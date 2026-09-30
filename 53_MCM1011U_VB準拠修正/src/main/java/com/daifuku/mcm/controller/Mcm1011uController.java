package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.List;
import java.util.stream.Collectors;

import jakarta.servlet.http.HttpSession;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.InitBinder;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.dto.UserInfo;
import com.daifuku.mcm.form.Mcm1011uForm;
import com.daifuku.mcm.form.Mcm1011uForm.KeiyakuRowForm;
import com.daifuku.mcm.service.Mcm1011uService;

/**
 * 【変換元】Mcm1011uScreen.vb
 * MCM1011U 審査・承認（取引先契約）コントローラ
 *
 * 元イベント対応:
 *   KensakuButton_Click    → search (POST /mcm1011u/search)
 *   ShoninButton_Click     → approve (POST /mcm1011u/approve)
 *   SashimodoshiButton_Click → reject (POST /mcm1011u/reject)
 */
@Controller
@RequestMapping("/mcm1011u")
public class Mcm1011uController extends BaseController {

    @Autowired
    private Mcm1011uService service;

    private static final String SESSION_FORM = "MCM1011U_FORM";
    private static final String SESSION_OWNER = "MCM1011U_OWNER";
    private static final String SESSION_LIST_VERSION = "MCM1011U_LIST_VERSION";

    /**
     * 【移植】元VB: CPCoreUserControl.GetAuthorityDivision() 相当。
     *   「取引先契約関連」（機能ID=MCM1011U）権限が「作成」でなければ承認・差戻しを許可しない。
     */
    private boolean canUpdate() {
        return hasUpdateAuthority() && service.canUpdate(getLoginUserId());
    }

    @Override
    @InitBinder
    public void initBinder(org.springframework.web.bind.WebDataBinder binder) {
        super.initBinder(binder);
        // Search results and approval amounts must never be bound from browser input.
        binder.setAllowedFields("nonyusakiCd", "nonyusakiNk", "supportId", "plantNk", "shinsachu", "shoninchu", "shoninzumi", "sashimodoshi");
    }

    @GetMapping("/link/mcm1005u")
    public String openContract(@RequestParam BigDecimal tkKeiyakuId, HttpSession session) {
        if (!hasInspectionAuthority() || !service.canReadAttachment(getLoginUserId())) throw new com.daifuku.mcm.exception.AuthorityException("権限がないため実行できません。");
        Mcm1011uForm displayed = (Mcm1011uForm) session.getAttribute(SESSION_FORM);
        UserInfo viewer = getLoginUserInfo(session);
        if (tkKeiyakuId == null || displayed == null || viewer == null
                || !java.util.Objects.equals(session.getAttribute(SESSION_OWNER), getLoginUserId())
                || displayed.getKeiyakuRows() == null
                || displayed.getKeiyakuRows().stream().noneMatch(row -> hasContractLink(row, tkKeiyakuId))
                || service.search(displayed, viewer).stream().noneMatch(row -> hasContractLink(row, tkKeiyakuId)))
            throw new com.daifuku.mcm.exception.AuthorityException("参照できない契約です。再検索してください。");
        session.setAttribute("mcm1005u.keiyakuId", tkKeiyakuId);
        session.setAttribute("mcm1005u.seniMotoKbn", 5);
        session.setAttribute("mcm1005u.returnTo", "/mcm1011u");
        session.setAttribute("MCM1005U_RETURN_" + tkKeiyakuId, "/mcm1011u");
        return "redirect:/mcm1005u";
    }

    /** Attachment link: validate both the visible list and the currently queryable period before delegation. */
    @GetMapping("/attachment/download")
    public String downloadAttachment(@RequestParam BigDecimal tkKeiyakuId,
                                     @RequestParam BigDecimal tkTenpuId, HttpSession session) {
        if (!hasInspectionAuthority() || getLoginUserInfo(session) == null
                || !service.canReadAttachment(getLoginUserId()))
            throw new com.daifuku.mcm.exception.AuthorityException("権限がないため実行できません。");
        Mcm1011uForm saved = (Mcm1011uForm) session.getAttribute(SESSION_FORM);
        if (saved == null || saved.getKeiyakuRows() == null
                || !java.util.Objects.equals(session.getAttribute(SESSION_OWNER), getLoginUserId())
                || tkKeiyakuId == null || tkTenpuId == null || tkKeiyakuId.signum() <= 0 || tkTenpuId.signum() <= 0)
            throw new com.daifuku.mcm.exception.AuthorityException("参照できない添付ファイルです。");
        boolean displayed = saved.getKeiyakuRows().stream().anyMatch(row ->
                row.getTkKeiyakuId() != null && row.getTkKeiyakuId().compareTo(tkKeiyakuId) == 0
                && row.getTenpuList() != null && row.getTenpuList().stream().anyMatch(file ->
                        file.getTkTenpuId() != null && file.getTkTenpuId().compareTo(tkTenpuId) == 0));
        boolean current = displayed && service.search(saved, getLoginUserInfo(session)).stream().anyMatch(row ->
                row.getTkKeiyakuId() != null && row.getTkKeiyakuId().compareTo(tkKeiyakuId) == 0
                && row.getTenpuList() != null && row.getTenpuList().stream().anyMatch(file ->
                        file.getTkTenpuId() != null && file.getTkTenpuId().compareTo(tkTenpuId) == 0));
        if (!current) throw new com.daifuku.mcm.exception.AuthorityException("参照できない添付ファイルです。");
        // MCM1005U also validates authorization and resolves the stored path in its attachment service.
        return "redirect:/mcm1005u/attachment/download?tkKeiyakuId=" + tkKeiyakuId.toPlainString()
                + "&tkTenpuId=" + tkTenpuId.toPlainString();
    }

    // ===================================================================
    // 初期表示
    // ===================================================================

    @GetMapping
    public String index(Model model, HttpSession session) {
        if (!hasInspectionAuthority() || getLoginUserInfo(session) == null)
            throw new com.daifuku.mcm.exception.AuthorityException("ログインし直してください。");
        if (!java.util.Objects.equals(session.getAttribute(SESSION_OWNER), getLoginUserId())) {
            session.removeAttribute(SESSION_FORM);
            session.removeAttribute(SESSION_OWNER);
            session.removeAttribute(SESSION_LIST_VERSION);
        }
        Mcm1011uForm form = (Mcm1011uForm) session.getAttribute(SESSION_FORM);
        model.addAttribute("resetGridState", form == null);
        if (form == null) {
            form = new Mcm1011uForm();
            UserInfo user = getLoginUserInfo(session);
            form.setShinsachu(user != null && user.isShinsa());
            form.setShoninchu(user != null && user.isShonin());
            if (form.isShinsachu() || form.isShoninchu()) {
                try { form.setKeiyakuRows(service.search(form, user)); }
                catch (Exception ex) { model.addAttribute("error", "検索に失敗しました。検索ボタンで再実行してください。"); }
            }
            session.setAttribute(SESSION_FORM, form);
            session.setAttribute(SESSION_OWNER, getLoginUserId());
        } else {
            // 権限・担当区分が変更された後も古い検索結果を表示しない。
            try { form.setKeiyakuRows(service.search(form, getLoginUserInfo(session))); }
            catch (Exception ex) {
                form.setKeiyakuRows(java.util.List.of());
                model.addAttribute("error", "検索に失敗しました。検索ボタンで再実行してください。");
            }
        }
        // Every freshly rendered list has a distinct snapshot token.
        session.setAttribute(SESSION_LIST_VERSION, java.util.UUID.randomUUID().toString());
        model.addAttribute("listVersion", session.getAttribute(SESSION_LIST_VERSION));
        addKingakuLimits(model, session);
        setCommonAttributes(model, session);
        model.addAttribute("form", form);
        UserInfo viewer = getLoginUserInfo(session);
        model.addAttribute("viewerShinsa", viewer != null && viewer.isShinsa());
        model.addAttribute("viewerShonin", viewer != null && viewer.isShonin());
        model.addAttribute("canDownloadAttachment", service.canReadAttachment(getLoginUserId()));
        // 【移植】元VB: SYONINButton/SASHIMODOSHIButton.AuthorityIsThrough=False
        model.addAttribute("canUpdate", canUpdate());

        return "mcm1011u/index";
    }

    // ===================================================================
    // 検索
    // 【変換元】KensakuButton_Click → Search()
    // ===================================================================

    @PostMapping("/search")
    public String search(@ModelAttribute Mcm1011uForm form,
                         HttpSession session, RedirectAttributes ra) {
        if (!hasInspectionAuthority() || getLoginUserInfo(session) == null)
            throw new com.daifuku.mcm.exception.AuthorityException("ログインし直してください。");
        // A new search invalidates an action form rendered in another tab.
        session.setAttribute(SESSION_LIST_VERSION, java.util.UUID.randomUUID().toString());
        try {
            List<KeiyakuRowForm> rows = service.search(form, getLoginUserInfo(session));
            form.setKeiyakuRows(rows);
            session.setAttribute(SESSION_FORM, form);
            session.setAttribute(SESSION_OWNER, getLoginUserId());
            // VB契約版のMSG_0002E呼出しはコメントアウト。0件時は空一覧のみ表示。
        } catch (Exception e) {
            form.setKeiyakuRows(java.util.List.of());
            ra.addFlashAttribute("error", "検索中にエラーが発生しました。");
            session.setAttribute(SESSION_FORM, form);
            session.setAttribute(SESSION_OWNER, getLoginUserId());
        }
        return "redirect:/mcm1011u";
    }

    // ===================================================================
    // 承認処理
    // 【変換元】ShoninButton_Click
    //   チェックされた行の承認状態を1段階進める
    //   審査中(1)→承認中(2)、承認中(2)→承認済(3)
    // ===================================================================

    @PostMapping("/approve")
    public String approve(
            @RequestParam(value = "checkedIds", required = false) List<String> checkedIds,
            @RequestParam(value = "listVersion", required = false) String listVersion,
            HttpSession session, RedirectAttributes ra) {

        // 【移植】元VB: SYONINButton.AuthorityIsThrough=False（参照専用権限で非活性）。
        //   画面側のボタン非活性だけに依存せず、サーバー側でも実効権限を検証する。
        if (!canUpdate()) {
            ra.addFlashAttribute("error", "権限がないため実行できません。");
            return "redirect:/mcm1011u";
        }

        if (listVersion == null || !listVersion.equals(session.getAttribute(SESSION_LIST_VERSION))) {
            ra.addFlashAttribute("error", "別のタブで一覧が更新されました。再検索してから操作してください。");
            return "redirect:/mcm1011u";
        }
        Mcm1011uForm savedForm = (Mcm1011uForm) session.getAttribute(SESSION_FORM);
        if (savedForm == null || !java.util.Objects.equals(session.getAttribute(SESSION_OWNER), getLoginUserId())
                || checkedIds == null || checkedIds.isEmpty()) {
            ra.addFlashAttribute("error", "審査・承認可能な申請が選定されていません。");
            return "redirect:/mcm1011u";
        }

        if (checkedIds.stream().anyMatch(java.util.Objects::isNull)
                || checkedIds.stream().distinct().count() != checkedIds.size()
                || savedForm.getKeiyakuRows() == null) {
            ra.addFlashAttribute("error", "選択した申請が無効です。再検索してください。");
            return "redirect:/mcm1011u";
        }
        List<KeiyakuRowForm> checkedRows = savedForm.getKeiyakuRows().stream()
            .filter(r -> checkedIds.contains(r.getRowKey()))
            .collect(Collectors.toList());
        if (checkedRows.size() != checkedIds.size()) {
            ra.addFlashAttribute("error", "選択した申請が更新されたか、一覧外です。再検索してください。");
            return "redirect:/mcm1011u";
        }

        if (checkedRows.isEmpty()) {
            ra.addFlashAttribute("error", "審査・承認可能な申請が選定されていません。");
            return "redirect:/mcm1011u";
        }

        List<String> warnings;
        try {
            UserInfo userInfo = getLoginUserInfo(session);
            warnings = service.approve(checkedRows, userInfo, getLoginUserId());
        } catch (Exception e) {
            ra.addFlashAttribute("error", e instanceof com.daifuku.mcm.exception.McmBusinessException
                    ? e.getMessage() : "処理に失敗しました。再検索して状態を確認してください。");
            return "redirect:/mcm1011u";
        }
        if (!warnings.isEmpty()) ra.addFlashAttribute("infoMessage", String.join("\n", warnings));
        session.setAttribute(SESSION_LIST_VERSION, java.util.UUID.randomUUID().toString());
        // コミット後の再検索障害を承認失敗と表示しない（二重承認を防ぐ）。
        try {
            savedForm.setKeiyakuRows(service.search(savedForm, getLoginUserInfo(session)));
            session.setAttribute(SESSION_FORM, savedForm);
            ra.addFlashAttribute("message", "審査・承認を完了しました。");
        } catch (Exception e) {
            savedForm.setKeiyakuRows(java.util.List.of());
            session.setAttribute(SESSION_FORM, savedForm);
            ra.addFlashAttribute("message", "承認は完了しましたが、一覧を再取得できません。再検索してください。");
        }
        return "redirect:/mcm1011u";
    }

    // ===================================================================
    // 差戻し処理
    // 【変換元】SashimodoshiButton_Click
    // ===================================================================

    @PostMapping("/reject")
    public String reject(
            @RequestParam(value = "checkedIds", required = false) List<String> checkedIds,
            @RequestParam(value = "listVersion", required = false) String listVersion,
            HttpSession session, RedirectAttributes ra) {

        // 【移植】元VB: SASHIMODOSHIButton.AuthorityIsThrough=False（参照専用権限で非活性）。
        if (!canUpdate()) {
            ra.addFlashAttribute("error", "権限がないため実行できません。");
            return "redirect:/mcm1011u";
        }

        if (listVersion == null || !listVersion.equals(session.getAttribute(SESSION_LIST_VERSION))) {
            ra.addFlashAttribute("error", "別のタブで一覧が更新されました。再検索してから操作してください。");
            return "redirect:/mcm1011u";
        }
        Mcm1011uForm savedForm = (Mcm1011uForm) session.getAttribute(SESSION_FORM);
        if (savedForm == null || !java.util.Objects.equals(session.getAttribute(SESSION_OWNER), getLoginUserId())
                || checkedIds == null || checkedIds.isEmpty()) {
            ra.addFlashAttribute("error", "差戻し可能な申請が選択されていません。");
            return "redirect:/mcm1011u";
        }

        if (checkedIds.stream().anyMatch(java.util.Objects::isNull)
                || checkedIds.stream().distinct().count() != checkedIds.size()
                || savedForm.getKeiyakuRows() == null) {
            ra.addFlashAttribute("error", "選択した申請が無効です。再検索してください。");
            return "redirect:/mcm1011u";
        }
        List<KeiyakuRowForm> checkedRows = savedForm.getKeiyakuRows().stream()
            .filter(r -> checkedIds.contains(r.getRowKey()))
            .collect(Collectors.toList());
        if (checkedRows.size() != checkedIds.size()) {
            ra.addFlashAttribute("error", "選択した申請が更新されたか、一覧外です。再検索してください。");
            return "redirect:/mcm1011u";
        }

        if (checkedRows.isEmpty()) {
            ra.addFlashAttribute("error", "差戻し可能な申請が選択されていません。");
            return "redirect:/mcm1011u";
        }

        try {
            service.reject(checkedRows, getLoginUserInfo(session), getLoginUserId());
        } catch (Exception e) {
            ra.addFlashAttribute("error", e instanceof com.daifuku.mcm.exception.McmBusinessException
                    ? e.getMessage() : "処理に失敗しました。再検索して状態を確認してください。");
            return "redirect:/mcm1011u";
        }
        session.setAttribute(SESSION_LIST_VERSION, java.util.UUID.randomUUID().toString());
        try {
            savedForm.setKeiyakuRows(service.search(savedForm, getLoginUserInfo(session)));
            session.setAttribute(SESSION_FORM, savedForm);
            ra.addFlashAttribute("message", "差戻しを完了しました。");
        } catch (Exception e) {
            savedForm.setKeiyakuRows(java.util.List.of());
            session.setAttribute(SESSION_FORM, savedForm);
            ra.addFlashAttribute("message", "差戻しは完了しましたが、一覧を再取得できません。再検索してください。");
        }
        return "redirect:/mcm1011u";
    }

    // ===================================================================
    // private
    // ===================================================================

    private void addKingakuLimits(Model model, HttpSession session) {
        UserInfo userInfo = getLoginUserInfo(session);
        model.addAttribute("keiyakuShinsaKin",
            (userInfo != null) ? userInfo.getKeiyakuShinsaKin() : null);
        model.addAttribute("keiyakuShoninKin",
            (userInfo != null) ? userInfo.getKeiyakuShoninKin() : null);
    }

    /** Validate the formal VB main-contract reference link. */
    private boolean hasContractLink(KeiyakuRowForm row, BigDecimal contractId) {
        return row != null && row.getTkKeiyakuId() != null && row.getTkKeiyakuId().compareTo(contractId) == 0;
    }

    @Override
    protected String getScreenTitle() { return "審査・承認（取引先契約）"; }

    @Override
    protected String getFunctionId() { return "MCM1011U"; }
}