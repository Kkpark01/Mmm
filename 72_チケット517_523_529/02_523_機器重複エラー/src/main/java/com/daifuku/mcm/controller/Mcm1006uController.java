package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

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
import com.daifuku.mcm.constants.Mcm1006uConstants;
import com.daifuku.mcm.form.Mcm1006uForm;
import com.daifuku.mcm.service.Mcm1006uService;

/**
 * 【変換元】Mcm1006u1Screen.vb / Mcm1006u2Screen.vb
 * MCM1006U 取引先契約機器選定 Controller
 *
 * 画面フロー:
 *   MCM1005U → GET /mcm1006u/step1   (1画面目初期表示)
 *           POST /mcm1006u/step1/next (次へ: バリデーション + step2へ)
 *           GET  /mcm1006u/step2      (2画面目初期表示)
 *           POST /mcm1006u/step2/back (前へ: step1に戻る)
 *           POST /mcm1006u/step2/complete (摘要へ: MCM1005Uに結果返却)
 *
 * 画面間データはセッション "MCM1006U_FORM" に格納。
 */
@Controller
@RequestMapping("/mcm1006u")
public class Mcm1006uController extends BaseController {
    @InitBinder
    public void protectSelectionVersion(org.springframework.web.bind.WebDataBinder binder) {
        binder.setDisallowedFields("selectionVersion", "restorePendingSelection", "warningMsg");
    }

    @Autowired
    private Mcm1006uService service;

    /**
     * 【移植】元VB: CPCoreUserControl.GetAuthorityDivision() 相当。
     *   MCM1006Uは常にMCM1005Uから起動されるため、機能ID=MCM1005Uの権限で判定する。
     */
    private boolean canUpdate() {
        return hasUpdateAuthority() && service.canUpdate(getLoginUserId());
    }

    // =========================================================================
    // 1画面目: 初期表示
    // 【変換元】Mcm1006u1Screen_Load()
    //   MCM1005Uからのパラメータ（deliveryデータ）をsessionに格納してから表示。
    // =========================================================================

    /**
     * 1画面目 初期表示
     * GET /mcm1006u/step1?seniMotoKbn=1&nonyusakiId=...
     *
     * MCM1005Uから受け取るパラメータ（すべてオプション）:
     *   seniMotoKbn, nonyusakiId, nonyusakiCd, nonyusakiNk,
     *   plantId, supportId, plantNk, kaisiDt, syuryoDt,
     *   torihikisakiId, torihikisakiCd, torihikisakiNk,
     *   keiyakuNo, tkKeiyakuId, tkKikanId, tabFlg,
     *   tmKeiyakujikanIds (comma-separated)
     */
    @GetMapping("/step1")
    public String step1(
            @RequestParam(required = false, defaultValue = "1") int seniMotoKbn,
            @RequestParam(required = false) BigDecimal nonyusakiId,
            @RequestParam(required = false) String nonyusakiCd,
            @RequestParam(required = false) String nonyusakiNk,
            @RequestParam(required = false) BigDecimal plantId,
            @RequestParam(required = false) String supportId,
            @RequestParam(required = false) String plantNk,
            @RequestParam(required = false) String kaisiDt,
            @RequestParam(required = false) String syuryoDt,
            @RequestParam(required = false) BigDecimal torihikisakiId,
            @RequestParam(required = false) String torihikisakiCd,
            @RequestParam(required = false) String torihikisakiNk,
            @RequestParam(required = false) String keiyakuNo,
            @RequestParam(required = false) BigDecimal tkKeiyakuId,
            @RequestParam(required = false) BigDecimal tkKikanId,
            @RequestParam(required = false, defaultValue = "0") int tabFlg,
            @RequestParam(required = false) String tmKeiyakujikanIdsStr,
            Model model,
            HttpSession session) {

        // MCM1005Uからの事前受け渡しがあればそちらを優先使用
        // （変換元: MCM1005U HenkoButton_Click → TabHenkoButtonClick → delivery）
        Mcm1006uForm delivery = (Mcm1006uForm) session.getAttribute("MCM1006U_DELIVERY");
        Mcm1006uForm form;
        if (delivery != null) {
            session.removeAttribute("MCM1006U_DELIVERY");
            form = delivery;
            // delivery の値で URL パラメータを上書き（nullの場合はdelivery値を使用）
            seniMotoKbn   = delivery.getSeniMotoKbn();
            plantId       = delivery.getPlantId();
            torihikisakiId = delivery.getTorihikisakiId();
        } else {
            // URL パラメータからフォームを構築
            form = new Mcm1006uForm();
            form.setSeniMotoKbn(seniMotoKbn);
            form.setNonyusakiId(nonyusakiId);
            form.setNonyusakiCd(nonyusakiCd);
            form.setNonyusakiNk(nonyusakiNk);
            form.setPlantId(plantId);
            form.setSupportId(supportId);
            form.setPlantNk(plantNk);
            form.setKaisiDt(kaisiDt);
            form.setSyuryoDt(syuryoDt);
            form.setTorihikisakiId(torihikisakiId);
            form.setTorihikisakiCd(torihikisakiCd);
            form.setTorihikisakiNk(torihikisakiNk);
            form.setKeiyakuNo(keiyakuNo);
            form.setTkKeiyakuId(tkKeiyakuId);
            form.setTkKikanId(tkKikanId);
            form.setTabFlg(tabFlg);
        }

        // TAB_HENKO時は既存のtmKeiyakujikanIdsを復元してcheckを立てる
        List<BigDecimal> selectedIds = delivery != null
                ? form.getTmKeiyakujikanIds()
                : parseIds(tmKeiyakujikanIdsStr);

        // 見積データロード
        if (plantId != null && torihikisakiId != null) {
            form.setMitsumoriRows(service.loadMitsumori(plantId, torihikisakiId, selectedIds));
        }

        session.setAttribute(Mcm1006uConstants.SESSION_KEY, form);
        model.addAttribute("form", form);
        // 【変換元】Mcm1006u1Screen 見積単価グリッド（参照専用）
        addStep1Tables(model, form);
        // TAB_HENKO時は開始日をreadonly
        model.addAttribute("kaisiReadonly", seniMotoKbn == Mcm1006uConstants.SENIMOTO_TAB_HENKO);
        // 【移植】元VB: nextButton.AuthorityIsThrough=False
        model.addAttribute("canUpdate", canUpdate());
        return "mcm1006u/step1/index";
    }

    /**
     * 1画面目 再表示（「前へ」ボタンでstep2から戻った場合）
     * Sessionからformを復元して表示。
     */
    @GetMapping("/step1/back")
    public String step1Back(Model model, HttpSession session) {
        Mcm1006uForm form = getFormFromSession(session);
        if (form == null) return "redirect:/mcm1006u/step1";
        model.addAttribute("form", form);
        addStep1Tables(model, form);
        model.addAttribute("kaisiReadonly",
            form.getSeniMotoKbn() == Mcm1006uConstants.SENIMOTO_TAB_HENKO);
        model.addAttribute("canUpdate", canUpdate());
        return "mcm1006u/step1/index";
    }

    // =========================================================================
    // 1画面目: 「次へ」ボタン
    // 【変換元】Mcm1006u1Screen.vb nextButton_Click()
    // =========================================================================

    /**
     * 「次へ」ボタン
     * POST /mcm1006u/step1/next
     * フォームのチェック状態を受け取り、バリデーション後にstep2へ遷移。
     */
    @PostMapping("/step1/next")
    public String step1Next(@ModelAttribute Mcm1006uForm submittedForm,
                            RedirectAttributes ra,
                            HttpSession session) {

        // 【移植】元VB: nextButton.AuthorityIsThrough=False（参照専用権限で非活性）。
        //   画面側のボタン非活性だけに依存せず、サーバー側でも実効権限を検証する。
        if (!canUpdate()) {
            ra.addFlashAttribute("errorMsg", "権限がないため実行できません。");
            return "redirect:/mcm1006u/step1/back";
        }

        // セッションのformにチェック状態を反映
        Mcm1006uForm form = getFormFromSession(session);
        if (form == null) {
            ra.addFlashAttribute("errorMsg", "画面の有効期限が切れました。取引先契約内容から開き直してください。");
            return "redirect:/mcm1006u/step1";
        } else {
            // 入力値（期間と見積チェック状態）を反映
            if (form.getSeniMotoKbn() != Mcm1006uConstants.SENIMOTO_TAB_HENKO) {
                form.setKaisiDt(submittedForm.getKaisiDt());
            }
            form.setSyuryoDt(submittedForm.getSyuryoDt());
            mergeMitsumoriCheckState(form, submittedForm);
        }

        // バリデーション
        String errorMsg = service.validateStep1(form);
        if (errorMsg != null) {
            ra.addFlashAttribute("errorMsg", errorMsg);
            ra.addFlashAttribute("form", form);
            session.setAttribute(Mcm1006uConstants.SESSION_KEY, form);
            return "redirect:/mcm1006u/step1/back";
        }

        // バリデーション通過: 2画面目データをロード
        service.loadStep2Data(form);
        session.setAttribute(Mcm1006uConstants.SESSION_KEY, form);

        return "redirect:/mcm1006u/step2";
    }

    // =========================================================================
    // 2画面目: 初期表示
    // 【変換元】Mcm1006u2Screen_Load()
    // =========================================================================

    /**
     * 2画面目 初期表示
     * GET /mcm1006u/step2
     */
    @GetMapping("/step2")
    public String step2(Model model, HttpSession session) {
        Mcm1006uForm form = getFormFromSession(session);
        if (form == null) return "redirect:/mcm1006u/step1";
        model.addAttribute("form", form);
        // 【移植】元VB: tekiyoButton.AuthorityIsThrough=False
        model.addAttribute("canUpdate", canUpdate());
        return "mcm1006u/step2/index";
    }

    // =========================================================================
    // 2画面目: 「前へ」ボタン
    // 【変換元】Mcm1006u2Screen.vb — ReDrawScreen(MCM1006U1)
    // =========================================================================

    /**
     * 「前へ」ボタン
     * POST /mcm1006u/step2/back
     * チェック状態をsessionに保存してstep1に戻る。
     */
    @PostMapping("/step2/back")
    public String step2Back(@ModelAttribute Mcm1006uForm submittedForm,
                            HttpSession session) {
        if (!canUpdate()) return "redirect:/mcm1006u/step2";
        Mcm1006uForm form = getFormFromSession(session);
        if (form != null) {
            mergeStep2CheckState(form, submittedForm);
            session.setAttribute(Mcm1006uConstants.SESSION_KEY, form);
        }
        return "redirect:/mcm1006u/step1/back";
    }

    // =========================================================================
    // 2画面目: 「摘要へ」ボタン
    // 【変換元】Mcm1006u2Screen.vb tekiyoButton_Click()
    // =========================================================================

    /**
     * 「摘要へ」ボタン
     * POST /mcm1006u/step2/complete
     * 月割り計算・数量補正を行い、MCM1005Uへリダイレクト。
     */
    @PostMapping("/step2/complete")
    public String step2Complete(@ModelAttribute Mcm1006uForm submittedForm,
                                RedirectAttributes ra,
                                HttpSession session) {

        // 【移植】元VB: tekiyoButton.AuthorityIsThrough=False（参照専用権限で非活性）。
        if (!canUpdate()) {
            ra.addFlashAttribute("errorMsg", "権限がないため実行できません。");
            return "redirect:/mcm1006u/step2";
        }

        Mcm1006uForm form = getFormFromSession(session);
        if (form == null) {
            ra.addFlashAttribute("errorMsg", "画面の有効期限が切れました。取引先契約内容から開き直してください。");
            return "redirect:/mcm1006u/step1";
        } else {
            mergeStep2CheckState(form, submittedForm);
        }

        String errorMsg = service.processComplete(form);
        if (errorMsg != null) {
            ra.addFlashAttribute("errorMsg", errorMsg);
            session.setAttribute(Mcm1006uConstants.SESSION_KEY, form);
            return "redirect:/mcm1006u/step2";
        }

        // 結果をセッションに格納してMCM1005Uへ返す
        session.setAttribute("MCM1006U_RESULT", form);
        // 新規契約を摘要した後も、1005の閉じるで元の検索画面へ戻る。
        Object caller = session.getAttribute("MCM1006U_RETURN_URL");
        if (form.getTkKeiyakuId() == null && caller instanceof String url
                && java.util.Set.of("/mcm1003u", "/mcm1009u", "/mcm1010u", "/mcm3007u").contains(url))
            session.setAttribute("MCM1005U_RETURN_0", url);
        // 摘要時は選定した期間タブへMCM1005U側で移動するため、閉じる用の戻り先は破棄する
        session.removeAttribute("MCM1006U_RETURN_ID");
        session.removeAttribute("MCM1006U_RETURN_URL");
        session.removeAttribute(Mcm1006uConstants.SESSION_KEY);

        // MCM1005Uに戻る（tkKeiyakuIdをパラメータで渡す）
        if (form.getTkKeiyakuId() != null) {
            return "redirect:/mcm1005u?tkKeiyakuId=" + form.getTkKeiyakuId()
                + "&mcm1006uCompleted=true";
        }
        return "redirect:/mcm1005u?mcm1006uCompleted=true";
    }

    // =========================================================================
    // 「閉じる」ボタン（1画面目・2画面目共通）
    // 【変換元】VB 閉じる — 選定結果を反映せず呼出元（MCM1005U）へ戻る
    // =========================================================================

    /**
     * 「閉じる」ボタン
     * GET /mcm1006u/close
     * 選定途中の状態を破棄し、結果（MCM1006U_RESULT）は設定せずに呼出元へ戻る。
     * ブラウザ履歴（history.back）に依存しないため、step1/step2間を往復した後でも確実に戻れる。
     */
    @GetMapping("/close")
    public String close(HttpSession session) {
        Mcm1006uForm form = getFormFromSession(session);
        Object returnId = session.getAttribute("MCM1006U_RETURN_ID");
        Object returnUrl = session.getAttribute("MCM1006U_RETURN_URL");
        session.removeAttribute(Mcm1006uConstants.SESSION_KEY);
        session.removeAttribute("MCM1006U_DELIVERY");
        session.removeAttribute("MCM1006U_RETURN_ID");
        session.removeAttribute("MCM1006U_RETURN_URL");
        if (form != null && form.getTkKeiyakuId() != null) {
            // 【VB準拠】「変更」を押した期間タブ（selectedKikanIndex）へ戻す。
            //   同一契約の戻り先が保持されている場合のみ使用し、別契約の古い値は使わない。
            if (returnId instanceof BigDecimal id && id.compareTo(form.getTkKeiyakuId()) == 0
                    && returnUrl instanceof String url && url.startsWith("/mcm1005u?")) {
                return "redirect:" + url;
            }
            // returnToは付けない（付けるとMCM1005U側で編集中の下書きが破棄されるため）
            return "redirect:/mcm1005u?tkKeiyakuId=" + form.getTkKeiyakuId();
        }
        if (returnUrl instanceof String url && java.util.Set.of("/mcm1003u", "/mcm1009u", "/mcm1010u", "/mcm3007u").contains(url)) {
            return "redirect:" + url;
        }
        // 呼出元がない場合の退避先
        return "redirect:/mcm1009u";
    }

    // =========================================================================
    // ユーティリティ
    // =========================================================================

    private Mcm1006uForm getFormFromSession(HttpSession session) {
        return (Mcm1006uForm) session.getAttribute(Mcm1006uConstants.SESSION_KEY);
    }

    /**
     * submitされたフォームの見積チェック状態をsessionのformに反映する。
     * フォームの mitsumoriRows はインデックス順で対応する。
     */
    private void mergeMitsumoriCheckState(Mcm1006uForm sessionForm, Mcm1006uForm submitted) {
        if (submitted.getMitsumoriRows() == null) return;
        List<Mcm1006uForm.MitsumoriRowForm> sessionRows = sessionForm.getMitsumoriRows();
        List<Mcm1006uForm.MitsumoriRowForm> subRows     = submitted.getMitsumoriRows();
        for (int i = 0; i < sessionRows.size() && i < subRows.size(); i++) {
            sessionRows.get(i).setCheckFlg(subRows.get(i).getCheckFlg());
        }
    }

    /**
     * submitされたフォームの2画面目チェック状態をsessionのformに反映する。
     */
    private void mergeStep2CheckState(Mcm1006uForm sessionForm, Mcm1006uForm submitted) {
        mergeCheckList(sessionForm.getKoseiRows(), submitted.getKoseiRows(),
                (s, sub) -> s.setCheckFlg(allowedCheck(s.getTmIraiNo(), sub.getCheckFlg())));
        mergeCheckList(sessionForm.getMeisaiRows(), submitted.getMeisaiRows(),
                (s, sub) -> s.setCheckFlg(allowedCheck(s.getTmIraiNo(), sub.getCheckFlg())));
        mergeCheckList(sessionForm.getKotaiRows(), submitted.getKotaiRows(),
                (s, sub) -> s.setCheckFlg(allowedCheck(s.getTmIraiNo(), sub.getCheckFlg())));
        // hidden単価フラグは信用せず、見積対象の選定済み明細から導出する。
        for (var detail : sessionForm.getMeisaiRows()) {
            boolean parentOn = sessionForm.getKoseiRows().stream().anyMatch(r -> r.getCheckFlg() == 1
                    && sameId(r.getKikikoseiId(), detail.getKikikoseiId()));
            if (!parentOn) detail.setCheckFlg(0);
        }
        for (var item : sessionForm.getKotaiRows()) {
            boolean parentOn = sessionForm.getMeisaiRows().stream().anyMatch(r -> r.getCheckFlg() == 1
                    && sameId(r.getKikikoseiId(), item.getKikikoseiId())
                    && sameId(r.getKikimeisaiId(), item.getKikimeisaiId()));
            if (!parentOn) item.setCheckFlg(0);
        }
        for (var price : sessionForm.getTankaRows()) {
            boolean selected = price.getTmTankaId() != null && price.getTmKikanId() != null
                    && sessionForm.getMeisaiRows().stream().anyMatch(r -> r.getCheckFlg() == 1
                    && sameId(r.getKikikoseiId(), price.getKikikoseiId())
                    && sameId(r.getKikimeisaiId(), price.getKikimeisaiId()));
            price.setCheckFlg(selected ? 1 : 0);
        }
    }

    private static int allowedCheck(String requestNo, int flag) {
        return requestNo != null && !requestNo.isBlank() && flag == 1 ? 1 : 0;
    }

    private static boolean sameId(BigDecimal a, BigDecimal b) {
        return a != null && b != null && a.compareTo(b) == 0;
    }

    private void addStep1Tables(Model model, Mcm1006uForm form) {
        model.addAttribute("step1KikanRows", service.loadStep1Kikan(form.getPlantId(), form.getTorihikisakiId()));
        model.addAttribute("step1TankaRows", service.loadStep1Tanka(form.getPlantId(), form.getTorihikisakiId()));
        model.addAttribute("step1Labels", service.estimateLabels(form.getMitsumoriRows()));
    }

    @FunctionalInterface
    private interface RowMerger<T> {
        void merge(T sessionRow, T submittedRow);
    }

    private <T> void mergeCheckList(List<T> sessionRows, List<T> subRows, RowMerger<T> merger) {
        if (subRows == null || sessionRows == null) return;
        for (int i = 0; i < sessionRows.size() && i < subRows.size(); i++) {
            merger.merge(sessionRows.get(i), subRows.get(i));
        }
    }

    private List<BigDecimal> parseIds(String idsStr) {
        List<BigDecimal> list = new ArrayList<>();
        if (idsStr == null || idsStr.isBlank()) return list;
        for (String s : idsStr.split(",")) {
            try { list.add(new BigDecimal(s.trim())); } catch (Exception ignore) {}
        }
        return list;
    }

    @Override
    protected String getScreenTitle() { return Mcm1006uConstants.SCREEN_NAME; }

    @Override
    protected String getFunctionId() { return Mcm1006uConstants.SCREEN_ID; }
}
