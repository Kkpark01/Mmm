package com.daifuku.mcm.controller;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.constants.Mcm1009uConstants;
import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.form.Mcm1009uForm;
import com.daifuku.mcm.service.Mcm1009uService;
import jakarta.servlet.http.HttpSession;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.math.BigDecimal;
import java.util.Collections;

/**
 * 【変換元】Mcm1009uScreen.vb
 *   MCM1009U パック契約・解約 Controller
 *
 * 【権限制御について】
 *   旧VB版 Mcm1009uScreen_Load() は
 *     GetAuthorityDivision("MCM1005U") = AUTHORITY_DIVISION_UPDATE
 *   の場合のみ「契約作成」「契約破棄」列を表示していた。
 *   Web版フレームワークの UserInfo は画面単位の権限区分マップを保持していないため、
 *   現時点では共通の更新権限（BaseController#hasUpdateAuthority / ROLE_UPDATE）を
 *   等価物として使用する。画面単位権限（MCM1005U基準）の完全再現は未確認事項。
 *   いずれの場合も「表示制御だけで終わらせず」サーバー側でも必ず検証する。
 */
@Controller
@RequestMapping("/mcm1009u")
public class Mcm1009uController extends BaseController {

    @Autowired
    private Mcm1009uService service;

    @Override
    protected String getScreenTitle() { return Mcm1009uConstants.SCREEN_TITLE; }

    @Override
    protected String getFunctionId() { return Mcm1009uConstants.SCREEN_ID; }

    /** 初期表示 【変換元】Mcm1009uScreen_Load() */
    @GetMapping
    public String index(Model model, HttpSession session) {
        model.addAttribute("form", new Mcm1009uForm());
        model.addAttribute("rows", Collections.emptyList());
        setCommonAttributes(model, session);
        model.addAttribute("hasUpdateAuth", hasUpdateAuthority());
        return "mcm1009u/index";
    }

    /** 検索処理 【変換元】SEARCHButton_Click() */
    @PostMapping("/search")
    public String search(@ModelAttribute("form") Mcm1009uForm form, Model model, HttpSession session) {
        session.setAttribute(Mcm1009uConstants.SESSION_KEY_SEARCH_FORM, form);
        Mcm1009uService.SearchResult result = service.search(form);
        return render(form, result, model, session);
    }

    /**
     * 契約破棄 【変換元】CellContentClick - Case KEIYAKU_HAKI
     *   POST + CSRF（Spring Security 既定で有効）。処理後は同一検索条件で再検索する。
     */
    @PostMapping("/cancel-contract")
    public String cancelContract(@RequestParam BigDecimal tkKeiyakuId,
                                 @RequestParam BigDecimal tmKeiyakujikanId,
                                 HttpSession session, RedirectAttributes ra) {
        // 画面のリンク非表示だけに依存せず、サーバー側で更新権限を必ず検証する
        requireUpdateAuthority();

        try {
            String errorMsg = service.cancelContract(tkKeiyakuId, tmKeiyakujikanId, getLoginUserId());
            if (errorMsg != null) {
                // 業務エラー（支払発生済み等）
                ra.addFlashAttribute("errors", Collections.singletonList(errorMsg));
            } else {
                ra.addFlashAttribute("message", Mcm1009uConstants.MSG_CANCEL_COMPLETE);
            }
        } catch (McmBusinessException e) {
            // 業務エラー（対象なし／二重破棄／改ざん）。SQLやスタックトレースは画面に出さない。
            logger.warn("MCM1009U 契約破棄 業務エラー: {}", e.getMessage());
            ra.addFlashAttribute("errors", Collections.singletonList(e.getMessage()));
        } catch (org.springframework.dao.DataAccessException e) {
            // システムエラー
            logger.error("MCM1009U 契約破棄 DBエラー", e);
            ra.addFlashAttribute("errors", Collections.singletonList(
                    "処理を完了できませんでした。再検索してから、もう一度操作してください。"));
        }
        return "redirect:/mcm1009u/re-search";
    }

    /** 再検索（PRGパターン）。完了・エラーメッセージはリダイレクト後に表示される。 */
    @GetMapping("/re-search")
    public String reSearch(Model model, HttpSession session) {
        Mcm1009uForm form = (Mcm1009uForm) session.getAttribute(
                Mcm1009uConstants.SESSION_KEY_SEARCH_FORM);
        if (form == null) return "redirect:/mcm1009u";
        Mcm1009uService.SearchResult result = service.search(form);
        return render(form, result, model, session);
    }

    /** 一覧表示の共通処理（検索条件・権限・メッセージを常に同じ形で設定する） */
    private String render(Mcm1009uForm form, Mcm1009uService.SearchResult result,
                          Model model, HttpSession session) {
        model.addAttribute("form", form);
        model.addAttribute("rows", result.getRows());
        setCommonAttributes(model, session);
        model.addAttribute("hasUpdateAuth", hasUpdateAuthority());
        if (result.hasErrors() && !model.containsAttribute("errors")) {
            model.addAttribute("errors", result.getErrors());
        }
        return "mcm1009u/index";
    }

    /** MCM1004U遷移 【変換元】Case TM_IRAI_NO（参照遷移のため権限は参照権限で足りる） */
    @GetMapping("/forward/mcm1004u")
    public String forwardMcm1004u(@RequestParam BigDecimal tmKeiyakujikanId, RedirectAttributes ra) {
        // 1004UはURLパラメータで受け取る。VB Delivery.TmKeiyakujikanIDと遷移区分1を引き継ぐ。
        ra.addAttribute("tmKeiyakujikanId", tmKeiyakujikanId);
        ra.addAttribute("seniMotoKbn", 1); // VB Mcm1004uConstant.SENIMOTO_MITSUMORI_LINK
        // 1004Uの既存の戻り先を維持し、前回表示したFormを流用せず選択対象を読み直す。
        ra.addAttribute("returnTo", "/mcm1003u");
        return "redirect:/mcm1004u";
    }

    /** MCM1005U遷移 【変換元】Case KEIYAKU_NO */
    @GetMapping("/forward/mcm1005u")
    public String forwardMcm1005u(@RequestParam BigDecimal tkKeiyakuId,
                                  @RequestParam BigDecimal tmKeiyakujikanId,
                                  @RequestParam BigDecimal torihikisakiId,
                                  @RequestParam String torihikisakiCd,
                                  @RequestParam String torihikisakiNk,
                                  @RequestParam String nonyusakiCd,
                                  @RequestParam String nonyusakiNk,
                                  @RequestParam String supportId,
                                  @RequestParam String plantNk,
                                  @RequestParam BigDecimal plantId,
                                  HttpSession session) {
        clearForwardAttributes(session, "mcm1005u.");
        session.setAttribute("mcm1005u.keiyakuId", tkKeiyakuId.intValue());
        session.setAttribute("mcm1005u.tmKeiyakujikanId", new String[]{tmKeiyakujikanId.toString()});
        session.setAttribute("mcm1005u.torihikisakiId", torihikisakiId);
        session.setAttribute("mcm1005u.torihikisakiCd", torihikisakiCd);
        session.setAttribute("mcm1005u.torihikisakiNk", torihikisakiNk);
        session.setAttribute("mcm1005u.nonyusakiCd", nonyusakiCd);
        session.setAttribute("mcm1005u.nonyusakiNk", nonyusakiNk);
        session.setAttribute("mcm1005u.supportId", supportId);
        session.setAttribute("mcm1005u.plantNk", plantNk);
        session.setAttribute("mcm1005u.plantId", plantId.intValue());
        session.setAttribute("mcm1005u.seniMotoKbn", Mcm1009uConstants.SENIMOTO_KEIYAKU_LINK);
        session.setAttribute("mcm1005u.returnTo", "/mcm1009u");
        return "redirect:/mcm1005u";
    }

    /**
     * MCM1006U1遷移 【変換元】Case KEIYAKU_KEIYAKU
     *   旧版では更新権限がある場合のみ「契約作成」列を表示していたため、
     *   Web版でもサーバー側で更新権限を検証する（URL直接アクセス対策）。
     */
    @GetMapping("/forward/mcm1006u1")
    public String forwardMcm1006u1(@RequestParam BigDecimal tmKeiyakujikanId,
                                   @RequestParam BigDecimal torihikisakiId,
                                   @RequestParam String torihikisakiCd,
                                   @RequestParam String torihikisakiNk,
                                   @RequestParam String nonyusakiCd,
                                   @RequestParam String nonyusakiNk,
                                   @RequestParam String supportId,
                                   @RequestParam String plantNk,
                                   @RequestParam BigDecimal plantId,
                                   HttpSession session) {
        requireUpdateAuthority();

        var delivery = new com.daifuku.mcm.form.Mcm1006uForm();
        delivery.setSeniMotoKbn(1);
        delivery.setTmKeiyakujikanIds(java.util.List.of(tmKeiyakujikanId));
        delivery.setTorihikisakiId(torihikisakiId);
        delivery.setTorihikisakiCd(torihikisakiCd);
        delivery.setTorihikisakiNk(torihikisakiNk);
        delivery.setNonyusakiCd(nonyusakiCd);
        delivery.setNonyusakiNk(nonyusakiNk);
        delivery.setSupportId(supportId);
        delivery.setPlantNk(plantNk);
        delivery.setPlantId(plantId);
        session.removeAttribute("MCM1006U_FORM");
        session.removeAttribute("MCM1006U_RESULT");
        session.removeAttribute("MCM1006U_RETURN_ID");
        session.setAttribute("MCM1006U_RETURN_URL", "/mcm1009u");
        session.setAttribute("MCM1006U_DELIVERY", delivery);
        return "redirect:/mcm1006u/step1";
    }

    /**
     * 遷移先の引き継ぎ属性を一旦全削除する。
     * 前回操作した契約ID・契約時間IDがセッションに残り、
     * 別契約を開いた際に誤って使用されることを防ぐ。
     */
    private void clearForwardAttributes(HttpSession session, String prefix) {
        for (String name : java.util.Collections.list(session.getAttributeNames())) {
            if (name.startsWith(prefix)) {
                session.removeAttribute(name);
            }
        }
    }
}
