package com.daifuku.mcm.controller;

import java.math.BigDecimal;

import jakarta.servlet.http.HttpSession;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.form.Mcm2003uForm;
import com.daifuku.mcm.service.Mcm2003uService;

/**
 * 【変換元】Mcm2003uScreen.vb
 * MCM2003U 店舗見積内容基本設定コントローラ
 *
 * 元イベント対応:
 *   MitsumoriButton_Click → save  (POST /mcm2003u/save)
 *   HakkoButton_Click     → publish (POST /mcm2003u/publish)
 *   SinseiButton_Click    → sinsei (POST /mcm2003u/sinsei)
 *   KensakuGamenButton_Click → back (GET /mcm2003u/back)
 *
 * MCM2004U → MCM2003U の遷移:
 *   セッションキー "mcm2003u.umKihonMitsumoriId" で UM_KIHON_MITSUMORI_ID を受取る
 */
@Controller
@RequestMapping("/mcm2003u")
public class Mcm2003uController extends BaseController {

    @Autowired
    private Mcm2003uService service;

    private static final String SESSION_FORM = "MCM2003U_FORM";
    private static final String SESSION_UM_ID = "mcm2003u.umKihonMitsumoriId";
    private static final String SESSION_SENI_MOTO_KBN = "mcm2003u.seniMotoKbn";

    // ===================================================================
    // 初期表示
    // ===================================================================

    @GetMapping
    public String index(Model model, HttpSession session) {
        BigDecimal umKihonMitsumoriId = (BigDecimal) session.getAttribute(SESSION_UM_ID);

        Mcm2003uForm form = (Mcm2003uForm) session.getAttribute(SESSION_FORM);
        boolean reload = form == null
            || (umKihonMitsumoriId != null
                && !umKihonMitsumoriId.equals(form.getUmKihonMitsumoriId()));

        if (reload) {
            if (umKihonMitsumoriId != null) {
                try {
                    form = service.load(umKihonMitsumoriId);
                } catch (Exception e) {
                    logger.error("MCM2003U: ロードエラー [umKihonMitsumoriId={}]", umKihonMitsumoriId, e);
                    model.addAttribute("error", "データの読み込みに失敗しました: " + e.getMessage());
                    form = new Mcm2003uForm();
                }
            } else {
                form = new Mcm2003uForm();
            }
            Integer seniMotoKbn = (Integer) session.getAttribute(SESSION_SENI_MOTO_KBN);
            if (seniMotoKbn != null) {
                form.setSeniMotoKbn(seniMotoKbn);
                session.removeAttribute(SESSION_SENI_MOTO_KBN);
            }
            session.setAttribute(SESSION_FORM, form);
        }

        setCommonAttributes(model, session);
        model.addAttribute("form", form);
        model.addAttribute("readOnly",
            service.isReadOnly(form.getSyouninJotai(), form.getSeniMotoKbn()));
        return "mcm2003u/index";
    }

    // ===================================================================
    // 見積登録（保存）
    // 【変換元】MitsumoriButton_Click → UpdateButton(BUTTON_FLG_SAKUSEI)
    // ===================================================================

    @PostMapping("/save")
    public String save(@ModelAttribute Mcm2003uForm form,
                       HttpSession session, RedirectAttributes ra) {
        form.setUmKihonMitsumoriId(getStoredId(session));
        try {
            String savedJotai = getSavedJotai(session);
            form.setJotai(savedJotai);
            service.save(form, getLoginUserId());
            session.removeAttribute(SESSION_FORM);
            ra.addFlashAttribute("message", "登録を完了しました。");
        } catch (Exception e) {
            ra.addFlashAttribute("error", "保存中にエラーが発生しました: " + e.getMessage());
        }
        return "redirect:/mcm2003u";
    }

    // ===================================================================
    // 見積発行
    // 【変換元】HakkoButton_Click → UpdateButton(BUTTON_FLG_HAKKO)
    //   状態が作成中(9)の場合のみ見積(1)に変更。
    // ===================================================================

    @PostMapping("/publish")
    public String publish(@ModelAttribute Mcm2003uForm form,
                          HttpSession session, RedirectAttributes ra) {
        form.setUmKihonMitsumoriId(getStoredId(session));
        try {
            String savedJotai = getSavedJotai(session);
            form.setJotai(savedJotai);
            service.publish(form, getLoginUserId());
            session.removeAttribute(SESSION_FORM);
            ra.addFlashAttribute("message", "見積を発行しました。");
        } catch (Exception e) {
            ra.addFlashAttribute("error", "見積発行中にエラーが発生しました: " + e.getMessage());
        }
        return "redirect:/mcm2003u";
    }

    // ===================================================================
    // 申請
    // 【変換元】SinseiButton_Click
    // ===================================================================

    @PostMapping("/sinsei")
    public String sinsei(HttpSession session, RedirectAttributes ra) {
        BigDecimal umKihonMitsumoriId = getStoredId(session);
        try {
            service.sinsei(umKihonMitsumoriId, getLoginUserId());
            session.removeAttribute(SESSION_FORM);
            ra.addFlashAttribute("message", "申請を完了しました。");
        } catch (IllegalStateException e) {
            ra.addFlashAttribute("error", e.getMessage());
        } catch (Exception e) {
            ra.addFlashAttribute("error", "申請中にエラーが発生しました: " + e.getMessage());
        }
        return "redirect:/mcm2003u";
    }

    // ===================================================================
    // 検索画面へ戻る
    // 【変換元】KensakuGamenButton_Click
    // ===================================================================

    @GetMapping("/back")
    public String back() {
        return "redirect:/mcm2004u";
    }

    // ===================================================================
    // private
    // ===================================================================

    private BigDecimal getStoredId(HttpSession session) {
        return (BigDecimal) session.getAttribute(SESSION_UM_ID);
    }

    /** セッションに保存済みのフォームから JOTAI を取得（UPDATE時の現在状態保持用）。 */
    private String getSavedJotai(HttpSession session) {
        Mcm2003uForm savedForm = (Mcm2003uForm) session.getAttribute(SESSION_FORM);
        return (savedForm != null) ? savedForm.getJotai() : null;
    }

    @Override
    protected String getScreenTitle() { return "店舗見積内容基本設定"; }

    @Override
    protected String getFunctionId() { return "MCM2003U"; }
}
