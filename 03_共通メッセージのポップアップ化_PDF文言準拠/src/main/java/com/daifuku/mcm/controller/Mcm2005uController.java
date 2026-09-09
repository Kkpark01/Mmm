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
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.form.Mcm2005uForm;
import com.daifuku.mcm.service.Mcm2005uService;

/**
 * 【変換元】Mcm2005uScreen.vb / Mcm2005uTabControl.vb
 * MCM2005U 店舗見積ブランド詳細設定コントローラ
 *
 * エントリポイント:
 *   MCM2003U ブランドリンク     → GET /mcm2005u?umKihonBrandId=XXX
 *   MCM2003U 期間リンク        → GET /mcm2005u?umMitsumoriId=XXX
 *   （両方指定も可能）
 *
 * 遷移:
 *   保存後      → redirect:/mcm2005u (PRG)
 *   戻るボタン  → redirect:/mcm2003u
 */
@Controller
@RequestMapping("/mcm2005u")
public class Mcm2005uController extends BaseController {

    @Autowired
    private Mcm2005uService service;

    private static final String SESSION_FORM      = "MCM2005U_FORM";
    private static final String SESSION_BRAND_ID  = "mcm2005u.umKihonBrandId";
    private static final String SESSION_MITSUMORI_ID = "mcm2005u.umMitsumoriId";

    // ===================================================================
    // 初期表示
    // 【変換元】Search() / View() / Delivery.UmMitsumoriID, UmKihonBrandID
    //
    // umKihonBrandId のみ → umMitsumoriId をDB補完
    // umMitsumoriId のみ  → umKihonBrandId をDB補完
    // 両方 null           → セッションの既存フォームを再表示
    // ===================================================================

    @GetMapping
    public String index(
            @RequestParam(required = false) BigDecimal umKihonBrandId,
            @RequestParam(required = false) BigDecimal umMitsumoriId,
            Model model, HttpSession session) {

        boolean hasParam = (umKihonBrandId != null || umMitsumoriId != null);

        if (hasParam) {
            // パラメータがあれば常にリロード
            session.setAttribute(SESSION_BRAND_ID, umKihonBrandId);
            session.setAttribute(SESSION_MITSUMORI_ID, umMitsumoriId);
            session.removeAttribute(SESSION_FORM);
        } else {
            // パラメータなし（PRGリダイレクト後など）はセッションから復元
            umKihonBrandId = (BigDecimal) session.getAttribute(SESSION_BRAND_ID);
            umMitsumoriId  = (BigDecimal) session.getAttribute(SESSION_MITSUMORI_ID);
        }

        Mcm2005uForm form = (Mcm2005uForm) session.getAttribute(SESSION_FORM);
        if (form == null) {
            if (umKihonBrandId != null || umMitsumoriId != null) {
                try {
                    form = service.load(umKihonBrandId, umMitsumoriId);
                    // ロード完了後、補完されたIDをセッションに保存
                    session.setAttribute(SESSION_BRAND_ID, form.getUmKihonBrandId());
                    session.setAttribute(SESSION_MITSUMORI_ID, form.getUmMitsumoriId());
                } catch (Exception e) {
                    logger.error("MCM2005U: ロードエラー [umKihonBrandId={}, umMitsumoriId={}]",
                            umKihonBrandId, umMitsumoriId, e);
                    model.addAttribute("error", "データの読み込みに失敗しました: " + e.getMessage());
                    form = new Mcm2005uForm();
                }
            } else {
                form = new Mcm2005uForm();
            }
            session.setAttribute(SESSION_FORM, form);
        }

        setCommonAttributes(model, session);
        model.addAttribute("form", form);
        model.addAttribute("readOnly",
            service.isReadOnly(form.getJotai(), form.getSyouninJotai()));
        return "mcm2005u/index";
    }

    // ===================================================================
    // 保存（更新）
    // 【変換元】UpdateButtonTabNaiyo()
    //   MCM_UM_BRAND (MERGE INTO) + MCM_UM_TANKA (UPDATE)
    // ===================================================================

    @PostMapping("/save")
    public String save(@ModelAttribute Mcm2005uForm form,
                       HttpSession session, RedirectAttributes ra) {
        // セッションに保存済みの ID とブランド基本情報を補完
        Mcm2005uForm saved = (Mcm2005uForm) session.getAttribute(SESSION_FORM);
        if (saved != null) {
            form.setUmKihonBrandId(saved.getUmKihonBrandId());
            form.setUmMitsumoriId(saved.getUmMitsumoriId());
            form.setUmKihonMitsumoriId(saved.getUmKihonMitsumoriId());
            // HOSHU_KIN 計算に必要な KIHON_BRAND 側の値を引き継ぎ
            form.setHardhosyuKinKihon(saved.getHardhosyuKinKihon());
            form.setHoseisofthosyuKin(saved.getHoseisofthosyuKin());
            // HARDHOSYU_KIN は編集不可のため保存済み値を使用
            form.setHardhosyuKin(saved.getHardhosyuKin());
        }
        try {
            service.save(form, getLoginUserId());
            session.removeAttribute(SESSION_FORM);
            ra.addFlashAttribute("message", "登録を完了しました。");
        } catch (Exception e) {
            logger.error("MCM2005U: 保存エラー", e);
            ra.addFlashAttribute("error", "保存中にエラーが発生しました: " + e.getMessage());
        }
        return "redirect:/mcm2005u";
    }

    // ===================================================================
    // MCM2003U へ戻る
    // 【変換元】戻るボタン
    // MCM2003U はセッション "mcm2003u.umKihonMitsumoriId" を保持しているので
    // パラメータなしでリダイレクトすれば MCM2003U が自動リロードする。
    // ===================================================================

    @GetMapping("/back")
    public String back() {
        return "redirect:/mcm2003u";
    }

    @Override
    protected String getScreenTitle() { return "店舗見積ブランド詳細設定"; }

    @Override
    protected String getFunctionId() { return "MCM2005U"; }
}
