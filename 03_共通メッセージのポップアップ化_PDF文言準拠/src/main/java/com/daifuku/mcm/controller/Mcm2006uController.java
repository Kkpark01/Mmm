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
import com.daifuku.mcm.constants.Mcm2006uConstants;
import com.daifuku.mcm.form.Mcm2006uForm;
import com.daifuku.mcm.form.Mcm2006uForm.KikanTabForm;
import com.daifuku.mcm.form.Mcm2007uForm;
import com.daifuku.mcm.service.Mcm2006uService;

/**
 * 【変換元】Mcm2006uScreen.vb / Mcm2006uTabControl.vb
 * MCM2006U ユーザ契約内容コントローラ
 *
 * エントリポイント:
 *   GET /mcm2006u?ukKeiyakuId=XXX      → 既存契約参照 (SENIMOTO_KBN_KEYAKU_LINK=0)
 *   GET /mcm2006u?seniMotoKbn=1        → 新規契約作成 (SENIMOTO_INSERT)
 *   GET /mcm2006u                      → セッション復元 (MCM2007U戻り後 / PRGリダイレクト後)
 *   GET /mcm2006u?kikanId=YYY          → タブ切り替え
 *
 * MCM2007U連携:
 *   GET /mcm2006u/change               → MCM2007Uへ（選択タブIDをセッションに保存）
 *   セッション "mcm2006u.pendingKikan" → MCM2007U適用後にタブ追加
 */
@Controller
@RequestMapping("/mcm2006u")
public class Mcm2006uController extends BaseController {

    @Autowired
    private Mcm2006uService service;

    private static final String SESSION_FORM       = "MCM2006U_FORM";
    private static final String SESSION_KEIYAKU_ID = "mcm2006u.ukKeiyakuId";
    private static final String SESSION_SENIMOTO   = "mcm2006u.seniMotoKbn";
    private static final String SESSION_PENDING    = "mcm2006u.pendingKikan"; // MCM2007U→MCM2006U受け渡し

    // ===================================================================
    // 初期表示
    // ===================================================================

    @GetMapping
    public String index(
            @RequestParam(required = false) BigDecimal ukKeiyakuId,
            @RequestParam(required = false) Integer seniMotoKbn,
            @RequestParam(required = false) BigDecimal kikanId,
            Model model, HttpSession session) {

        // セッションからMCM2007U適用データを取り込む
        KikanTabForm pendingKikan = (KikanTabForm) session.getAttribute(SESSION_PENDING);
        session.removeAttribute(SESSION_PENDING);

        Mcm2006uForm form;

        if (ukKeiyakuId != null) {
            // 既存契約を参照
            session.setAttribute(SESSION_KEIYAKU_ID, ukKeiyakuId);
            session.setAttribute(SESSION_SENIMOTO,
                seniMotoKbn != null ? seniMotoKbn : Mcm2006uConstants.SENIMOTO_KBN_KEYAKU_LINK);
            session.removeAttribute(SESSION_FORM);
            try {
                form = service.load(ukKeiyakuId);
                form.setSeniMotoKbn(
                    seniMotoKbn != null ? seniMotoKbn : Mcm2006uConstants.SENIMOTO_KBN_KEYAKU_LINK);
            } catch (Exception e) {
                logger.error("MCM2006U: ロードエラー ukKeiyakuId={}", ukKeiyakuId, e);
                model.addAttribute("error", "データの読み込みに失敗しました: " + e.getMessage());
                form = new Mcm2006uForm();
            }
        } else if (seniMotoKbn != null && seniMotoKbn == Mcm2006uConstants.SENIMOTO_INSERT) {
            // 新規契約作成（MCM2007U経由）
            session.setAttribute(SESSION_SENIMOTO, Mcm2006uConstants.SENIMOTO_INSERT);
            session.removeAttribute(SESSION_KEIYAKU_ID);
            session.removeAttribute(SESSION_FORM);
            form = (Mcm2006uForm) session.getAttribute(SESSION_FORM);
            if (form == null) {
                form = new Mcm2006uForm();
                form.setSeniMotoKbn(Mcm2006uConstants.SENIMOTO_INSERT);
            }
        } else {
            // セッション復元（PRG後 / MCM2007U戻り後）
            form = (Mcm2006uForm) session.getAttribute(SESSION_FORM);
            if (form == null) {
                BigDecimal savedId = (BigDecimal) session.getAttribute(SESSION_KEIYAKU_ID);
                if (savedId != null) {
                    try {
                        form = service.load(savedId);
                        Integer savedSeni = (Integer) session.getAttribute(SESSION_SENIMOTO);
                        if (savedSeni != null) form.setSeniMotoKbn(savedSeni);
                    } catch (Exception e) {
                        logger.error("MCM2006U: セッション復元エラー", e);
                        form = new Mcm2006uForm();
                    }
                } else {
                    form = new Mcm2006uForm();
                }
            }
        }

        // MCM2007Uから戻った場合: 新しい期間タブを追加
        if (pendingKikan != null) {
            service.applyPendingKikan(form, pendingKikan);
        }

        // タブ切り替え
        if (kikanId != null) {
            form.setSelectedKikanId(kikanId);
        }

        session.setAttribute(SESSION_FORM, form);

        setCommonAttributes(model, session);
        model.addAttribute("form", form);
        model.addAttribute("readOnly",
            service.isReadOnly(form.getJotai(), form.getSeniMotoKbn()));
        model.addAttribute("selectedTab", form.getSelectedTab());
        return "mcm2006u/index";
    }

    // ===================================================================
    // 保存（更新）
    // 【変換元】UpdateButton_Click → UpdateButtonTabNaiyo
    // ===================================================================

    @PostMapping("/save")
    public String save(@ModelAttribute Mcm2006uForm form,
                       HttpSession session, RedirectAttributes ra) {
        // セッションのフォームから非HTML送信フィールド（タブ内グリッド等）を補完
        Mcm2006uForm saved = (Mcm2006uForm) session.getAttribute(SESSION_FORM);
        if (saved != null) {
            // HTML送信でないフィールドをセッションから補完
            form.setUkKeiyakuId(saved.getUkKeiyakuId());
            form.setNonyusakiId(saved.getNonyusakiId());
            form.setNonyusakiCd(saved.getNonyusakiCd());
            form.setNonyusakiNk(saved.getNonyusakiNk());
            form.setNonyusakijusyo1Nk(saved.getNonyusakijusyo1Nk());
            form.setNonyusakijusyo2Nk(saved.getNonyusakijusyo2Nk());
            form.setPlantId(saved.getPlantId());
            form.setSupportId(saved.getSupportId());
            form.setPlantNk(saved.getPlantNk());
            form.setSeniMotoKbn(saved.getSeniMotoKbn());
            // 期間タブ一覧もセッションから引き継ぎ
            if (form.getKikanTabs() == null || form.getKikanTabs().isEmpty()) {
                form.setKikanTabs(saved.getKikanTabs());
            }
            form.setSelectedKikanId(saved.getSelectedKikanId());
        }

        try {
            service.save(form, getLoginUserId());
            session.removeAttribute(SESSION_FORM);
            ra.addFlashAttribute("message", "登録を完了しました。");
            // PRG後は保存済みのukKeiyakuIdでリダイレクト
            session.setAttribute(SESSION_KEIYAKU_ID, form.getUkKeiyakuId());
        } catch (Exception e) {
            logger.error("MCM2006U: 保存エラー", e);
            ra.addFlashAttribute("error", "保存中にエラーが発生しました: " + e.getMessage());
        }
        return "redirect:/mcm2006u";
    }

    // ===================================================================
    // MCM2007Uへ（変更ボタン）
    // 【変換元】TabHenkoButtonClick — Delivery生成 → MCM2007U起動
    //
    // セッションにMCM2006Uのコンテキストを保存してMCM2007Uへリダイレクト
    // ===================================================================

    @GetMapping("/change")
    public String change(@RequestParam(required = false) BigDecimal kikanId,
                         HttpSession session) {
        Mcm2006uForm form = (Mcm2006uForm) session.getAttribute(SESSION_FORM);
        if (form != null && kikanId != null) {
            form.setSelectedKikanId(kikanId);
            session.setAttribute(SESSION_FORM, form);
        }
        // MCM2007Uにセッション経由でコンテキストを渡す
        // mcm2007u.* セッションキーはMCM2007UControllerで設定
        return "redirect:/mcm2007u?seniMotoKbn=" + Mcm2006uConstants.SENIMOTO_TAB_HENKO;
    }

    // ===================================================================
    // 戻る（呼び出し元へ）
    // ===================================================================

    @GetMapping("/back")
    public String back() {
        return "redirect:/mcm1001l";
    }

    @Override
    protected String getScreenTitle() { return "ユーザ契約内容"; }

    @Override
    protected String getFunctionId() { return "MCM2006U"; }
}
