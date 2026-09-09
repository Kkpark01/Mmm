package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.List;

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
import com.daifuku.mcm.form.Mcm2006uForm.KikanTabForm;
import com.daifuku.mcm.form.Mcm2007uForm;
import com.daifuku.mcm.service.Mcm2006uService;
import com.daifuku.mcm.service.Mcm2007uService;

/**
 * 【変換元】Mcm2007uScreen1.vb / Mcm2007uScreen2.vb
 * MCM2007U ユーザ契約内容変更コントローラ（2ステップウィザード）
 *
 * Step1: GET /mcm2007u         → 店舗見積+ブランド選択
 * Step2: POST /mcm2007u/step2  → 機器構成・明細・個体選択（Step1から遷移）
 * 適用:  POST /mcm2007u/apply  → MCM2006Uへ戻る
 * 戻る:  GET  /mcm2007u/back   → MCM2006Uへ戻る（キャンセル）
 *
 * DB書き込みなし（SELECT専用画面）。
 */
@Controller
@RequestMapping("/mcm2007u")
public class Mcm2007uController extends BaseController {

    @org.springframework.web.bind.annotation.ExceptionHandler(org.springframework.dao.DataAccessException.class)
    public String dataAccessFailure(org.springframework.dao.DataAccessException e, RedirectAttributes ra) {
        logger.error("MCM2007U data access failed",e);
        ra.addFlashAttribute("errors",List.of("画面の情報を取得できませんでした。再検索してから、もう一度操作してください。"));
        return "redirect:/mcm2004u";
    }

    @Autowired
    private Mcm2007uService service;

    @Autowired
    private Mcm2006uService mcm2006uService;

    private static final String SESSION_FORM      = "MCM2007U_FORM";
    private static final String SESSION_STEP1_FORM = "MCM2007U_STEP1_FORM";

    // MCM2006Uセッションキー（MCM2006Uが設定したコンテキストを読む）
    private static final String MCM2006U_SESSION_FORM    = "MCM2006U_FORM";
    private static final String MCM2006U_PENDING_KIKAN   = "mcm2006u.pendingKikan";

    // ===================================================================
    // Step1: 初期表示（店舗見積+ブランド選択）
    // 【変換元】Mcm2007uScreen1.vb — Form_Load
    // ===================================================================

    @GetMapping
    public String step1(
            @RequestParam(required = false) Integer seniMotoKbn,
            Model model, HttpSession session) {

        // MCM2006Uのコンテキスト（plantId等）をセッションから取得
        com.daifuku.mcm.form.Mcm2006uForm mcm2006uForm =
            (com.daifuku.mcm.form.Mcm2006uForm) session.getAttribute(MCM2006U_SESSION_FORM);

        Mcm2007uForm form = new Mcm2007uForm();

        if (mcm2006uForm != null) {
            form.setUkKeiyakuId(mcm2006uForm.getUkKeiyakuId());
            form.setPlantId(mcm2006uForm.getPlantId());
            form.setNonyusakiId(mcm2006uForm.getNonyusakiId());
            form.setNonyusakiCd(mcm2006uForm.getNonyusakiCd());
            form.setNonyusakiNk(mcm2006uForm.getNonyusakiNk());
            form.setNonyusakijusyo1Nk(mcm2006uForm.getNonyusakijusyo1Nk());
            form.setNonyusakijusyo2Nk(mcm2006uForm.getNonyusakijusyo2Nk());
            form.setSupportId(mcm2006uForm.getSupportId());
            form.setPlantNk(mcm2006uForm.getPlantNk());

            // 変更ボタン時: 選択中タブの期間情報を初期値に設定
            KikanTabForm selectedTab = mcm2006uForm.getSelectedTab();
            if (selectedTab != null) {
                form.setUkKikanId(selectedTab.getUkKikanId());
                form.setKaisiDt(selectedTab.getSyuryoDt()); // 前期の翌日が開始日候補
            }
        }

        form.setSeniMotoKbn(seniMotoKbn != null ? seniMotoKbn : Mcm2006uConstants.SENIMOTO_TAB_HENKO);

        // 店舗見積一覧ロード
        if (form.getPlantId() != null) {
            service.loadStep1(form);
        }

        // 2004U「契約作成」: VB同様、当期(なければ最も近い未来)の見積とブランドを選択する。
        if (form.getSeniMotoKbn() == Mcm2006uConstants.SENIMOTO_INSERT
                && session.getAttribute("mcm2007u.from2004") instanceof BigDecimal estimateId) {
            service.selectInitialEstimate(form, estimateId);
            session.removeAttribute("mcm2007u.from2004");
        }

        session.setAttribute(SESSION_FORM, form);
        setCommonAttributes(model, session);
        model.addAttribute("form", form);
        return "mcm2007u/index";
    }

    // ===================================================================
    // Step1→Step2: バリデーション通過後にStep2へ
    // 【変換元】NextButton_Click
    // ===================================================================

    @PostMapping("/step2")
    public String step2(@ModelAttribute Mcm2007uForm form,
                        HttpSession session, Model model) {
        // Step1のフォームデータとセッションのコンテキストをマージ
        Mcm2007uForm saved = (Mcm2007uForm) session.getAttribute(SESSION_FORM);
        if (saved != null) {
            form.setPlantId(saved.getPlantId());
            form.setNonyusakiId(saved.getNonyusakiId());
            form.setNonyusakiCd(saved.getNonyusakiCd());
            form.setNonyusakiNk(saved.getNonyusakiNk());
            form.setNonyusakijusyo1Nk(saved.getNonyusakijusyo1Nk());
            form.setNonyusakijusyo2Nk(saved.getNonyusakijusyo2Nk());
            form.setSupportId(saved.getSupportId());
            form.setPlantNk(saved.getPlantNk());
            form.setUkKeiyakuId(saved.getUkKeiyakuId());
            form.setUkKikanId(saved.getUkKikanId());
            form.setSeniMotoKbn(saved.getSeniMotoKbn());
        }

        // バリデーション
        List<String> errors = service.validateStep1(form);
        if (!errors.isEmpty()) {
            // Step1に戻る
            service.loadStep1(form);
            session.setAttribute(SESSION_FORM, form);
            setCommonAttributes(model, session);
            model.addAttribute("form", form);
            model.addAttribute("errors", errors);
            return "mcm2007u/index";
        }

        // Step1選択データを保存
        session.setAttribute(SESSION_STEP1_FORM, form);

        // Step2データをロード
        service.loadStep2(form);
        session.setAttribute(SESSION_FORM, form);

        setCommonAttributes(model, session);
        model.addAttribute("form", form);
        return "mcm2007u/step2";
    }

    // ===================================================================
    // Step2 表示（バック後に再表示する場合）
    // ===================================================================

    @GetMapping("/step2")
    public String step2Get(HttpSession session, Model model) {
        Mcm2007uForm form = (Mcm2007uForm) session.getAttribute(SESSION_FORM);
        if (form == null) {
            return "redirect:/mcm2007u";
        }
        if (form.getKoseiRows().isEmpty()) {
            service.loadStep2(form);
            session.setAttribute(SESSION_FORM, form);
        }
        setCommonAttributes(model, session);
        model.addAttribute("form", form);
        return "mcm2007u/step2";
    }

    // ===================================================================
    // 適用ボタン: Step2選択内容をMCM2006Uへ渡す
    // 【変換元】ApplyButton_Click → CloseAndSendDataToBaseForm(Delivery)
    // ===================================================================

    @PostMapping("/apply")
    public String apply(@ModelAttribute Mcm2007uForm form,
                        HttpSession session, RedirectAttributes ra) {
        // セッションのコンテキストを復元
        Mcm2007uForm saved = (Mcm2007uForm) session.getAttribute(SESSION_FORM);
        if (saved != null) {
            form.setPlantId(saved.getPlantId());
            form.setNonyusakiId(saved.getNonyusakiId());
            form.setNonyusakiCd(saved.getNonyusakiCd());
            form.setNonyusakiNk(saved.getNonyusakiNk());
            form.setNonyusakijusyo1Nk(saved.getNonyusakijusyo1Nk());
            form.setNonyusakijusyo2Nk(saved.getNonyusakijusyo2Nk());
            form.setSupportId(saved.getSupportId());
            form.setPlantNk(saved.getPlantNk());
            form.setUkKeiyakuId(saved.getUkKeiyakuId());
            form.setUkKikanId(saved.getUkKikanId());
            form.setSeniMotoKbn(saved.getSeniMotoKbn());
            form.setKaisiDt(saved.getKaisiDt());
            form.setSyuryoDt(saved.getSyuryoDt());
            // Step1選択データを復元
            form.setMitsumoriRows(saved.getMitsumoriRows());
        }

        // バリデーション（重複個体チェック）
        List<String> errors = service.validateStep2(form);
        if (!errors.isEmpty()) {
            session.setAttribute(SESSION_FORM, form);
            setCommonAttributes(ra, session);
            ra.addFlashAttribute("errors", errors);
            return "redirect:/mcm2007u/step2";
        }

        // MCM2006UへKikanTabFormとして渡す
        KikanTabForm pendingKikan = mcm2006uService.buildKikanFromDelivery(form);
        session.setAttribute(MCM2006U_PENDING_KIKAN, pendingKikan);

        session.removeAttribute("mcm2007u.from2004Cancel");
        // セッションクリア
        session.removeAttribute(SESSION_FORM);
        session.removeAttribute(SESSION_STEP1_FORM);

        return "redirect:/mcm2006u";
    }

    // ===================================================================
    // 戻る（MCM2006Uへキャンセル）
    // ===================================================================

    @GetMapping("/back")
    public String back(HttpSession session) {
        boolean fromSearch = Boolean.TRUE.equals(session.getAttribute("mcm2007u.from2004Cancel"));
        session.removeAttribute("mcm2007u.from2004Cancel");
        session.removeAttribute("mcm2007u.from2004");
        session.removeAttribute(SESSION_FORM);
        session.removeAttribute(SESSION_STEP1_FORM);
        if(fromSearch) {
            session.removeAttribute(MCM2006U_SESSION_FORM);
            session.removeAttribute(MCM2006U_PENDING_KIKAN);
            return "redirect:/mcm2004u";
        }
        return "redirect:/mcm2006u";
    }

    // RedirectAttributesへ共通属性をセット（フラッシュ属性経由のエラー表示用）
    private void setCommonAttributes(RedirectAttributes ra, HttpSession session) {
        // 必要に応じて共通属性を設定（現在は空実装）
    }

    @Override
    protected String getScreenTitle() { return "ユーザ契約内容変更"; }

    @Override
    protected String getFunctionId() { return "MCM2007U"; }
}
