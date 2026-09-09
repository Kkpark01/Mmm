/**
 * 変換元: Mcm1004uScreen.vb - Screen_Load / UpdateButton_Click / TabTsuikaButton_Click / LockReleaseButton_Click
 * 需要家契約一覧 Controller
 */
package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.List;

import jakarta.servlet.http.HttpSession;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.form.Mcm1004uForm;
import com.daifuku.mcm.service.Mcm1004uService;

import lombok.RequiredArgsConstructor;

@Controller
@RequestMapping("/mcm1004u")
@RequiredArgsConstructor
public class Mcm1004uController  extends BaseController{

    private final Mcm1004uService service;

    /**
     * 【変換元】Mcm1004uScreen_Load + Search + View
     *   画面初期表示 - 見積+契約時間帯読込、期間タブ構築、編集可否判定
     */
    @GetMapping
    public String index(@RequestParam("tmKeiyakujikanId") BigDecimal tmKeiyakujikanId,
                        @RequestParam(value = "seniMotoKbn", defaultValue = "1") int seniMotoKbn,
                        @RequestParam(value = "lockRelease", defaultValue = "false") boolean lockRelease,
                        Model model, HttpSession session) {
        Mcm1004uForm form = service.loadScreen(tmKeiyakujikanId, seniMotoKbn, lockRelease);
        model.addAttribute("form", form);
        return "mcm1004u/index";
    }

    /**
     * 【変換元】UpdateButton_Click
     *   登録ボタン - 変更チェック、期間チェック、状態更新、タブ内テーブル更新
     *   元コード: McmDBUtility.CleanProcedure("SP_TM")
     *             Me.checkChangedStatus() / Me.checkKikan()
     *             MyBase.DisplaySaveConfirm() -> JS confirm()
     *             Me.BeginTransaction() -> @Transactional
     */
    @PostMapping("/update")
    public String update(@ModelAttribute("form") Mcm1004uForm form,
                         RedirectAttributes ra, HttpSession session) {
        // 変換元: checkChangedStatus() - 変更有無チェック
        // TODO: フォーム変更検知ロジック

        // 変換元: checkKikan() - 期間バリデーション
        List<String> errors = service.validatePeriods(form.getTabs());
        if (!errors.isEmpty()) {
            ra.addFlashAttribute("errors", errors);
            return "redirect:/mcm1004u?tmKeiyakujikanId=" + form.getTmKeiyakujikanId();
        }

        // 変換元: DisplaySaveConfirm() -> JavaScript confirm() で実装
        String loginUser = (String) session.getAttribute("loginUser");

        service.updateAll(form, loginUser);
        ra.addFlashAttribute("message", "登録を完了しました。");
        return "redirect:/mcm1004u?tmKeiyakujikanId=" + form.getTmKeiyakujikanId();
    }

    /**
     * 【変換元】TabTsuikaButton_Click
     *   タブ追加 - 最終タブの開始日/終了日必須チェック、データコピー
     */
    @PostMapping("/addTab")
    public String addTab(@ModelAttribute("form") Mcm1004uForm form,
                         RedirectAttributes ra) {
        List<String> errors = service.validateForTabAdd(form);
        if (!errors.isEmpty()) {
            ra.addFlashAttribute("errors", errors);
        } else {
            service.addNewTab(form);
        }
        ra.addFlashAttribute("form", form);
        return "redirect:/mcm1004u?tmKeiyakujikanId=" + form.getTmKeiyakujikanId();
    }

    /**
     * 【変換元】LockReleaseButton_Click
     *   ロック解除 - lockReleaseFlg=true で再検索
     */
    @PostMapping("/lockRelease")
    public String lockRelease(@RequestParam("tmKeiyakujikanId") BigDecimal id,
                              RedirectAttributes ra) {
        return "redirect:/mcm1004u?tmKeiyakujikanId=" + id + "&lockRelease=true";
    }

	@Override
	protected String getScreenTitle() {
		// TODO 自動生成されたメソッド・スタブ
		return null;
	}

	@Override
	protected String getFunctionId() {
		// TODO 自動生成されたメソッド・スタブ
		return null;
	}
}
