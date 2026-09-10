/**
 * 【変換元】Mcm0011uScreen.vb
 * 【説明】納入先マスタ（単票編集画面）のコントローラ
 * 元イベント対応:
 *   MCM0011uScreen_Load    → index (GET)
 *   UpdateButton_Click     → save (POST)
 *   AddPlantButton_Click   → plantAdd (POST)
 * 画面構成（Designer.vb）:
 *   GroupBox「納入先」→ 6項目
 *   GroupBox「連絡先」→ 6項目
 *   備考（Multiline）、DTS連携チェック、削除チェック
 *   GroupBox「更新履歴」→ 4項目（ReadOnly）
 *   ボタン: 登録、プラント追加
 */
package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.Collections;
import java.util.Map;
import java.util.Optional;

import jakarta.validation.Valid;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.entity.NonyusakiEntity;
import com.daifuku.mcm.entity.TantoEntity;
import com.daifuku.mcm.form.Mcm0011uForm;
import com.daifuku.mcm.repository.TantoRepository;
import com.daifuku.mcm.service.Mcm0011uService;

@Controller
@RequestMapping("/mcm0011u")
public class Mcm0011uController extends BaseController {

    @Autowired
    private Mcm0011uService service;

    @Autowired
    private TantoRepository tantoRepository;

    /** 初期表示・入力エラー・登録後の再表示で同じ更新権限を使用する。 */
    @ModelAttribute("isUpdateAuthority")
    public boolean updateAuthority() {
        return hasUpdateAuthority();
    }

    /**
     * 初期表示
     * 元VB: MCM0011uScreen_Load
     * nonyusakiId パラメータなし → 新規モード
     * nonyusakiId パラメータあり → 編集モード
     */
    @GetMapping
    public String index(
            @RequestParam(required = false) BigDecimal nonyusakiId,
            @RequestParam(required = false) String returnUrl,
            Model model) {
    	
    	model.addAttribute("dataNotFound", false);

        if (nonyusakiId == null || nonyusakiId.compareTo(BigDecimal.ZERO) == 0) {
            // ===== 新規モード =====
            // 元VB: nonyusakiId = 0 → 新規行追加, DTS_FLG=0, 削除チェック無効, プラント追加無効
            Mcm0011uForm form = service.initNewForm();
            model.addAttribute("form", form);
            model.addAttribute("entity", null);
        } else {
            // ===== 編集モード =====
            // 元VB: MCM_MA_NONYUSAKITableAdapter.Fill(... , nonyusakiId)
            Optional<NonyusakiEntity> opt = service.findByNonyusakiId(nonyusakiId);
            if (opt.isPresent()) {
                NonyusakiEntity entity = opt.get();
                Mcm0011uForm form = service.entityToForm(entity);
                model.addAttribute("form", form);
                model.addAttribute("entity", entity);
            } else {
                // データ0件 → MSG_0003 + ボタン無効化
                // 元VB: DisplayMessage(MSG_0003), 登録ボタン・プラント追加ボタン無効化
                Mcm0011uForm form = service.initNewForm();
                form.setNonyusakiId(nonyusakiId);
                form.setNewMode(false);
                model.addAttribute("form", form);
                model.addAttribute("entity", null);
                model.addAttribute("errorMessage", "指定された納入先データが見つかりません。");
                model.addAttribute("dataNotFound", true);
            }
        }

        model.addAttribute("returnUrl", returnUrl);
        return "mcm0011u/index";

    }

    /**
     * 納入先削除チェックON時、「登録」ボタン押下タイミングでプラント紐付きの有無を確認するためのAPI。
     * 画面側は、この結果に応じて確認ダイアログを表示するかエラーメッセージを表示するかを判定する。
     * サーバー側 save() でも同様のチェックを行っており、これはUX向上のための事前チェック（多重防御）。
     */
    @GetMapping("/checkPlantExists")
    @ResponseBody
    public Map<String, Boolean> checkPlantExists(@RequestParam(required = false) BigDecimal nonyusakiId) {
        if (nonyusakiId == null || nonyusakiId.compareTo(BigDecimal.ZERO) == 0) {
            return Collections.singletonMap("hasPlant", false);
        }
        boolean hasPlant = service.countPlantByNonyusakiId(nonyusakiId) > 0;
        return Collections.singletonMap("hasPlant", hasPlant);
    }

    /**
     * 登録/更新/削除処理
     * 元VB: UpdateButton_Click
     * → CheckExistUpdated → DisplaySaveConfirm → UpdateProcess → EnterUpdate → MSG_0008
     */

    @PostMapping("/save")
    public String save(
            @Valid @ModelAttribute("form") Mcm0011uForm form,
            BindingResult bindingResult,
            Model model,
            RedirectAttributes ra,
            @RequestParam(required = false) String returnUrl) {

        requireUpdateAuthority();

        if (bindingResult.hasErrors()) {
            reloadEntity(form, model);
            model.addAttribute("dataNotFound", false);
            model.addAttribute("returnUrl", returnUrl);
            return "mcm0011u/index";
        }

        String userId = getCurrentUserId();

        // ===== 削除チェックON → 削除処理へ分岐 =====
        if (form.isNonyusakiSakujo()) {
            boolean deleted = service.delete(form.getNonyusakiId());
            if (!deleted) {
                // プラントマスタで使用中のため削除不可（MSG_0006相当）
                reloadEntity(form, model);
                model.addAttribute("errorMessage", "既にプラント情報が登録されている為、削除する事が出来ません。");
                model.addAttribute("dataNotFound", false);
                model.addAttribute("returnUrl", returnUrl);
                return "mcm0011u/index";
            }
            // 削除成功 → 新規モードへ初期化（元VB: データクリア→新規モードへ）
            Mcm0011uForm newForm = service.initNewForm();
            model.addAttribute("form", newForm);
            model.addAttribute("entity", null);
            model.addAttribute("successMessage", "登録が完了しました。");
            model.addAttribute("dataNotFound", false);
            model.addAttribute("returnUrl", returnUrl);
            return "mcm0011u/index";
        }

        // ===== 通常の登録/更新処理 =====
        BigDecimal savedId = service.save(form, userId);
        form.setNonyusakiId(savedId);
        form.setNewMode(false);

        reloadEntity(form, model);
        model.addAttribute("successMessage", "登録が完了しました。");
        model.addAttribute("dataNotFound", false);
        model.addAttribute("returnUrl", returnUrl);

        return "mcm0011u/index";
    }


    /**
     * プラント追加ボタン
     * 元VB: AddPlantButton_Click
     * → If Me.IsChangedStatus Then UpdateProcess() End If
     * → ForwardScreen(MCM0012U, Mcm0012uDelivery(NonyusakiId, NonyusakiCd, NonyusakiNk))
     */
	@PostMapping("/plantAdd")
	public String plantAdd(
	        @Valid @ModelAttribute("form") Mcm0011uForm form,
	        BindingResult bindingResult,
	        Model model,
	        RedirectAttributes ra,
	        @RequestParam(required = false) String returnUrl) {
	
	    requireUpdateAuthority();
	
	    if (!bindingResult.hasErrors()) {
	        String userId = getCurrentUserId();
	        BigDecimal savedId = service.save(form, userId);
	        form.setNonyusakiId(savedId);
	    }
	
	    String backTo = "/mcm0011u?nonyusakiId=" + form.getNonyusakiId();
	    
		return "redirect:/mcm0012u?nonyusakiId=" + form.getNonyusakiId()
		        + "&nonyusakiCd=" + encode(form.getNonyusakiCd())
		        + "&nonyusakiNk=" + encode(form.getNonyusakiNk())
		        + "&returnUrl=" + encode(backTo);

	}

    // ===== ヘルパー =====
    private void reloadEntity(Mcm0011uForm form, Model model) {
        if (form.getNonyusakiId() != null
            && form.getNonyusakiId().compareTo(BigDecimal.ZERO) != 0) {
            Optional<NonyusakiEntity> opt = service.findByNonyusakiId(form.getNonyusakiId());
            model.addAttribute("entity", opt.orElse(null));
        } else {
            model.addAttribute("entity", null);
        }
    }

    /**
     * ログインユーザーに紐づく担当者マスタの「担当者名」を取得する。
     * 【変換元】CPSettingInfo.GetUserInfo 相当
     * 登録者・更新者（CREATED_BY / LASTUPDATE_BY）に設定する値として使用する。
     */
    private String getCurrentUserId() {
        String loginId = getLoginUserId();
        if (loginId == null) {
            return "SYSTEM";
        }
        return tantoRepository.findByLoginId(loginId)
                .map(TantoEntity::getTantoNk)
                .orElse(loginId);
    }

    private String encode(String s) {
        try {
            return java.net.URLEncoder.encode(s != null ? s : "", "UTF-8");
        } catch (Exception e) {
            return s != null ? s : "";
        }
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
