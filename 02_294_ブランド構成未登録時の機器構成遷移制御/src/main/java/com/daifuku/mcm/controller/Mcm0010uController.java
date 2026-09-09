/**
 * 【変換元】Mcm0010uScreen.vb
 * 【説明】納入先マスタ検索画面のコントローラ
 * 元イベント対応:
 *   Screen_Load        → index (GET)
 *   SEARCHButton_Click → search (POST)
 *   NONYUSAKIDataGridView_CellContentClick → navigateNonyusaki
 *   PLANTDataGridView_CellContentClick     → navigatePlant
 *   RowDeleteButton_Click                  → deletePlant (POST)
 */
package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

import jakarta.servlet.http.HttpSession;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.SessionAttribute;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.dto.Mcm0010uBrandKoseiDto;
import com.daifuku.mcm.dto.Mcm0010uKikiKoseiDto;
import com.daifuku.mcm.entity.NonyusakiEntity;
import com.daifuku.mcm.entity.PlantEntity;
import com.daifuku.mcm.form.Mcm0010uForm;
import com.daifuku.mcm.service.Mcm0010uService;

@Controller
@RequestMapping("/mcm0010u")
public class Mcm0010uController extends BaseController {

    public static final String MSG_0110 = "ブランド構成が設定されていない為、機器構成の設定画面には遷移することは出来ません。";

    @ModelAttribute("brandKoseiRequiredMessage")
    public String brandKoseiRequiredMessage() {
        return MSG_0110;
    }

    @Autowired
    private Mcm0010uService service;
    
    private static final String SESSION_MCM0010_SEARCH_FORM = "mcm0010u_search_form";

    /**
     * 初期表示
     * 元VB: Mcm0010uScreen_Load
     */

	@GetMapping
    public String index(Model model, HttpSession session) {
        Mcm0010uForm savedForm = (Mcm0010uForm) session.getAttribute(SESSION_MCM0010_SEARCH_FORM);

        // ★ MCM関連のセッションをクリア
         clearMcmSession(session);
        
        if (savedForm != null) {
            model.addAttribute("form", savedForm);

            List<NonyusakiEntity> nonyusakiList = service.searchNonyusaki(savedForm);
            List<PlantEntity> plantList = service.searchPlants(savedForm, getAuthorityDivision());
            List<Mcm0010uBrandKoseiDto> brandList = service.searchBrandKosei(savedForm);
            List<Mcm0010uKikiKoseiDto> kikiList = service.searchKikiKosei(savedForm);

            model.addAttribute("nonyusakiList", nonyusakiList);
            model.addAttribute("plantList", plantList);
            model.addAttribute("brandList", brandList);
            model.addAttribute("kikiList", kikiList);
            model.addAttribute("searched", true);
        } else {
            model.addAttribute("form", new Mcm0010uForm());
            model.addAttribute("nonyusakiList", Collections.emptyList());
            model.addAttribute("plantList", Collections.emptyList());
            model.addAttribute("brandList", Collections.emptyList());
            model.addAttribute("kikiList", Collections.emptyList());
            model.addAttribute("searched", false);
        }

        model.addAttribute("isMaintenance", isMaintenance());
        return "mcm0010u/index";
    }


    /**
     * 検索
     * 元VB: SEARCHButton_Click
     */

	@PostMapping("/search")
	public String search(@ModelAttribute Mcm0010uForm form, Model model, HttpSession session) {
	    session.setAttribute(SESSION_MCM0010_SEARCH_FORM, form);

	    model.addAttribute("form", form);
	    model.addAttribute("isMaintenance", isMaintenance());
	    model.addAttribute("searched", true);

	    if (service.isAllEmpty(form)) {
	        model.addAttribute("errorMessage", "検索条件は1項目以上選択して下さい。");
	        model.addAttribute("nonyusakiList", Collections.emptyList());
	        model.addAttribute("plantList", Collections.emptyList());
	        model.addAttribute("brandList", Collections.emptyList());
	        model.addAttribute("kikiList", Collections.emptyList());
	        return "mcm0010u/index";
	    }

	    // ★追加：納入先コード／サポートIDに全角文字が含まれる場合は
	    //         検索を実行せず「0件」相当のエラー表示とする
	    if (containsZenkaku(form.getNonyusakiCd()) || containsZenkaku(form.getSupportId())) {
	        model.addAttribute("errorMessage", "検索結果が1件も存在しません。");
	        model.addAttribute("nonyusakiList", Collections.emptyList());
	        model.addAttribute("plantList", Collections.emptyList());
	        model.addAttribute("brandList", Collections.emptyList());
	        model.addAttribute("kikiList", Collections.emptyList());
	        return "mcm0010u/index";
	    }

	    List<NonyusakiEntity> nonyusakiList = service.searchNonyusaki(form);

	    if (nonyusakiList.isEmpty()) {
	        model.addAttribute("errorMessage", "検索結果が1件も存在しません。");
	        model.addAttribute("nonyusakiList", Collections.emptyList());
	        model.addAttribute("plantList", Collections.emptyList());
	        model.addAttribute("brandList", Collections.emptyList());
	        model.addAttribute("kikiList", Collections.emptyList());
	        return "mcm0010u/index";
	    }

	    String authority = getAuthorityDivision();
	    List<PlantEntity> plantList = service.searchPlants(form, authority);
	    List<Mcm0010uBrandKoseiDto> brandList = service.searchBrandKosei(form);
	    List<Mcm0010uKikiKoseiDto> kikiList = service.searchKikiKosei(form);

	    model.addAttribute("nonyusakiList", nonyusakiList);
	    model.addAttribute("plantList", plantList);
	    model.addAttribute("brandList", brandList);
	    model.addAttribute("kikiList", kikiList);

	    return "mcm0010u/index";
	}

	// ===== ヘルパー（★追加） =====
	/**
	 * 対象文字列に全角文字が含まれるか判定する。
	 * Shift-JIS(MS932)換算のバイト数と文字数が一致しない場合、
	 * 2バイト文字（全角）が含まれると判定する。
	 */
	private boolean containsZenkaku(String value) {
	    if (value == null || value.isEmpty()) {
	        return false;
	    }
	    byte[] bytes = value.getBytes(java.nio.charset.Charset.forName("MS932"));
	    return bytes.length != value.length();
	}


    /**
     * プラント一覧取得（納入先選択時のリレーション連動 — AJAX用）
     * 元VB: DataRelation MCM_MA_PLANT_MCM_MA_NONYUSAKI
     */
    @GetMapping("/api/nonyusaki/{nonyusakiId}/plants")
    @ResponseBody
    public List<PlantEntity> getPlantsByNonyusakiId(@PathVariable BigDecimal nonyusakiId) {
        return service.getPlantsByNonyusakiId(nonyusakiId);
    }

    /**
     * ブランド構成取得（プラント選択時のリレーション連動 — AJAX用）
     * 元VB: DataRelation MCM_MA_PLANT_MCM_MA_BRAND_KOSEI
     * セッションから全件データをplantIdでフィルタ
     */
    @GetMapping("/api/plant/{plantId}/brands")
    @ResponseBody
    public List<Mcm0010uBrandKoseiDto> getBrandsByPlantId(
            @PathVariable BigDecimal plantId,
            @SessionAttribute(name = "allBrands", required = false) List<Mcm0010uBrandKoseiDto> allBrands) {
        if (allBrands != null) {
            return allBrands.stream()
                .filter(b -> plantId.equals(b.getPlantId()))
                .collect(Collectors.toList());
        }
        return Collections.emptyList();
    }

    /**
     * 機器構成取得（プラント選択時のリレーション連動 — AJAX用）
     * 元VB: DataRelation MCM_MA_PLANT_MCM_MA_KIKIKOSEI
     */
    @GetMapping("/api/plant/{plantId}/kikikosei")
    @ResponseBody
    public List<Mcm0010uKikiKoseiDto> getKikiByPlantId(
            @PathVariable BigDecimal plantId,
            @SessionAttribute(name = "allKiki", required = false) List<Mcm0010uKikiKoseiDto> allKiki) {
        if (allKiki != null) {
            return allKiki.stream()
                .filter(k -> plantId.equals(k.getPlantId()))
                .collect(Collectors.toList());
        }
        return Collections.emptyList();
    }

    /**
     * プラント行削除
     * 元VB: RowDeleteButton_Click → McmDBUtility.DeletePlant(plantId)
     */
    @PostMapping("/plant/delete")
    public String deletePlant(@RequestParam BigDecimal plantId, RedirectAttributes ra) {
        service.deletePlant(plantId);
        ra.addFlashAttribute("infoMessage", "プラントを削除しました。");
        return "redirect:/mcm0010u";
    }

    /**
     * ブランド構成データ存在チェック
     * 元VB: MCM_MA_BRAND_KOSEI.Select("PLANT_ID = " & plantId).Length
     */
    @GetMapping("/api/plant/{plantId}/hasBrand")
    @ResponseBody
    public boolean hasBrandKosei(@PathVariable BigDecimal plantId) {
        return service.hasBrandKosei(plantId);
    }

    /** #294: 別タブやJavaScript無効時も、機器構成リンクからの遷移前にDBを確認する。 */
    @GetMapping("/kikikosei")
    public String navigateKikikosei(@RequestParam BigDecimal plantId, RedirectAttributes ra) {
        if (plantId.signum() <= 0 || plantId.stripTrailingZeros().scale() > 0
                || plantId.compareTo(BigDecimal.valueOf(Long.MAX_VALUE)) > 0
                || !service.hasBrandKosei(plantId)) {
            ra.addFlashAttribute("errorMessage", MSG_0110);
            return "redirect:/mcm0010u";
        }
        return "redirect:/mcm0013u?plantId=" + plantId.toBigIntegerExact()
                + "&returnTo0012Url=/mcm0010u";
    }

    // ===== ヘルパー（BaseControllerで提供想定） =====
    private boolean isMaintenance() {
        // TODO: CPSettingInfo.GetUserInfo.Maintenance に相当するユーザ権限チェック
        // 認証モジュール連携後に実装
        return true;
    }

    private String getAuthorityDivision() {
        // TODO: GetAuthorityDivision() に相当する権限区分取得
        // CPConstant.AUTHORITY_DIVISION_INSPECTION or AUTHORITY_DIVISION_UPDATE
        return "UPDATE";
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
