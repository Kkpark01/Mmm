/**
 * 【変換元】Mcm1001uScreen.vb - Screen_Load() / SearchButton_Click()
 *   元ファイル行数: 約4,296行（Screen.vb全体）
 *   MCM1001U（取引先見積依頼作成検索）コントローラ
 *
 * @author MCM Migration Tool
 * @since v8
 */
package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.Collections;
import java.util.List;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.dto.Mcm1001uNonyusakiDto;
import com.daifuku.mcm.dto.Mcm1001uPlantDto;
import com.daifuku.mcm.form.Mcm1001uForm;
import com.daifuku.mcm.service.Mcm1001uService;

/**
 * MCM1001U 取引先見積依頼作成検索 コントローラ.
 *
 * 【変換元】Mcm1001uScreen.vb
 *   - Screen_Load() → index() [GET /mcm1001u]
 *   - SearchButton_Click() → search() [POST /mcm1001u/search]
 *   - DataGridView_NONYUSAKI_SelectionChanged → getPlant() [GET /mcm1001u/plant] (Ajax)
 */
@Controller
@RequestMapping("/mcm1001u")
public class Mcm1001uController extends BaseController {

    private final Mcm1001uService service;

    public Mcm1001uController(Mcm1001uService service) {
        this.service = service;
    }

    // ========================================================
    // 初期表示
    // 【変換元】Mcm1001uScreen.vb - McmXXXXxScreen_Load()
    // ========================================================

    @GetMapping
    public String index(Model model) {
        if (!model.containsAttribute("form")) {
            model.addAttribute("form", new Mcm1001uForm());
        }
        if (!model.containsAttribute("nonyusakiList")) {
            model.addAttribute("nonyusakiList", Collections.emptyList());
        }
        if (!model.containsAttribute("plantList")) {
            model.addAttribute("plantList", Collections.emptyList());
        }
        return "mcm1001u/index";
    }

    // ========================================================
    // 検索（PRGパターン）
    // 【変換元】Mcm1001uScreen.vb - SearchButton_Click()
    // ========================================================

    @PostMapping("/search")
    public String search(@ModelAttribute("form") Mcm1001uForm form,
                         RedirectAttributes ra) {

        List<String> errors = service.validate(
                form.getNonyusakiCd(), form.getSupportId(),
                form.getNonyusakiNk(), form.getPlantNk(), form.getEtc());
        if (!errors.isEmpty()) {
            ra.addFlashAttribute("errors", errors);
            ra.addFlashAttribute("form", form);
            return "redirect:/mcm1001u";
        }

        List<Mcm1001uNonyusakiDto> nonyusakiList = service.searchNonyusaki(
                form.getNonyusakiCd(), form.getSupportId(),
                form.getNonyusakiNk(), form.getPlantNk(), form.getEtc());

        List<Mcm1001uPlantDto> plantList = service.searchPlant(
                form.getNonyusakiCd(), form.getSupportId(),
                form.getNonyusakiNk(), form.getPlantNk(), form.getEtc());

        if (nonyusakiList.isEmpty()) {
            ra.addFlashAttribute("message", "検索結果が1件も存在しません。");
        }

        ra.addFlashAttribute("form", form);
        ra.addFlashAttribute("nonyusakiList", nonyusakiList);
        ra.addFlashAttribute("plantList", plantList);

        return "redirect:/mcm1001u";
    }

    // ========================================================
    // プラント取得（Ajax用）
    // 【変換元】Mcm1001uScreen.vb - DataGridView_NONYUSAKI_SelectionChanged
    // ========================================================

    @GetMapping("/plant")
    @ResponseBody
    public List<Mcm1001uPlantDto> getPlant(
            @RequestParam("nonyusakiId") BigDecimal nonyusakiId) {
        return service.getPlantByNonyusakiId(nonyusakiId);
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
