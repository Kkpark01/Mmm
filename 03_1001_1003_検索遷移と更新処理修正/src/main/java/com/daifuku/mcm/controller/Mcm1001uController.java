package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.List;
import jakarta.servlet.http.HttpSession;
import org.springframework.dao.DataAccessException;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.dto.Mcm1001uPlantDto;
import com.daifuku.mcm.form.Mcm1001uForm;
import com.daifuku.mcm.service.Mcm1001uService;

/** 取引先見積依頼作成検索。子画面から戻った場合は検索時の条件・結果を復元する。 */
@Controller
@RequestMapping("/mcm1001u")
public class Mcm1001uController extends BaseController {
    private static final String PREFIX = "mcm1001u.";
    private final Mcm1001uService service;
    public Mcm1001uController(Mcm1001uService service) { this.service = service; }

    @GetMapping
    public String index(@RequestParam(defaultValue = "false") boolean fresh, Model model, HttpSession session) {
        if (fresh) {
            session.removeAttribute(PREFIX + "form");
            clearResults(session);
        }
        Object form = session.getAttribute(PREFIX + "form");
        if (!model.containsAttribute("form")) model.addAttribute("form", form != null ? form : new Mcm1001uForm());
        for (String key : List.of("nonyusakiList", "plantList")) {
            Object rows = session.getAttribute(PREFIX + key);
            if (!model.containsAttribute(key)) model.addAttribute(key, rows != null ? rows : List.of());
        }
        return "mcm1001u/index";
    }

    @PostMapping("/search")
    public String search(@ModelAttribute("form") Mcm1001uForm form, RedirectAttributes ra, HttpSession session) {
        session.setAttribute(PREFIX + "form", form);
        clearResults(session);
        List<String> errors = service.validate(form.getNonyusakiCd(), form.getSupportId(),
                form.getNonyusakiNk(), form.getPlantNk(), form.getEtc());
        if (!errors.isEmpty()) {
            ra.addFlashAttribute("errors", errors);
            return "redirect:/mcm1001u";
        }
        try {
            var nonyusakiList = service.searchNonyusaki(form.getNonyusakiCd(), form.getSupportId(),
                    form.getNonyusakiNk(), form.getPlantNk(), form.getEtc());
            var plantList = nonyusakiList.isEmpty() ? List.of() : service.searchPlant(form.getNonyusakiCd(), form.getSupportId(),
                    form.getNonyusakiNk(), form.getPlantNk(), form.getEtc());
            session.setAttribute(PREFIX + "nonyusakiList", nonyusakiList);
            session.setAttribute(PREFIX + "plantList", plantList);
            if (nonyusakiList.isEmpty()) ra.addFlashAttribute("errors", List.of("検索結果が1件も存在しません。"));
        } catch (DataAccessException error) {
            logger.warn("MCM1001U 検索を中断", error);
            ra.addFlashAttribute("errors", List.of("処理中にエラーが発生しました。再度検索してから操作してください。"));
        }
        return "redirect:/mcm1001u";
    }

    @GetMapping("/plant")
    @ResponseBody
    public List<Mcm1001uPlantDto> getPlant(@RequestParam BigDecimal nonyusakiId) {
        return service.getPlantByNonyusakiId(nonyusakiId);
    }

    private void clearResults(HttpSession session) {
        session.setAttribute(PREFIX + "nonyusakiList", List.of());
        session.setAttribute(PREFIX + "plantList", List.of());
    }
    @Override protected String getScreenTitle() { return "取引先見積依頼作成検索"; }
    @Override protected String getFunctionId() { return "MCM1001U"; }
}
