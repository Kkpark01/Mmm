package com.daifuku.mcm.controller;

import com.daifuku.mcm.common.AppConstants;
import com.daifuku.mcm.common.Mcm1002uConstants;
import com.daifuku.mcm.constants.Mcm1006uConstants;
import com.daifuku.mcm.dto.Mcm1003uRowDto;
import com.daifuku.mcm.exception.AuthorityException;
import com.daifuku.mcm.form.Mcm1002uDeliveryDto;
import com.daifuku.mcm.form.Mcm1003uForm;
import com.daifuku.mcm.form.Mcm1006uForm;
import com.daifuku.mcm.service.Mcm1003uService;
import jakarta.servlet.http.HttpSession;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.dao.DataAccessException;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.math.BigDecimal;
import java.util.List;

/** MCM1003U。VBの機能別権限・検索結果保持・Delivery・破棄/行削除に対応する。 */
@Controller
@RequestMapping("/mcm1003u")
public class Mcm1003uController {
    private static final Logger log = LoggerFactory.getLogger(Mcm1003uController.class);
    private static final String SESSION_FORM = "mcm1003u.form";
    private static final String SESSION_RESULTS = "mcm1003u.searchResults";
    private static final String SESSION_MCM1002P_DATA = "mcm1002pDataList";
    private final Mcm1003uService service;

    public Mcm1003uController(Mcm1003uService service) { this.service = service; }

    @ModelAttribute
    public void commonModel(Model model, HttpSession session) {
        String loginId = loginId(session);
        model.addAttribute("iraisyaList", List.of());
        model.addAttribute("canEditMitsumori", false);
        model.addAttribute("canEditKeiyaku", false);
        model.addAttribute("canDelete", false);
        try {
            model.addAttribute("iraisyaList", service.getIraitantosyaList());
            model.addAttribute("canEditMitsumori", service.canUpdate(loginId, "MCM1002U"));
            model.addAttribute("canEditKeiyaku", service.canUpdate(loginId, "MCM1005U"));
            model.addAttribute("canDelete", service.canDelete(loginId));
        } catch (DataAccessException error) {
            // DB障害時も同じURLへのリダイレクトを繰り返さず、画面で通知する。
            log.warn("MCM1003U 初期データ取得を中断", error);
            model.addAttribute("errors", List.of("処理中にエラーが発生しました。再度検索してから操作してください。"));
        }
    }

    @GetMapping
    public String index(Model model, HttpSession session) {
        Object savedForm = session.getAttribute(SESSION_FORM);
        if (!model.containsAttribute("form")) model.addAttribute("form", savedForm != null ? savedForm : new Mcm1003uForm());
        List<Mcm1003uRowDto> rows = savedRows(session);
        model.addAttribute("list", rows);
        model.addAttribute("resultCount", rows.size());
        return "mcm1003u/index";
    }

    @PostMapping("/search")
    public String search(@ModelAttribute("form") Mcm1003uForm form, Model model,
                         HttpSession session, RedirectAttributes ra) {
        // 検索開始時点で旧一覧/帳票を失効。0件・入力エラー・SQL失敗でも旧行を出力しない。
        session.setAttribute(SESSION_FORM, form);
        clearResults(session);
        var result = service.search(form);
        session.setAttribute(SESSION_RESULTS, result.getRows());
        model.addAttribute("form", form);
        model.addAttribute("list", result.getRows());
        model.addAttribute("resultCount", result.getRows().size());
        if (result.hasErrors()) model.addAttribute("errors", result.getErrors());
        return "mcm1003u/index";
    }

    @GetMapping("/new")
    public String newMitsumori(HttpSession session) {
        service.requireUpdate(loginId(session), "MCM1002U");
        return "redirect:/mcm1001u?fresh=true";
    }

    @GetMapping("/close")
    public String close(HttpSession session) {
        session.removeAttribute(SESSION_FORM);
        clearResults(session);
        return "redirect:/menu";
    }

    @PostMapping("/discard-mitsumori")
    public String discardMitsumori(@RequestParam BigDecimal tmKeiyakujikanId,
            @RequestParam(required = false) BigDecimal tkKeiyakuId, HttpSession session, RedirectAttributes ra) {
        service.discardMitsumori(tmKeiyakujikanId, tkKeiyakuId, loginId(session));
        clearResults(session);
        return "redirect:/mcm1003u/re-search";
    }

    @PostMapping("/discard-keiyaku")
    public String discardKeiyaku(@RequestParam BigDecimal tkKeiyakuId,
            @RequestParam BigDecimal tmKeiyakujikanId, HttpSession session, RedirectAttributes ra) {
        service.discardKeiyaku(tkKeiyakuId, tmKeiyakujikanId, loginId(session));
        clearResults(session);
        return "redirect:/mcm1003u/re-search";
    }

    @PostMapping("/delete-row")
    public String deleteRow(@RequestParam BigDecimal tmIraiId, HttpSession session, RedirectAttributes ra) {
        service.deleteRow(tmIraiId, loginId(session));
        clearResults(session);
        return "redirect:/mcm1003u/re-search";
    }

    @GetMapping("/re-search")
    public String reSearchGet(HttpSession session, Model model) {
        Mcm1003uForm form = (Mcm1003uForm) session.getAttribute(SESSION_FORM);
        if (form == null) return "redirect:/mcm1003u";
        clearResults(session);
        var result = service.search(form);
        session.setAttribute(SESSION_RESULTS, result.getRows());
        model.addAttribute("form", form);
        model.addAttribute("list", result.getRows());
        model.addAttribute("resultCount", result.getRows().size());
        if (result.hasErrors()) model.addAttribute("errors", result.getErrors());
        return "mcm1003u/index";
    }

    @GetMapping("/excel-download")
    public String excelDownload(HttpSession session, RedirectAttributes ra) {
        var rows = savedRows(session);
        session.removeAttribute(SESSION_MCM1002P_DATA);
        if (rows.isEmpty()) {
            ra.addFlashAttribute("errors", List.of("検索結果が1件も存在しません。"));
            return "redirect:/mcm1003u";
        }
        session.setAttribute(SESSION_MCM1002P_DATA, service.buildExcelDataList(rows));
        return "redirect:/mcm1002p/download";
    }

    @GetMapping("/link/irai-no")
    public String linkIraiNo(@RequestParam BigDecimal tmKeiyakujikanId, HttpSession session, RedirectAttributes ra) {
        service.findRows(tmKeiyakujikanId);
        ra.addAttribute("tmKeiyakujikanId", tmKeiyakujikanId);
        ra.addAttribute("seniMotoKbn", 1); // VB SENIMOTO_MITSUMORI_LINK
        return "redirect:/mcm1004u";
    }

    @GetMapping("/link/keiyaku-no")
    public String linkKeiyakuNo(@RequestParam BigDecimal tkKeiyakuId,
            @RequestParam BigDecimal tmKeiyakujikanId, @RequestParam(required = false) BigDecimal plantId,
            HttpSession session, RedirectAttributes ra) {
        Mcm1003uRowDto row = findRow(tmKeiyakujikanId, plantId, tkKeiyakuId);
        ra.addAttribute("tkKeiyakuId", row.getTkKeiyakuId());
        ra.addAttribute("torihikisakiNk", row.getTorihikisakiNk());
        ra.addAttribute("nonyusakiNk", row.getNonyusakiNk());
        return "redirect:/mcm1005u";
    }

    @GetMapping("/link/mitsumori-irai")
    public String linkMitsumoriIrai(@RequestParam BigDecimal tmKeiyakujikanId,
            @RequestParam(required = false) BigDecimal plantId, HttpSession session) {
        var row = findRow(tmKeiyakujikanId, plantId, null);
        session.setAttribute("mcm1002uDelivery", delivery(row, row.getPlantId(), Mcm1002uConstants.SYUSEI, false));
        return "redirect:/mcm1002u";
    }

    @GetMapping("/link/mitsumori-copy")
    public String linkMitsumoriCopy(@RequestParam BigDecimal tmKeiyakujikanId,
            @RequestParam(required = false) BigDecimal plantId, HttpSession session, RedirectAttributes ra) {
        service.requireUpdate(loginId(session), "MCM1002U");
        var row = findRow(tmKeiyakujikanId, plantId, null);
        List<BigDecimal> plantIds = service.findPlantIds(tmKeiyakujikanId);
        if (plantIds.isEmpty()) throw new IllegalStateException("プラント情報にデータが存在しません。");
        if (plantIds.size() == 1) {
            session.setAttribute("mcm1002uDelivery", delivery(row, plantIds.get(0), Mcm1002uConstants.FUKUSEI, true));
            return "redirect:/mcm1002u";
        }
        // 1008Uの現在の受け口に合わせる。選択後に同画面がDeliveryを作成する。
        session.setAttribute("mcm1008u.plantIds", plantIds);
        session.setAttribute("mcm1002u.tmKeiyakujikanId", tmKeiyakujikanId.longValueExact());
        session.setAttribute("mcm1002u.jotai", row.getJotai());
        return "redirect:/mcm1008u";
    }

    @GetMapping("/link/keiyaku-create")
    public String linkKeiyakuCreate(@RequestParam BigDecimal tmKeiyakujikanId,
            @RequestParam(required = false) BigDecimal plantId, HttpSession session) {
        service.requireUpdate(loginId(session), "MCM1005U");
        var row = findRow(tmKeiyakujikanId, plantId, null);
        if (row.getPlantId() == null || row.getTorihikisakiId() == null
                || row.getKeiyakuKeiyaku() == null || row.getKeiyakuKeiyaku().isBlank()) {
            throw new IllegalStateException("取引先見積情報が存在しません。");
        }
        var delivery = new Mcm1006uForm();
        delivery.setSeniMotoKbn(Mcm1006uConstants.SENIMOTO_INSERT);
        delivery.setTmKeiyakujikanIds(List.of(tmKeiyakujikanId));
        delivery.setNonyusakiId(row.getNonyusakiId());
        delivery.setNonyusakiCd(row.getNonyusakiCd());
        delivery.setNonyusakiNk(row.getNonyusakiNk());
        delivery.setPlantId(row.getPlantId());
        delivery.setSupportId(row.getSupportId());
        delivery.setPlantNk(row.getPlantNk());
        delivery.setTorihikisakiId(row.getTorihikisakiId());
        delivery.setTorihikisakiCd(row.getTorihikisakiCd());
        delivery.setTorihikisakiNk(row.getTorihikisakiNk());
        session.setAttribute("MCM1006U_DELIVERY", delivery);
        return "redirect:/mcm1006u/step1";
    }

    private Mcm1003uRowDto findRow(BigDecimal tmId, BigDecimal plantId, BigDecimal contractId) {
        // 同じ見積時間でも複数プラント/契約の行があり得る。別行へ暗黙に差し替えない。
        return service.findRows(tmId).stream()
                .filter(row -> plantId == null || sameId(plantId, row.getPlantId()))
                .filter(row -> contractId == null || sameId(contractId, row.getTkKeiyakuId()))
                .findFirst().orElseThrow(() -> new IllegalStateException("取引先見積情報が存在しません。"));
    }

    private Mcm1002uDeliveryDto delivery(Mcm1003uRowDto row, BigDecimal plantId, int mode, boolean copy) {
        var delivery = new Mcm1002uDeliveryDto();
        delivery.setMitsumoriFlg(mode);
        delivery.setPlantId(plantId);
        delivery.setPlantNk(row.getPlantNk());
        delivery.setNonyusakiCd(row.getNonyusakiCd());
        delivery.setNonyusakiNk(row.getNonyusakiNk());
        delivery.setSupportId(row.getSupportId());
        delivery.setTmKeiyakujikanId(row.getTmKeiyakujikanId().longValueExact());
        delivery.setJotai(row.getJotai());
        delivery.setCopyFlg(copy);
        delivery.setReturnUrl("/mcm1003u");
        return delivery;
    }

    @ExceptionHandler({IllegalStateException.class, AuthorityException.class, DataAccessException.class})
    public String actionError(Exception error, RedirectAttributes ra) {
        log.warn("MCM1003U 操作を中断", error);
        String message = error instanceof DataAccessException
                ? "処理中にエラーが発生しました。再度検索してから操作してください。" : error.getMessage();
        ra.addFlashAttribute("errors", List.of(message));
        return "redirect:/mcm1003u";
    }

    @SuppressWarnings("unchecked")
    private List<Mcm1003uRowDto> savedRows(HttpSession session) {
        Object rows = session.getAttribute(SESSION_RESULTS);
        return rows instanceof List<?> ? (List<Mcm1003uRowDto>) rows : List.of();
    }
    private void clearResults(HttpSession session) {
        session.setAttribute(SESSION_RESULTS, List.of());
        session.removeAttribute(SESSION_MCM1002P_DATA);
    }
    private String loginId(HttpSession session) {
        Object loginId = session.getAttribute(AppConstants.SESSION_USER_ID);
        return loginId instanceof String ? (String) loginId : null;
    }
    private static boolean sameId(BigDecimal left, BigDecimal right) {
        return left != null && right != null && left.compareTo(right) == 0;
    }
}
