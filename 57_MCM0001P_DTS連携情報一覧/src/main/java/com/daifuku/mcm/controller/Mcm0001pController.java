package com.daifuku.mcm.controller;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.common.Mcm0020uConstants;
import com.daifuku.mcm.dto.Mcm0001pData;
import com.daifuku.mcm.dto.Mcm0020uSearchState;
import com.daifuku.mcm.service.Mcm0001pExcelService;
import jakarta.servlet.http.HttpSession;
import org.springframework.context.MessageSource;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.Locale;

/** VBのExecuteReport(MCM0001P, DataGridView)に対応するExcelダウンロード。 */
@Controller
@RequestMapping("/mcm0001p")
public class Mcm0001pController extends BaseController {
    private final Mcm0001pExcelService excelService;
    private final MessageSource messages;

    public Mcm0001pController(Mcm0001pExcelService excelService, MessageSource messages) {
        this.excelService = excelService;
        this.messages = messages;
    }

    @Override protected String getScreenTitle() { return "DTS連携情報一覧"; }
    @Override protected String getFunctionId() { return "MCM0001P"; }

    @GetMapping("/download")
    public Object download(@RequestParam(required = false) String state, HttpSession session,
                           RedirectAttributes ra) {
        Mcm0020uSearchState saved = Mcm0020uSearchState.find(session, state);
        if (saved == null) {
            ra.addFlashAttribute("error", state == null || state.isBlank()
                ? Mcm0020uConstants.MSG_NO_SEARCH_RESULT
                : "検索結果の保持期限が切れました。検索をやり直してください。");
            return "redirect:/mcm0020u";
        }
        if (saved.results().isEmpty()) {
            ra.addFlashAttribute("error", Mcm0020uConstants.MSG_NO_SEARCH_RESULT);
            ra.addAttribute("state", state);
            return "redirect:/mcm0020u/return";
        }
        // 全タブ共通の一時データを使わず、検索済みの表示用スナップショットだけを渡す。
        var data = saved.results().stream().map(row -> {
            var item = new Mcm0001pData();
            item.setLogDt(row.getLogDt());
            item.setNonyusakiCd(row.getNonyusakiCd());
            item.setNonyusakiNk(row.getNonyusakiNk());
            item.setSupportId(row.getSupportIdHyoji());
            item.setPlantNk(row.getPlantNk());
            item.setNaiyo(row.getNaiyo());
            return item;
        }).toList();
        try {
            byte[] bytes = excelService.generate(data);
            return ResponseEntity.ok()
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .header(HttpHeaders.CONTENT_DISPOSITION, ContentDisposition.attachment()
                    .filename(Mcm0001pExcelService.FILE_NAME, StandardCharsets.UTF_8).build().toString())
                .header(HttpHeaders.CACHE_CONTROL, "no-store")
                .body(bytes);
        } catch (IOException | RuntimeException ex) {
            logger.error("[MCM0001P] Excel出力に失敗しました", ex);
            ra.addFlashAttribute("error", messages.getMessage("FWM_0014E", null, Locale.JAPANESE));
            ra.addAttribute("state", state);
            return "redirect:/mcm0020u/return";
        }
    }
}
