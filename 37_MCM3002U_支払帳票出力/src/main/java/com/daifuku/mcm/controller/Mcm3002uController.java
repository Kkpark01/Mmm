package com.daifuku.mcm.controller;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.form.Mcm3002uForm;
import com.daifuku.mcm.form.Mcm3002uForm.RowForm;
import com.daifuku.mcm.service.Mcm3002uService;
import com.daifuku.mcm.service.Mcm3002uExcelService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ContentDisposition;
import java.nio.charset.StandardCharsets;
import jakarta.servlet.http.HttpSession;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.util.List;

/**
 * 【変換元】Mcm3002uScreen.vb
 * MCM3002U 支払い明細出力指示コントローラ
 *
 * 元イベント対応:
 *   Mcm3002uScreen_Load → index (GET /mcm3002u)
 *   OutputButton_Click  → search (POST /mcm3002u/search)
 *
 * OutputButton_Click → output。2帳票のxlsをZIPでまとめて返す。
 */
@Controller
@RequestMapping("/mcm3002u")
public class Mcm3002uController extends BaseController {

    @Autowired
    private Mcm3002uService service;

    @Autowired
    private Mcm3002uExcelService excelService;

    @InitBinder
    public void bindConditions(org.springframework.web.bind.WebDataBinder binder) {
        binder.setAllowedFields("shiharaiTsuki", "siharaiKbn");
    }

    private static final String SESSION_FORM = "MCM3002U_FORM";

    // ===================================================================
    // 初期表示
    // 【変換元】Mcm3002uScreen_Load
    //   支払い月: システム日付の YYYYMM（Form のデフォルト値で対応）
    //   支払区分: 年払（Form のデフォルト値で対応）
    // ===================================================================

    @GetMapping
    public String index(Model model, HttpSession session) {
        Mcm3002uForm form = (Mcm3002uForm) session.getAttribute(SESSION_FORM);
        if (form == null) {
            form = new Mcm3002uForm();
            session.setAttribute(SESSION_FORM, form);
        }
        model.addAttribute("form", form);
        return "mcm3002u/index";
    }

    // ===================================================================
    // 検索（VB版: OutputButton_Click の Fill + 帳票出力に相当）
    // 【変換元】Mcm3002uScreen.vb OutputButton_Click
    //   Web版の検索一覧表示は既存操作として維持する。
    //   0件の場合はメッセージ表示。
    // ===================================================================

    @PostMapping("/search")
    public String search(@ModelAttribute Mcm3002uForm form,
                         HttpSession session, RedirectAttributes ra) {
        try {
            List<RowForm> rows = service.search(form);
            form.setRows(rows);
            session.setAttribute(SESSION_FORM, form);
            if (rows.isEmpty()) {
                // 【変換元】VB版 MSG_0002「対象データが存在しません」
                ra.addFlashAttribute("error", "対象データが存在しません。");
            }
        } catch (IllegalArgumentException e) {
            ra.addFlashAttribute("error", e.getMessage());
            session.setAttribute(SESSION_FORM, form);
        } catch (Exception e) {
            ra.addFlashAttribute("error", "検索中にエラーが発生しました: " + e.getMessage());
            session.setAttribute(SESSION_FORM, form);
        }
        return "redirect:/mcm3002u";
    }

    @PostMapping("/output")
    public String output(@ModelAttribute Mcm3002uForm form, HttpSession session, RedirectAttributes ra,
                         jakarta.servlet.http.HttpServletResponse response) {
        long started = System.nanoTime();
        // 入力エラー・DB障害・帳票生成失敗のいずれも入力条件を残す。
        session.setAttribute(SESSION_FORM, form);
        try {
            var rows = service.searchForOutput(form);
            form.setRows(rows);
            if (rows.isEmpty()) {
                logger.info("MCM3002U 帳票出力対象なし");
                ra.addFlashAttribute("error", "対象データが存在しません。");
                return "redirect:/mcm3002u";
            }
            String month = form.getShiharaiTsuki().trim();
            byte[] data = excelService.generateZip(month, form.getSiharaiKbn(), rows);
            String filename = month + "_" + form.getSiharaiKbn() + "_支払帳票.zip";
            response.setContentType("application/zip");
            response.setHeader(HttpHeaders.CONTENT_DISPOSITION, ContentDisposition.attachment()
                    .filename(filename, StandardCharsets.UTF_8).build().toString());
            response.setHeader(HttpHeaders.CACHE_CONTROL, "no-store");
            response.setContentLength(data.length);
            response.getOutputStream().write(data);
            response.getOutputStream().flush();
            logger.info("MCM3002U 帳票送信完了: 明細件数={}, bytes={}, elapsedMs={}",
                    rows.size(), data.length, (System.nanoTime() - started) / 1_000_000);
            return null;
        } catch (IllegalArgumentException e) {
            // 期間・金額等の不備やPOIの境界エラーも、画面案内だけで終わらせない。
            logger.warn("MCM3002U 帳票出力を中止しました", e);
            ra.addFlashAttribute("error", e.getMessage());
        } catch (Exception e) {
            logger.error("MCM3002U 帳票出力に失敗しました", e);
            if (response.isCommitted()) return null;
            response.reset();
            ra.addFlashAttribute("error", "帳票を出力できませんでした。入力条件を保持しています。");
        }
        return "redirect:/mcm3002u";
    }

    @Override
    protected String getScreenTitle() { return "支払い明細出力指示"; }

    @Override
    protected String getFunctionId() { return "MCM3002U"; }
}
