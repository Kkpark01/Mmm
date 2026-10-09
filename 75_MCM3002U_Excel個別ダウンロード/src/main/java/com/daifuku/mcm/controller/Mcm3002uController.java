package com.daifuku.mcm.controller;

import java.util.List;

import jakarta.servlet.http.HttpSession;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.InitBinder;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.form.Mcm3002uForm;
import com.daifuku.mcm.form.Mcm3002uForm.RowForm;
import com.daifuku.mcm.service.Mcm3002uExcelService;
import com.daifuku.mcm.service.Mcm3002uService;

/**
 * 【変換元】Mcm3002uScreen.vb
 * MCM3002U 支払い明細出力指示コントローラ
 *
 * 元イベント対応:
 *   Mcm3002uScreen_Load → index (GET /mcm3002u)
 *   OutputButton_Click  → search (POST /mcm3002u/search)
 *
 * OutputButton_Click → output。2帳票のxlsを個別ダウンロードするためのデータを返す。
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
        // VB版の画面読込と同様、通常の表示では当月・年払いに初期化する。
        // 出力エラー後と既存の検索画面からの戻りでは、直前の入力条件を復元する。
        boolean restore = model.containsAttribute("error") || model.containsAttribute("errors") || model.containsAttribute("mcm3002uSearchReturn");
        Mcm3002uForm form = restore ? (Mcm3002uForm) session.getAttribute(SESSION_FORM) : null;
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
        ra.addFlashAttribute("mcm3002uSearchReturn", true);
        return "redirect:/mcm3002u";
    }

    /**
     * 帳票のブラウザー取得（fetch）。同一画面で2帳票を受け取り、
     * 待機中は common.js の fetch フックが共通のぐるぐるを表示する。
     */
    @PostMapping(value = "/output", headers = "X-Requested-With=XMLHttpRequest")
    public org.springframework.http.ResponseEntity<?> download(@ModelAttribute Mcm3002uForm form) {
        try {
            var rows = service.searchForOutput(form);
            if (rows.isEmpty()) return downloadError(400, "対象データが存在しません。");
            String month = form.getShiharaiTsuki().trim().replace("/", "");
            var reports = excelService.generateReports(month, form.getSiharaiKbn(), rows);
            // byte[]はJacksonでBase64として返す。ZIPやサーバー上の一時ファイルは作成しない。
            return org.springframework.http.ResponseEntity.ok()
                    .header(HttpHeaders.CACHE_CONTROL, "no-store")
                    .contentType(org.springframework.http.MediaType.APPLICATION_JSON)
                    .body(java.util.Map.of("reports", reports));
        } catch (IllegalArgumentException e) {
            logger.warn("MCM3002U 帳票出力を中止しました", e);
            return downloadError(400, e.getMessage());
        } catch (Exception e) {
            logger.error("MCM3002U 帳票出力に失敗しました", e);
            return downloadError(500, "帳票を出力できませんでした。");
        }
    }

    private org.springframework.http.ResponseEntity<?> downloadError(int status, String message) {
        return org.springframework.http.ResponseEntity.status(status).header(HttpHeaders.CACHE_CONTROL, "no-store")
                .body(java.util.Map.of("message", message));
    }

    @PostMapping(value = "/output", headers = "!X-Requested-With")
    public String output(@ModelAttribute Mcm3002uForm form, HttpSession session, RedirectAttributes ra,
                         jakarta.servlet.http.HttpServletResponse response) {
        // 通常フォーム送信では2ファイルを返せないため、誤ってZIPや片方だけを返さない。
        session.setAttribute(SESSION_FORM, form);
        ra.addFlashAttribute("error", "帳票の出力にはJavaScriptを有効にしてください。");
        return "redirect:/mcm3002u";
    }

    @Override
    protected String getScreenTitle() { return "支払い明細出力指示"; }

    @Override
    protected String getFunctionId() { return "MCM3002U"; }
}
