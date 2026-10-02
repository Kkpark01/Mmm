package com.daifuku.mcm.controller;

import java.nio.charset.StandardCharsets;
import java.util.List;

import jakarta.servlet.http.HttpServletResponse;
import jakarta.servlet.http.HttpSession;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.validation.BindingResult;
import java.util.Map;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.WebDataBinder;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.InitBinder;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.constants.Mcm3002pConstants;
import com.daifuku.mcm.constants.Mcm3003pConstants;
import com.daifuku.mcm.form.Mcm3003uForm;
import com.daifuku.mcm.form.Mcm3003uForm.RowForm;
import com.daifuku.mcm.service.Mcm3003uExcelService;
import com.daifuku.mcm.service.Mcm3003uService;

/**
 * 【変換元】Mcm3003uScreen.vb
 * MCM3003U 保守点検出力指示コントローラ
 *
 * 元イベント対応:
 *   Mcm3003uScreen_Load           → index   (GET  /mcm3003u)
 *   DBTorokuyoButton_Click        → outputDbTorokuyo  (POST /mcm3003u/output/db)
 *   SagyoyoListOutputButton_Click → outputSagyoyoList (POST /mcm3003u/output/list)
 *
 * VB版は一覧グリッドを持たない「帳票出力指示」専用画面であるため、
 * Web版でも検索一覧・検索ボタンは設けず、2つの出力ボタンから直接
 * Excelファイル（.xls）をダウンロードさせる。
 * VB版の SaveFileDialog ＋ 保存後のExcel自動起動はブラウザでは再現できないため、
 * Content-Disposition: attachment によるダウンロードに置き換えている。
 *
 * 【移植】画面アクセス権限（機能ID=MCM3003U）は全画面共通の
 *   com.daifuku.mcm.config.ScreenAuthorityInterceptor（SESSION_DENIED_FUNCTION_IDS参照）
 *   により URL の先頭セグメントから一元的に判定される。サブパス（/output/**）も
 *   同様に保護されるため、個別の権限チェックは不要。
 */
@Controller
@RequestMapping("/mcm3003u")
public class Mcm3003uController extends BaseController {

    @Autowired
    private Mcm3003uService service;

    @Autowired
    private Mcm3003uExcelService excelService;

    private static final String SESSION_FORM = "MCM3003U_FORM";

    /** 出力条件のみをバインド対象とする（意図しないプロパティの外部設定を防ぐ）。 */
    @InitBinder
    public void bindConditions(WebDataBinder binder) {
        binder.setAllowedFields("fromTsuki", "toTsuki", "syutsuryokuKbn", "sabunFrom", "sabunTo");
    }

    // ===================================================================
    // 初期表示
    // 【変換元】Mcm3003uScreen_Load
    //   ZenkenRadioButton.Checked = True（全件出力）
    //   FromTextBox / ToTextBox = システム日付の YYYY/MM
    //   ※ Form のフィールド初期値で対応
    // ===================================================================

    @GetMapping
    public String index(Model model, HttpSession session) {
        // 出力エラー後は入力条件を復元する。通常表示ではVB版の画面読込と同様に当月・全件へ初期化する。
        boolean restore = model.containsAttribute("error") || model.containsAttribute("errors");
        Mcm3003uForm form = model.containsAttribute("form")
                ? (Mcm3003uForm) model.getAttribute("form")
                : restore ? (Mcm3003uForm) session.getAttribute(SESSION_FORM) : null;
        if (form == null) {
            form = new Mcm3003uForm();
        }
        session.setAttribute(SESSION_FORM, form);
        model.addAttribute("form", form);
        return "mcm3003u/index";
    }

    // ===================================================================
    // ＤＢ登録用データ出力
    // 【変換元】Mcm3003uScreen.vb DBTorokuyoButton_Click
    //   Search() → 0件なら MSG_0002 → Mcm3002pExcel で出力
    // ===================================================================

    @PostMapping(value = "/output/db", headers = "!X-Requested-With")
    public String outputDbTorokuyo(@ModelAttribute("form") Mcm3003uForm form, BindingResult binding, HttpSession session,
                                   RedirectAttributes ra, HttpServletResponse response) {
        if (binding.hasErrors()) return inputError(form, ra);
        return output(form, session, ra, response, true);
    }

    /** DB登録用データのブラウザー取得。入力条件はリクエスト単位で扱い、別タブと混ぜない。 */
    @PostMapping(value = "/output/db", headers = "X-Requested-With=XMLHttpRequest")
    public ResponseEntity<?> downloadDbTorokuyo(@ModelAttribute("form") Mcm3003uForm form,
                                               BindingResult binding) {
        if (binding.hasErrors()) return downloadError(400, "出力区分を選択してください。");
        try {
            List<RowForm> rows = service.searchForOutput(form);
            if (rows.isEmpty()) return downloadError(400, "検索結果が1件も存在しません。");
            byte[] data = excelService.generateDbTorokuyo(rows);
            String filename = service.buildFileNk(form) + Mcm3002pConstants.CREATE_FILE_NAME;
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, ContentDisposition.attachment()
                            .filename(filename, StandardCharsets.UTF_8).build().toString())
                    .header(HttpHeaders.CACHE_CONTROL, "no-store")
                    .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
                    .contentLength(data.length).body(data);
        } catch (IllegalArgumentException e) {
            return downloadError(400, e.getMessage());
        } catch (Exception e) {
            logger.error("MCM3002P 帳票出力に失敗しました", e);
            return downloadError(500, "帳票を出力できませんでした。");
        }
    }

    private ResponseEntity<?> downloadError(int status, String message) {
        return ResponseEntity.status(status).header(HttpHeaders.CACHE_CONTROL, "no-store")
                .body(Map.of("message", message));
    }

    private String inputError(Mcm3003uForm form, RedirectAttributes ra) {
        ra.addFlashAttribute("form", form);
        ra.addFlashAttribute("error", "出力区分を選択してください。");
        return "redirect:/mcm3003u";
    }

    // ===================================================================
    // 作業用リスト出力
    // 【変換元】Mcm3003uScreen.vb SagyoyoListOutputButton_Click
    //   Search() → 0件なら MSG_0002 → Mcm3003pExcel で出力
    //   ※ VB版は0件でも Exit Sub せず空ファイルを作成してしまうが、
    //     Web版はメッセージのみ表示しダウンロードを行わない。
    // ===================================================================

    @PostMapping("/output/list")
    public String outputSagyoyoList(@ModelAttribute("form") Mcm3003uForm form, BindingResult binding, HttpSession session,
                                    RedirectAttributes ra, HttpServletResponse response) {
        if (binding.hasErrors()) return inputError(form, ra);
        return output(form, session, ra, response, false);
    }

    /**
     * 帳票出力の共通処理。
     *
     * @param dbTorokuyo true=ＤＢ登録用データ（MCM3002P） / false=作業用リスト（MCM3003P）
     */
    private String output(Mcm3003uForm form, HttpSession session, RedirectAttributes ra,
                          HttpServletResponse response, boolean dbTorokuyo) {
        long started = System.nanoTime();
        String screen = dbTorokuyo ? "MCM3003U ＤＢ登録用データ" : "MCM3003U 作業用リスト";
        // 入力エラー・DB障害・帳票生成失敗のいずれの場合も入力条件を残す。
        session.setAttribute(SESSION_FORM, form);
        ra.addFlashAttribute("form", form);
        try {
            List<RowForm> rows = service.searchForOutput(form);
            if (rows.isEmpty()) {
                // 【変換元】MSG_0002E「検索結果が1件も存在しません。」
                logger.info("{} 出力対象なし", screen);
                ra.addFlashAttribute("error", "検索結果が1件も存在しません。");
                return "redirect:/mcm3003u";
            }

            // 【変換元】Delivery_Change() → Mcm300xpExcel.Initialize の CreateFileName
            String fileNk = service.buildFileNk(form);
            String filename = fileNk + (dbTorokuyo
                    ? Mcm3002pConstants.CREATE_FILE_NAME
                    : Mcm3003pConstants.CREATE_FILE_NAME);

            byte[] data = dbTorokuyo
                    ? excelService.generateDbTorokuyo(rows)
                    : excelService.generateSagyoyoList(rows);

            // 原紙が .xls（HSSF）形式のため、生成物も同形式で返す
            response.setContentType("application/vnd.ms-excel");
            response.setHeader(HttpHeaders.CONTENT_DISPOSITION, ContentDisposition.attachment()
                    .filename(filename, StandardCharsets.UTF_8).build().toString());
            response.setHeader(HttpHeaders.CACHE_CONTROL, "no-store");
            response.setContentLength(data.length);
            response.getOutputStream().write(data);
            response.getOutputStream().flush();
            logger.info("{} 送信完了: 件数={}, bytes={}, elapsedMs={}",
                    screen, rows.size(), data.length, (System.nanoTime() - started) / 1_000_000);
            return null;
        } catch (IllegalArgumentException e) {
            // 入力チェックエラー・行数上限超過など
            logger.warn("{} 出力を中止しました", screen, e);
            ra.addFlashAttribute("error", e.getMessage());
        } catch (Exception e) {
            logger.error("{} 出力に失敗しました", screen, e);
            if (response.isCommitted()) {
                return null;
            }
            response.reset();
            ra.addFlashAttribute("error", "帳票を出力できませんでした。");
        }
        return "redirect:/mcm3003u";
    }

    @Override
    protected String getScreenTitle() { return "保守点検出力指示"; }

    @Override
    protected String getFunctionId() { return "MCM3003U"; }
}
