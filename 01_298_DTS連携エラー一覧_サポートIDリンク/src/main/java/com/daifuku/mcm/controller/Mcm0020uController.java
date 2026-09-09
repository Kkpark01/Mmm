/**
 * 【変換元】Mcm0020uScreen.vb（全130行）
 * MCM0020U — DTS連携エラー一覧 コントローラー
 *
 * パターンC: 検索・参照型（ReadOnly DataGridView）
 *
 * 画面フロー:
 *   1. 初期表示（GET）   → ラジオ=DTSログ, FROM=昨日, TO=空
 *   2. 検索ボタン（POST /search）→ モードに応じた検索実行
 *   3. Excel出力（POST /excel）  → MCM0001P帳票起動
 *   4. サポートIDリンク  → MCM0021Uへ遷移（plantId パラメータ付き）
 *
 * @since 2026-06-08
 */
package com.daifuku.mcm.controller;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.format.ResolverStyle;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import jakarta.servlet.http.HttpSession;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.common.Mcm0020uConstants;
import com.daifuku.mcm.dto.DtslogSearchResultDto;
import com.daifuku.mcm.form.Mcm0020uForm;
import com.daifuku.mcm.service.Mcm0020uService;

@Controller
@RequestMapping("/mcm0020u")
public class Mcm0020uController extends BaseController {

	// 抽出期間の日付フォーマット（存在しない日付を厳格に弾くため STRICT を指定）
	// ※ "yyyy"（Year of Era＝元号年）はEraが未指定だとSTRICTモードで解決不能になるため、
//	    西暦年を表す "uuuu"（Year）を使用する
	private static final DateTimeFormatter LOG_DATE_FORMATTER =
	        DateTimeFormatter.ofPattern("uuuu/MM/dd").withResolverStyle(ResolverStyle.STRICT);

    private final Mcm0020uService service;

    public Mcm0020uController(Mcm0020uService service) {
        this.service = service;
    }

    @Override
    protected String getScreenTitle() {
        return Mcm0020uConstants.SCREEN_TITLE;
    }

    @Override
    protected String getFunctionId() {
        return Mcm0020uConstants.SCREEN_ID;
    }

    // ========================================================================
    // 初期表示
    // 【変換元】Mcm0020uScreen.vb — Mcm0020uScreen_Load()
    //   元コード: Me.DtsMstRenkeiLogRadioButton.Checked = True
    //             Me.LogDateFromTextBox.Text = Format(Now.AddDays(-1), "yyyy/MM/dd")
    // ========================================================================
    @GetMapping
    public String index(Model model, HttpSession session) {

        // フォーム初期値
        Mcm0020uForm form = new Mcm0020uForm();
        form.setSearchMode(Mcm0020uConstants.MODE_DTSLOG);
        form.setLogDateFrom(LocalDate.now().minusDays(1)
                .format(DateTimeFormatter.ofPattern("yyyy/MM/dd")));
        form.setLogDateTo("");

        model.addAttribute("screenTitle", getScreenTitle());
        model.addAttribute("screenId", getFunctionId());
        model.addAttribute("form", form);
        model.addAttribute("results", Collections.emptyList());
        model.addAttribute("resultCount", 0);

        // リンク列表示制御
        // 【変換元】Mcm0020uScreen.vb — Mcm0020uScreen_Load()
        //   元コード: If Me.GetAuthorityData.AuthorityDivision_INSPECTION = ... Then
        //             Me.DataGridView.Columns("SUPPORT_ID_DTSLOG_Link").Visible = False
        boolean showLink = hasUpdateAuthority();
        model.addAttribute("showLink", showLink);
        // DTSログモード初期表示ではリンクは常に非表示
        model.addAttribute("linkEnabled", false);

        return "mcm0020u/index";
    }

    // ========================================================================
    // 検索ボタン
    // 【変換元】Mcm0020uScreen.vb — SearchButton_Click()
    //   元コード: If Me.DtsMstRenkeiLogRadioButton.Checked Then
    //                 Me.DaoContainer.Fill(dtslogSQL)
    //             Else
    //                 Me.DaoContainer.Fill(genzaijotaiSQL)
    //             End If
    // ========================================================================
    @PostMapping("/search")
    public String search(@ModelAttribute Mcm0020uForm form,
                         Model model,
                         HttpSession session) {

        // 抽出期間の入力チェック（DTSマスタ連携ログ検索時のみ対象）
        List<String> errors = validateSearchForm(form);
        if (!errors.isEmpty()) {
            model.addAttribute("screenTitle", getScreenTitle());
            model.addAttribute("screenId", getFunctionId());
            model.addAttribute("form", form);
            model.addAttribute("results", Collections.emptyList());
            model.addAttribute("resultCount", 0);
            model.addAttribute("showLink", hasUpdateAuthority());
            model.addAttribute("linkEnabled", false);
            model.addAttribute("errors", errors);
            return "mcm0020u/index";
        }

        List<DtslogSearchResultDto> results;
        boolean linkEnabled = false;

        if (Mcm0020uConstants.MODE_GENZAIJOTAI.equals(form.getSearchMode())) {
            // 現在状態検索
            results = service.searchGenzaijotai();

            // 【変換元】Mcm0020uScreen.vb — SearchButton_Click()
            //   元コード: Dim authorityMcm0021u = GetAuthorityDataFromScreenId("MCM0021U")
            //             If authorityMcm0021u.AuthorityDivision_UPDATE Then
            //                 Columns("SUPPORT_ID_DTSLOG_Link").Visible = True
            boolean hasUpdateForMcm0021u = hasUpdateAuthority();
            linkEnabled = hasUpdateForMcm0021u;
        } else {
            // DTSマスタ連携ログ検索
            results = service.searchDtslog(form.getLogDateFrom(), form.getLogDateTo());
            // DTSログモードではリンク列は常に非表示
            linkEnabled = false;
        }

        model.addAttribute("screenTitle", getScreenTitle());
        model.addAttribute("screenId", getFunctionId());
        model.addAttribute("form", form);
        model.addAttribute("results", results);
        model.addAttribute("resultCount", results.size());
        model.addAttribute("showLink", hasUpdateAuthority());
        model.addAttribute("linkEnabled", linkEnabled);

        // 検索結果が0件の場合、利用者へ「該当データなし」を明示するメッセージを表示する
        // 【期待仕様】異常終了と区別するため、正常終了かつ0件である旨をメッセージ表示する
        if (results.isEmpty()) {
            model.addAttribute("message", Mcm0020uConstants.MSG_NO_SEARCH_RESULT);
        }

        return "mcm0020u/index";
    }

    // ========================================================================
    // Excel出力ボタン
    // 【変換元】Mcm0020uScreen.vb — ErrorListOutputButton_Click()
    //   元コード: Dim mcm0001p As New Mcm0001pScreen
    //             mcm0001p.SetReportData(Me.DataGridView)
    //             mcm0001p.ShowDialog()
    // ========================================================================
    @PostMapping("/excel")
    public String exportExcel(@ModelAttribute Mcm0020uForm form,
                              RedirectAttributes ra) {
        // TODO: MCM0001P帳票連携 — DataGridViewの検索結果をMcm0001pExcelServiceで出力
        //       現状はMCM0001P帳票がすでに変換済みのため、
        //       検索結果をセッションに保持してMCM0001Pを呼び出す形で実装予定
        ra.addFlashAttribute("message", "Excel出力機能は現在実装中です");
        return "redirect:/mcm0020u";
    }

    // ========================================================================
    // 抽出期間バリデーション
    //   ・GENZAIJOTAI（現在状態）モードは日付項目が無効化されているため対象外
    //   ・DTSLOGモードはログ出力日（FROM/TO）の入力を必須とする（現行MCM準拠）
    //     → 未入力のまま検索させない（全件検索の抑止）
    //   ・"yyyy/MM/dd" 形式かつ暦上実在する日付のみを許可（STRICT判定）
    // ========================================================================
    private List<String> validateSearchForm(Mcm0020uForm form) {
        List<String> errors = new ArrayList<>();

        if (Mcm0020uConstants.MODE_DTSLOG.equals(form.getSearchMode())) {

            // 必須チェック（FROM・TOともに未入力の場合は検索処理を実行しない）
            if (isBlank(form.getLogDateFrom()) && isBlank(form.getLogDateTo())) {
                errors.add(Mcm0020uConstants.MSG_LOG_DATE_REQUIRED);
                // 必須未充足の時点で書式チェックは行わない（メッセージ重複を防ぐ）
                return errors;
            }

            if (!isValidLogDate(form.getLogDateFrom())) {
                errors.add("ログ出力日（開始）は yyyy/MM/dd 形式で正しい日付を入力してください。");
            }
            if (!isValidLogDate(form.getLogDateTo())) {
                errors.add("ログ出力日（終了）は yyyy/MM/dd 形式で正しい日付を入力してください。");
            }
        }
        return errors;
    }

    private boolean isBlank(String value) {
        return value == null || value.trim().isEmpty();
    }

    private boolean isValidLogDate(String value) {
        if (value == null || value.isEmpty()) {
            return true;
        }
        try {
            LocalDate.parse(value, LOG_DATE_FORMATTER);
            return true;
        } catch (DateTimeParseException e) {
            return false;
        }
    }

    // #298: ログイン時のSpring SecurityのROLE_UPDATEを共通基底クラスで判定する。
    // userInfoという未設定のセッション属性の有無で判定しない。
}
