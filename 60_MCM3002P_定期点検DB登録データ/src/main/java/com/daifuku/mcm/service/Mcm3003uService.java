package com.daifuku.mcm.service;

import com.daifuku.mcm.form.Mcm3003uForm;
import com.daifuku.mcm.form.Mcm3003uForm.RowForm;
import com.daifuku.mcm.repository.Mcm3003uRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.format.ResolverStyle;
import java.util.List;

/**
 * 【変換元】Mcm3003uScreen.vb Search() / Delivery_Change() / GetSyutsuryokuKubun()
 * MCM3003U 保守点検出力指示サービス
 *
 * VB版の入力チェック順序・メッセージをそのまま移植する。
 *   1. 点検対象月（開始 or 終了）が未入力          → FWM_0001E「点検対象月は必ず入力してください」
 *   2. 点検対象月（開始）が YYYY/MM 形式でない     → FWM_0006E「点検対象月はYYYY/MMで書式を入力してください」
 *   3. 点検対象月（終了）が YYYY/MM 形式でない     → 同上
 *   4. 差分出力かつ 出力対象日付が開始・終了とも未入力 → FWM_0001E「出力対象日付は必ず入力してください」
 *   5. 差分出力かつ 入力済みの出力対象日付が YYYY/MM/DD 形式でない
 *                                                  → FWM_0006E「出力対象日付はYYYY/MM/DDで書式を入力してください」
 *
 * 注意: VB版に「開始 &gt; 終了」のチェックは存在しない（該当データ0件として MSG_0002 になる）。
 *       Web版でも独自チェックは追加しない。
 */
@Service
public class Mcm3003uService {

    /** 【変換元】CPConstant.FORMAT_YYYYMM */
    private static final DateTimeFormatter YYYYMM =
            DateTimeFormatter.ofPattern("uuuu/MM").withResolverStyle(ResolverStyle.STRICT);

    /** 【変換元】CPConstant.FORMAT_YYYYMMDD */
    private static final DateTimeFormatter YYYYMMDD =
            DateTimeFormatter.ofPattern("uuuu/MM/dd").withResolverStyle(ResolverStyle.STRICT);

    /** 【変換元】Mcm3003uConstant.TENKEN_TAISYOZUKI */
    private static final String LABEL_TENKEN_TAISYOZUKI = "点検対象月";

    /** 【変換元】Mcm3003uConstant.TAISHO_DT / TAISYO_DT */
    private static final String LABEL_TAISYO_DT = "出力対象日付";

    /** 【変換元】Mcm3003uConstant.ZENKEN */
    private static final String SUFFIX_ZENKEN = "_全件";

    /** 【変換元】Mcm3003uConstant.SABUN */
    private static final String SUFFIX_SABUN = "_差分";

    @Autowired
    private Mcm3003uRepository repository;

    // 【移植】画面アクセス権限（機能ID=MCM3003U）は全画面共通の ScreenAuthorityInterceptor
    //   （SESSION_DENIED_FUNCTION_IDS 参照）で一元的にチェックされるため、本Service内での
    //   個別チェックは不要。

    /**
     * 出力対象データ検索。
     * 【変換元】Mcm3003uScreen.vb Search()
     */
    @Transactional(readOnly = true)
    public List<RowForm> searchForOutput(Mcm3003uForm form) {

        if (form.getSyutsuryokuKbn() != Mcm3003uForm.SYUTSURYOKU_KBN_ZENKEN
                && form.getSyutsuryokuKbn() != Mcm3003uForm.SYUTSURYOKU_KBN_SABUN) {
            throw new IllegalArgumentException("出力区分を選択してください。");
        }
        String kaishiText = trim(form.getFromTsuki());
        String shuryoText = trim(form.getToTsuki());

        // 必須チェック（VB: IsNull(kaishi) Or IsNull(shuryo)）
        if (isNull(kaishiText) || isNull(shuryoText)) {
            throw new IllegalArgumentException(required(LABEL_TENKEN_TAISYOZUKI));
        }

        // 点検対象月（開始月）: 日付チェック → 月初日へ変換
        LocalDate fromDate = parseMonth(kaishiText).atDay(1);
        // 点検対象月（終了月）: 日付チェック → 月末日へ変換
        LocalDate toDate = parseMonth(shuryoText).atEndOfMonth();

        LocalDate sabunFrom = null;
        LocalDate sabunTo = null;

        // 出力対象日付チェック（差分出力の場合のみ）
        if (form.isSabun()) {
            String sabunKaishiText = trim(form.getSabunFrom());
            String sabunShuryoText = trim(form.getSabunTo());

            // 必須チェック（VB: IsNull(開始) And IsNull(終了) → 両方未入力のときのみエラー）
            if (isNull(sabunKaishiText) && isNull(sabunShuryoText)) {
                throw new IllegalArgumentException(required(LABEL_TAISYO_DT));
            }
            if (!isNull(sabunKaishiText)) {
                sabunFrom = parseDate(sabunKaishiText);
            }
            if (!isNull(sabunShuryoText)) {
                sabunTo = parseDate(sabunShuryoText);
            }
        }

        return repository.search(fromDate, toDate, sabunFrom, sabunTo);
    }

    /**
     * 出力ファイル名の先頭部分を組み立てる。
     * 【変換元】Mcm3003uScreen.vb Delivery_Change()
     *   全件出力: 点検対象月(開始)＋"_"＋点検対象月(終了)＋"_全件"（スラッシュは除去）
     *   差分出力: 出力対象日付(開始)＋"_"＋出力対象日付(終了)＋"_差分"
     *             片方のみ入力されている場合はその値のみ（スラッシュは除去）
     */
    public String buildFileNk(Mcm3003uForm form) {
        if (!form.isSabun()) {
            return removeSlash(trim(form.getFromTsuki())) + "_"
                 + removeSlash(trim(form.getToTsuki())) + SUFFIX_ZENKEN;
        }
        String sabunKaishi = trim(form.getSabunFrom());
        String sabunShuryo = trim(form.getSabunTo());
        String fileNk;
        if (!isNull(sabunKaishi) && !isNull(sabunShuryo)) {
            fileNk = removeSlash(sabunKaishi) + "_" + removeSlash(sabunShuryo);
        } else if (!isNull(sabunKaishi)) {
            fileNk = removeSlash(sabunKaishi);
        } else if (!isNull(sabunShuryo)) {
            fileNk = removeSlash(sabunShuryo);
        } else {
            fileNk = "";
        }
        return fileNk + SUFFIX_SABUN;
    }

    // ===================================================================
    // 内部処理
    // ===================================================================

    /** 【変換元】CPCommonUtility.IsCheckDate(value, FORMAT_YYYYMM) */
    private YearMonth parseMonth(String value) {
        try {
            if (!value.matches("[0-9]{4}/[0-9]{2}")) throw new DateTimeParseException("format", value, 0);
            YearMonth month = YearMonth.parse(value, YYYYMM);
            if (month.getYear() < 1) throw new DateTimeParseException("year", value, 0);
            return month;
        } catch (DateTimeParseException e) {
            throw new IllegalArgumentException(format(LABEL_TENKEN_TAISYOZUKI, "YYYY/MM"));
        }
    }

    /** 【変換元】CPCommonUtility.IsCheckDate(value, FORMAT_YYYYMMDD) */
    private LocalDate parseDate(String value) {
        try {
            if (!value.matches("[0-9]{4}/[0-9]{2}/[0-9]{2}")) throw new DateTimeParseException("format", value, 0);
            LocalDate date = LocalDate.parse(value, YYYYMMDD);
            if (date.getYear() < 1) throw new DateTimeParseException("year", value, 0);
            return date;
        } catch (DateTimeParseException e) {
            throw new IllegalArgumentException(format(LABEL_TAISYO_DT, "YYYY/MM/DD"));
        }
    }

    /** 【変換元】FWM_0001E「{0}は必ず入力してください」 */
    private String required(String label) {
        return label + "は必ず入力してください。";
    }

    /** 【変換元】FWM_0006E「{0}は{1}で書式を入力してください」 */
    private String format(String label, String pattern) {
        return label + "は" + pattern + "で書式を入力してください。";
    }

    private static String trim(String value) {
        return value == null ? "" : value.trim();
    }

    /** 【変換元】CPCommonUtility.IsNull */
    private static boolean isNull(String value) {
        return value == null || value.isEmpty();
    }

    private static String removeSlash(String value) {
        return value.replace("/", "");
    }
}
