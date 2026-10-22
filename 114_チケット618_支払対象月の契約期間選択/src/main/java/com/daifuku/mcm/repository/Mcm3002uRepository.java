package com.daifuku.mcm.repository;

import java.sql.Date;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Repository;

import com.daifuku.mcm.form.Mcm3002uForm.RowForm;

/**
 * 【変換元】Mcm3002u_MCM_TK_SIHARAI_VTableAdapter.xml
 * MCM3002U 支払い明細出力指示リポジトリ
 *
 * Oracle → SQL Server 変換ポイント:
 *   ROWNUM <= 1         → TOP(1) サブクエリ
 *   TRUNC(TSUKI, 'DD')  → CAST(TSUKI AS DATE)（TSUKI は DATE型、月初日で格納）
 *   MAL.YUKO_FLG = '0'  → MAL.YUKO_FLG = 0（SQL Server は数値比較）
 *   MCM_FN_SIHARAI(...) → 元関数の取得条件＋Java側の月配分計算
 *
 * 前提:
 *   MCM_TU_KEIYAKU_V と関連テーブルが存在すること。
 *   支払ビューの期間切替条件は画面専用SQLに明示し、配備先ビューの定義に依存させない。
 */
@Repository
public class Mcm3002uRepository {

    @Autowired
    private JdbcTemplate jdbc;

    private static final DateTimeFormatter TSUKI_FMT = DateTimeFormatter.ofPattern("yyyy/MM");

    /**
     * 正式VBの MCM_TK_SIHARAI_V + Mcm3002u の抽出条件。
     * 開始月が支払対象月以前の有効期間から、最大開始日を採用する。
     * 終了日による独自の足切りや、期間ID最大値への置換は行わない。
     * DISTINCTにはVBと同じ業務キーを残し、別の製番・支払明細を任意に統合しない。
     */
    private static final String SQL_SEARCH = """
        WITH PAYMENT_ROWS AS (
            SELECT DISTINCT
                   ISNULL(MAA.NONYUSAKI_ID, TKB.NONYUSAKI_ID) AS NONYUSAKI_ID,
                   ISNULL(MAB.PLANT_ID, TKB.PLANT_ID) AS PLANT_ID,
                   TKA.TK_KEIYAKU_ID, TKB.TK_KIKAN_ID,
                   TKJ.TK_SIHARAI_ID, TKK.TK_SIHARAIMEISAI_ID, TUV.UK_KEIYAKU_ID,
                   ISNULL(MAA.NONYUSAKI_CD, TKB.NONYUSAKI_CD) AS NONYUSAKI_CD,
                   ISNULL(MAB.SUPPORT_ID, TKB.SUPPORT_ID) AS SUPPORT_ID,
                   ISNULL(MAA.NONYUSAKI_NK, TKB.NONYUSAKI_NK) AS NONYUSAKI_NK,
                   ISNULL(MAB.PLANT_NK, TKB.PLANT_NK) AS PLANT_NK,
                   ISNULL(MAR.MEISHO4_NK, TUV.IRAIMEISHO4_NK) AS TENPO_NK,
                   TKA.KEIYAKU_NO,
                   ISNULL(MAK.TORIHIKISAKI_ID, TKB.TORIHIKISAKI_ID) AS TORIHIKISAKI_ID,
                   ISNULL(MAK.TORIHIKISAKI_CD, TKB.TORIHIKISAKI_CD) AS TORIHIKISAKI_CD,
                   ISNULL(MAK.TORIHIKISAKI_NK, TKB.TORIHIKISAKI_NK) AS TORIHIKISAKI_NK,
                   MAK.TORIHIKISAKI_KN,
                   CASE ISNULL(TKJ.KAISU, 0)
                       WHEN 1 THEN N'年払' WHEN 0 THEN N'不明' ELSE N'月払' END AS SIHARAI,
                   TKJ.KAISU, TKJ.BIKO, TKK.TSUKI, TKB.KAISI_DT, TKB.SYURYO_DT,
                   CASE TKA.JOTAI WHEN '2' THEN N'契約' WHEN '4' THEN N'解約' END AS JOTAI,
                   TKA.KAIYAKU_DT,
                   CASE WHEN TKK.TSUKI < ISNULL(UKA.KAIYAKU_DT, CAST('9999-12-31' AS DATE))
                       THEN N'' ELSE N'* ' END + ISNULL(UKI.HARD_SEIBAN, N'') AS HARD_SEIBAN
              FROM MCM.MCM_TU_KEIYAKU_V TUV
              JOIN MCM.MCM_TK_KEIYAKU TKA ON TKA.TK_KEIYAKU_ID = TUV.TK_KEIYAKU_ID
              JOIN MCM.MCM_TK_KIKAN TKB ON TKB.TK_KIKAN_ID = TUV.TK_KIKAN_ID
              JOIN MCM.MCM_TK_SIHARAI TKJ ON TKJ.TK_KIKAN_ID = TKB.TK_KIKAN_ID
              JOIN MCM.MCM_TK_SIHARAIMEISAI TKK ON TKK.TK_SIHARAI_ID = TKJ.TK_SIHARAI_ID
              LEFT JOIN MCM.MCM_MA_TORIHIKISAKI MAK ON MAK.TORIHIKISAKI_ID = TKB.TORIHIKISAKI_ID
              LEFT JOIN MCM.MCM_MA_PLANT MAB ON MAB.PLANT_ID = TKB.PLANT_ID
              LEFT JOIN MCM.MCM_MA_NONYUSAKI MAA ON MAA.NONYUSAKI_ID = MAB.NONYUSAKI_ID
              LEFT JOIN MCM.MCM_UK_KEIYAKU UKA ON UKA.UK_KEIYAKU_ID = TUV.UK_KEIYAKU_ID
              LEFT JOIN MCM.MCM_UK_SEIBAN UKI ON UKI.UK_KEIYAKU_ID = TUV.UK_KEIYAKU_ID
                   AND TKK.TSUKI BETWEEN UKI.KAISI_DT
                       AND ISNULL(UKI.SYURYO_DT, CAST('9999-12-31' AS DATE))
              LEFT JOIN MCM.MCM_MA_TENPO MAR ON MAR.TENPO_ID = TUV.IRAITENPO_ID
             WHERE TKA.JOTAI IN ('2', '4')
               AND CAST(TKK.TSUKI AS DATE) = ?
               AND TKB.YUKO_FLG <= 0
               AND TKK.ON_FLG = 1 AND TKK.OFF_FLG = 0
               AND TKK.TSUKI <= ISNULL(TKA.KAIYAKU_DT, CAST('9999-12-31' AS DATE))
               AND TKB.KAISI_DT = (
                   SELECT MAX(SB.KAISI_DT)
                     FROM MCM.MCM_TK_KIKAN SB
                    WHERE SB.TK_KEIYAKU_ID = TKA.TK_KEIYAKU_ID
                      AND SB.YUKO_FLG <= 0
                      AND DATEFROMPARTS(YEAR(SB.KAISI_DT), MONTH(SB.KAISI_DT), 1)
                          <= CAST(TKK.TSUKI AS DATE))
               AND UKI.KAKUNIN_KBN = '1'
        )
        SELECT TKV.TK_KIKAN_ID, TKV.KAISI_DT, TKV.SYURYO_DT, TKV.TORIHIKISAKI_ID,
               TKV.NONYUSAKI_CD, TKV.SUPPORT_ID, TKV.JOTAI, TKV.NONYUSAKI_NK,
               TKV.TENPO_NK, TKV.KEIYAKU_NO, TKV.TORIHIKISAKI_CD, TKV.TORIHIKISAKI_NK,
               (SELECT TOP(1) MAL.TORISYUTANTOSYA_NK
                  FROM MCM.MCM_MA_TORIMADOGUCHI MAL
                 WHERE MAL.TORIHIKISAKI_ID = TKV.TORIHIKISAKI_ID
                   AND MAL.MADOGUCHI_KBN = '2' AND MAL.YUKO_FLG = 0) AS TORISYUTANTOSYA_NK,
               TKV.SIHARAI, TKV.KAISU, TKV.BIKO, TKV.TSUKI, TKV.HARD_SEIBAN
          FROM PAYMENT_ROWS TKV
         WHERE TKV.SIHARAI = ?
         ORDER BY TKV.TORIHIKISAKI_ID, TKV.KEIYAKU_NO, TKV.NONYUSAKI_CD, TKV.SUPPORT_ID
        """;

    // ===================================================================
    // 支払い明細検索
    // 【変換元】Mcm3002u_MCM_TK_SIHARAI_VTableAdapter.xml SELECT
    //   パラメーター:
    //     shiharaiDate : 支払い月の月初日（YYYYMM → LocalDate.of(y, m, 1)）
    //     siharaiKbn   : 支払区分（"年払" or "月払"）
    // ===================================================================

    public List<RowForm> search(LocalDate shiharaiDate, String siharaiKbn) {
        return jdbc.query(SQL_SEARCH, new Object[]{Date.valueOf(shiharaiDate), siharaiKbn},
            (rs, rowNum) -> {
                RowForm row = new RowForm();
                row.setTkKikanId(rs.getBigDecimal("TK_KIKAN_ID"));
                Date start = rs.getDate("KAISI_DT");
                Date end = rs.getDate("SYURYO_DT");
                row.setKaishiDt(start == null ? null : start.toLocalDate());
                row.setSyuryoDt(end == null ? null : end.toLocalDate());
                row.setTorihikisakiId(rs.getBigDecimal("TORIHIKISAKI_ID"));
                row.setNonyusakiCd(rs.getString("NONYUSAKI_CD"));
                row.setSupportId(rs.getString("SUPPORT_ID"));
                row.setJotai(rs.getString("JOTAI"));
                row.setNonyusakiNk(rs.getString("NONYUSAKI_NK"));
                row.setTenpoNk(rs.getString("TENPO_NK"));
                row.setKeiyakuNo(rs.getString("KEIYAKU_NO"));
                row.setTorihikisakiCd(rs.getString("TORIHIKISAKI_CD"));
                row.setTorihikisakiNk(rs.getString("TORIHIKISAKI_NK"));
                row.setTorisyutantosyaNk(rs.getString("TORISYUTANTOSYA_NK"));
                row.setSiharai(rs.getString("SIHARAI"));
                row.setKaisu(rs.getBigDecimal("KAISU"));
                row.setBiko(rs.getString("BIKO"));
                // TSUKI は DATE 型。表示用に YYYY/MM 形式の文字列に変換
                java.sql.Date tsuki = rs.getDate("TSUKI");
                row.setTsuki(tsuki != null ? tsuki.toLocalDate().format(TSUKI_FMT) : "");
                row.setHardSeiban(rs.getString("HARD_SEIBAN"));
                return row;
            });
    }
    // 支払区分・状態名はSQL_SEARCHでVBと同じ規則から生成する。
    // MCM_TK_SIHARAI_Vの文字列定義・期間選択に差異があっても、この画面では採用しない。
    /** MCM_FN_SIHARAI.LCUR_KANE。EXISTSで同じ見積期間の重複加算を避ける。 */
    public record RatePeriod(LocalDate start, LocalDate end, java.math.BigDecimal amount) {}

    public List<RatePeriod> findRatePeriods(java.math.BigDecimal tkKikanId) {
        return jdbc.query("""
            SELECT TMG.KAISI_DT, TMG.SYURYO_DT, TMG.SIKIRIGOKEI_KIN
              FROM MCM.MCM_TM_KIKAN TMG
             WHERE TMG.SIKIRIGOKEI_KIN > 0
               AND EXISTS (
                   SELECT 1 FROM MCM.MCM_TK_KIKIKOSEI TKC
                   JOIN MCM.MCM_TK_KIKIMEISAI TKD ON TKD.TK_KIKIKOSEI_ID = TKC.TK_KIKIKOSEI_ID
                   JOIN MCM.MCM_TK_TANKA TKF ON TKF.TK_KIKIMEISAI_ID = TKD.TK_KIKIMEISAI_ID
                   JOIN MCM.MCM_TM_TANKA TMF ON TMF.TM_TANKA_ID = TKF.TM_TANKA_ID
                   WHERE TMF.TM_KIKAN_ID = TMG.TM_KIKAN_ID AND TKC.TK_KIKAN_ID = ?)
            """, (rs, n) -> {
                Date start = rs.getDate("KAISI_DT"), end = rs.getDate("SYURYO_DT");
                return new RatePeriod(start == null ? null : start.toLocalDate(),
                        end == null ? null : end.toLocalDate(), rs.getBigDecimal("SIKIRIGOKEI_KIN"));
            }, tkKikanId);
    }

    /** MCM_FN_SIHARAI.LCUR_TUKI。終了境界は含めない。 */
    public List<LocalDate> findPaymentDates(java.math.BigDecimal tkKikanId, LocalDate start, LocalDate end) {
        return jdbc.query("""
            SELECT TKK.TSUKI FROM MCM.MCM_TK_SIHARAI TKJ
            JOIN MCM.MCM_TK_SIHARAIMEISAI TKK ON TKK.TK_SIHARAI_ID = TKJ.TK_SIHARAI_ID
            WHERE TKJ.TK_KIKAN_ID = ? AND TKK.ON_FLG = 1 AND TKK.OFF_FLG = 0
              AND TKK.TSUKI >= ? AND TKK.TSUKI < ?
            """, (rs, n) -> rs.getDate("TSUKI").toLocalDate(), tkKikanId, Date.valueOf(start), Date.valueOf(end));
    }
}
