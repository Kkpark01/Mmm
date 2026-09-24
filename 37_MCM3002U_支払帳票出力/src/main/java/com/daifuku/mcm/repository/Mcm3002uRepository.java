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
 *   MCM.MCM_TK_SIHARAI_V が SQL Server に VIEW として存在すること。
 *   Oracle 版の同名ビューと同等のカラムを持つこと。
 */
@Repository
public class Mcm3002uRepository {

    @Autowired
    private JdbcTemplate jdbc;

    private static final DateTimeFormatter TSUKI_FMT = DateTimeFormatter.ofPattern("yyyy/MM");

    // ===================================================================
    // 支払い明細検索
    // 【変換元】Mcm3002u_MCM_TK_SIHARAI_VTableAdapter.xml SELECT
    //   パラメーター:
    //     shiharaiDate : 支払い月の月初日（YYYYMM → LocalDate.of(y, m, 1)）
    //     siharaiKbn   : 支払区分（"年払" or "月払"）
    // ===================================================================

    public List<RowForm> search(LocalDate shiharaiDate, String siharaiKbn) {
        String sql =
            "SELECT " +
            "    TKV.TK_KIKAN_ID, " +
            "    TKV.KAISI_DT, TKV.SYURYO_DT, " +
            "    TKV.TORIHIKISAKI_ID, " +
            "    TKV.NONYUSAKI_CD, " +
            "    TKV.SUPPORT_ID, " +
            "    TKV.JOTAI, " +
            "    TKV.NONYUSAKI_NK, " +
            "    TKV.TENPO_NK, " +
            "    TKV.KEIYAKU_NO, " +
            "    TKV.TORIHIKISAKI_CD, " +
            "    TKV.TORIHIKISAKI_NK, " +
            "    (SELECT TOP(1) MAL.TORISYUTANTOSYA_NK " +
            "       FROM MCM.MCM_MA_TORIMADOGUCHI MAL " +
            "      WHERE MAL.TORIHIKISAKI_ID = TKV.TORIHIKISAKI_ID " +
            "        AND MAL.MADOGUCHI_KBN   = '2' " +
            "        AND MAL.YUKO_FLG        = 0) AS TORISYUTANTOSYA_NK, " +
            "    TKV.SIHARAI, " +
            "    TKV.KAISU, " +
            "    TKV.BIKO, " +
            "    TKV.TSUKI, " +
            "    TKV.HARD_SEIBAN " +
            "FROM MCM.MCM_TK_SIHARAI_V TKV " +
            "WHERE CAST(TKV.TSUKI AS DATE) = ? " +
            "  AND TKV.SIHARAI = ? " +
            "ORDER BY TKV.TORIHIKISAKI_ID ASC, TKV.KEIYAKU_NO ASC, " +
            "         TKV.NONYUSAKI_CD ASC, TKV.SUPPORT_ID ASC";

        return jdbc.query(sql, new Object[]{Date.valueOf(shiharaiDate), siharaiKbn},
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
