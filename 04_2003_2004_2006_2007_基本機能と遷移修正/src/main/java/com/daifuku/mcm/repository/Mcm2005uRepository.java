package com.daifuku.mcm.repository;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Repository;

import com.daifuku.mcm.form.Mcm2005uForm;

/**
 * 【変換元】Mcm2005uDataSet.Designer.vb / Mcm2005uTabControl.vb
 * MCM2005U 店舗見積ブランド詳細設定リポジトリ
 *
 * SQL Server 構文使用。全テーブル名に MCM. スキーマ付与。
 * Excel帳票出力・SP_UMストアドプロシージャはWeb版省略。
 */
@Repository
public class Mcm2005uRepository {

    @Autowired
    private JdbcTemplate jdbc;

    private static final DateTimeFormatter DT_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");

    // ===================================================================
    // エントリポイント補完クエリ
    // 【変換元】Mcm2005uDelivery.vb — UmMitsumoriID / UmKihonBrandID プロパティ参照
    // ===================================================================

    /**
     * umKihonBrandId のみ指定時: 最初の期間の umMitsumoriId を取得する
     */
    public BigDecimal findMitsumoriIdByBrandId(BigDecimal umKihonBrandId) {
        return jdbc.queryForObject(
            "SELECT TOP 1 M.UM_MITSUMORI_ID " +
            "FROM MCM.MCM_UM_MITSUMORI M " +
            "INNER JOIN MCM.MCM_UM_KIHON_BRAND B ON B.UM_KIHON_MITSUMORI_ID = M.UM_KIHON_MITSUMORI_ID " +
            "WHERE B.UM_KIHON_BRAND_ID = ? " +
            "ORDER BY M.KAISI_DT",
            BigDecimal.class, umKihonBrandId);
    }

    /**
     * umMitsumoriId のみ指定時: 最初のブランドの umKihonBrandId を取得する
     */
    public BigDecimal findBrandIdByMitsumoriId(BigDecimal umMitsumoriId) {
        return jdbc.queryForObject(
            "SELECT TOP 1 B.UM_KIHON_BRAND_ID " +
            "FROM MCM.MCM_UM_KIHON_BRAND B " +
            "INNER JOIN MCM.MCM_UM_MITSUMORI M ON M.UM_KIHON_MITSUMORI_ID = B.UM_KIHON_MITSUMORI_ID " +
            "WHERE M.UM_MITSUMORI_ID = ? " +
            "ORDER BY B.UM_KIHON_BRAND_ID",
            BigDecimal.class, umMitsumoriId);
    }

    // ===================================================================
    // ブランド基本情報取得 (MCM_UM_KIHON_BRAND)
    // 【変換元】MCM_UM_KIHON_BRANDTableAdapter.FillBrand
    // ===================================================================

    public void fillKihonBrand(Mcm2005uForm form, BigDecimal umKihonBrandId) {
        List<Mcm2005uForm> list = jdbc.query(
            "SELECT B.UM_KIHON_BRAND_ID, B.UM_KIHON_MITSUMORI_ID, B.BRAND_NK, B.BRANDSYOSAI_NK, " +
            "       B.KEIYAKUJIKANTAI, B.HARDHOSYU_KIN, B.HOSEISOFTHOSHU_KIN " +
            "FROM MCM.MCM_UM_KIHON_BRAND B " +
            "WHERE B.UM_KIHON_BRAND_ID = ?",
            new Object[]{umKihonBrandId},
            (rs, rowNum) -> {
                Mcm2005uForm tmp = new Mcm2005uForm();
                tmp.setUmKihonBrandId(rs.getBigDecimal("UM_KIHON_BRAND_ID"));
                tmp.setUmKihonMitsumoriId(rs.getBigDecimal("UM_KIHON_MITSUMORI_ID"));
                tmp.setBrandNk(rs.getString("BRAND_NK"));
                tmp.setBrandsyosaiNk(rs.getString("BRANDSYOSAI_NK"));
                tmp.setKeiyakujikantai(rs.getString("KEIYAKUJIKANTAI"));
                tmp.setHardhosyuKinKihon(rs.getBigDecimal("HARDHOSYU_KIN"));
                tmp.setHoseisofthosyuKin(rs.getBigDecimal("HOSEISOFTHOSHU_KIN"));
                return tmp;
            });
        if (!list.isEmpty()) {
            Mcm2005uForm src = list.get(0);
            form.setUmKihonBrandId(src.getUmKihonBrandId());
            form.setUmKihonMitsumoriId(src.getUmKihonMitsumoriId());
            form.setBrandNk(src.getBrandNk());
            form.setBrandsyosaiNk(src.getBrandsyosaiNk());
            form.setKeiyakujikantai(src.getKeiyakujikantai());
            form.setHardhosyuKinKihon(src.getHardhosyuKinKihon());
            form.setHoseisofthosyuKin(src.getHoseisofthosyuKin());
        }
    }

    // ===================================================================
    // 期間情報取得 (MCM_UM_MITSUMORI)
    // 【変換元】Mcm2005uDelivery.vb — UmMitsumoriID 参照
    // ===================================================================

    public void fillMitsumori(Mcm2005uForm form, BigDecimal umMitsumoriId) {
        List<String[]> list = jdbc.query(
            "SELECT UM_MITSUMORI_ID, KAISI_DT, SYURYO_DT " +
            "FROM MCM.MCM_UM_MITSUMORI " +
            "WHERE UM_MITSUMORI_ID = ?",
            new Object[]{umMitsumoriId},
            (rs, rowNum) -> {
                BigDecimal mId = rs.getBigDecimal("UM_MITSUMORI_ID");
                LocalDateTime kaisi = rs.getObject("KAISI_DT", LocalDateTime.class);
                LocalDateTime syuryo = rs.getObject("SYURYO_DT", LocalDateTime.class);
                return new String[]{
                    mId != null ? mId.toPlainString() : null,
                    kaisi != null ? DT_FMT.format(kaisi) : "",
                    syuryo != null ? DT_FMT.format(syuryo) : ""
                };
            });
        if (!list.isEmpty()) {
            String[] row = list.get(0);
            form.setUmMitsumoriId(row[0] != null ? new java.math.BigDecimal(row[0]) : null);
            form.setKaisiDt(row[1]);
            form.setSyuryoDt(row[2]);
        }
    }

    // ===================================================================
    // 見積状態取得 (MCM_UM_KIHON_MITSUMORI)
    // 【変換元】Mcm2005uTabControl.vb — jotaiValue / shoninJotaiValue
    // ===================================================================

    public void fillJotai(Mcm2005uForm form, BigDecimal umKihonMitsumoriId) {
        List<String[]> list = jdbc.query(
            "SELECT JOTAI, SYOUNIN_JOTAI " +
            "FROM MCM.MCM_UM_KIHON_MITSUMORI " +
            "WHERE UM_KIHON_MITSUMORI_ID = ?",
            new Object[]{umKihonMitsumoriId},
            (rs, rowNum) -> new String[]{rs.getString("JOTAI"), rs.getString("SYOUNIN_JOTAI")});
        if (!list.isEmpty()) {
            String[] row = list.get(0);
            form.setJotai(row[0]);
            form.setSyouninJotai(row[1]);
        }
    }

    // ===================================================================
    // 見積条件取得 (MCM_UM_BRAND)
    // 【変換元】MCM_UM_BRANDTableAdapter.Fill (umMitsumoriId + umKihonBrandId)
    // ===================================================================

    public void fillBrand(Mcm2005uForm form, BigDecimal umMitsumoriId, BigDecimal umKihonBrandId) {
        List<Mcm2005uForm> rows = jdbc.query(
            "SELECT MITSUMORICHUKI, HARDHOSYU_KIN, CHOSEI_KIN, HOSHU_KIN, BIKO " +
            "FROM MCM.MCM_UM_BRAND " +
            "WHERE UM_MITSUMORI_ID = ? AND UM_KIHON_BRAND_ID = ?",
            new Object[]{umMitsumoriId, umKihonBrandId},
            (rs, rowNum) -> {
                Mcm2005uForm tmp = new Mcm2005uForm();
                tmp.setMitsumorichuki(rs.getString("MITSUMORICHUKI"));
                tmp.setHardhosyuKin(rs.getBigDecimal("HARDHOSYU_KIN"));
                tmp.setChoseiKin(rs.getBigDecimal("CHOSEI_KIN"));
                tmp.setHoshuKin(rs.getBigDecimal("HOSHU_KIN"));
                tmp.setBiko(rs.getString("BIKO"));
                return tmp;
            });
        if (!rows.isEmpty()) {
            Mcm2005uForm src = rows.get(0);
            form.setMitsumorichuki(src.getMitsumorichuki());
            form.setHardhosyuKin(src.getHardhosyuKin());
            form.setChoseiKin(src.getChoseiKin());
            form.setHoshuKin(src.getHoshuKin());
            form.setBiko(src.getBiko());
        }
    }

    // ===================================================================
    // 機器単価グリッド取得 (MCM_UM_TANKA JOIN MCM_UM_KIKIKOSEI LEFT JOIN MCM_UM_KIKIMEISAI)
    // 【変換元】MCM_UM_TANKATableAdapter.Fill (umMitsumoriId + umKihonBrandId)
    //          親子リレーション: MCM_UM_KIKIKOSEI → MCM_UM_TANKA
    // ===================================================================

    public List<Mcm2005uForm.TankaRowForm> findTankaRows(BigDecimal umMitsumoriId, BigDecimal umKihonBrandId) {
        return jdbc.query(
            "SELECT T.UM_TANKA_ID,K.UM_KIKIKOSEI_ID,ISNULL(V.SURYO_NM,M.SURYO_NM) AS SURYO_NM,T.HYOJUN_KIN, " +
            "ISNULL(ISNULL(V.SURYO_NM,M.SURYO_NM),0)*ISNULL(T.HYOJUN_KIN,0) AS HYOJUNSHOKEI_KIN, " +
            "ISNULL(ISNULL(V.SURYO_NM,M.SURYO_NM),0)*ISNULL(T.SIKIRI_KIN,0) AS SIKIRISHOKEI_KIN,K.CONTROLLER_FLG,K.SET_NM,K.HYOJIJUN,M.KIKIHINMEI_NK AS MEISAI_NM " +
            "FROM MCM.MCM_UM_TANKA T JOIN MCM.MCM_UM_KIKIMEISAI M ON M.UM_KIKIMEISAI_ID=T.UM_KIKIMEISAI_ID JOIN MCM.MCM_UM_KIKIKOSEI K ON K.UM_KIKIKOSEI_ID=M.UM_KIKIKOSEI_ID " +
            "LEFT JOIN (SELECT SUM(MAF.SURYO_NM) AS SURYO_NM,UMF.UM_TANKA_ID,TMD.SURYO_NM AS SOSU_NM " +
            "FROM MCM.MCM_UM_TANKA UMF JOIN MCM.MCM_UM_KIKIMEISAI UMD ON UMD.UM_KIKIMEISAI_ID=UMF.UM_KIKIMEISAI_ID " +
            "JOIN MCM.MCM_UM_KOTAIMEISAI UME ON UME.UM_KIKIMEISAI_ID=UMD.UM_KIKIMEISAI_ID " +
            "JOIN MCM.MCM_TM_TANKA TMF ON TMF.TM_TANKA_ID=UMF.TM_TANKA_ID " +
            "JOIN MCM.MCM_TM_KIKIMEISAI TMD ON TMD.TM_KIKIMEISAI_ID=TMF.TM_KIKIMEISAI_ID AND TMD.KIKIMEISAI_ID=UMD.KIKIMEISAI_ID " +
            "JOIN MCM.MCM_TM_KOTAIMEISAI TME ON TME.TM_KIKIMEISAI_ID=TMD.TM_KIKIMEISAI_ID AND TME.KOTAIKANRI_ID=UME.KOTAIKANRI_ID " +
            "JOIN MCM.MCM_MA_KIKIMEISAI MAF ON MAF.KIKIMEISAI_ID=TMD.KIKIMEISAI_ID GROUP BY UMF.UM_TANKA_ID,TMD.SURYO_NM) V ON V.UM_TANKA_ID=T.UM_TANKA_ID " +
            "WHERE T.UM_MITSUMORI_ID=? AND K.UM_KIHON_BRAND_ID=? ORDER BY K.HYOJIJUN,M.HYOJIJUN,T.UM_TANKA_ID",
            new Object[]{umMitsumoriId, umKihonBrandId},
            (rs, rowNum) -> {
                Mcm2005uForm.TankaRowForm row = new Mcm2005uForm.TankaRowForm();
                row.setUmTankaId(rs.getBigDecimal("UM_TANKA_ID"));
                row.setUmKikoseiId(rs.getBigDecimal("UM_KIKIKOSEI_ID"));
                row.setSetNm(rs.getString("SET_NM"));
                row.setMeisaiNm(rs.getString("MEISAI_NM"));
                row.setSuryoNm(rs.getString("SURYO_NM"));
                row.setHyojunKin(rs.getBigDecimal("HYOJUN_KIN"));
                row.setHyojunshokeiKin(rs.getBigDecimal("HYOJUNSHOKEI_KIN"));
                row.setSikirishokeiKin(rs.getBigDecimal("SIKIRISHOKEI_KIN"));
                row.setControllerFlg(rs.getString("CONTROLLER_FLG"));
                return row;
            });
    }

    // ===================================================================
    // 見積条件保存 MCM_UM_BRAND — MERGE INTO (UPSERT)
    // 【変換元】UpdateButtonTabNaiyo → MyBase.UpdateAll(MCM_UM_BRAND)
    //
    // HOSHU_KIN = HARDHOSYU_KIN(MCM_UM_BRAND) + HOSEISOFTHOSHU_KIN(MCM_UM_KIHON_BRAND) - CHOSEI_KIN
    // 新規INSERT時は HARDHOSYU_KIN の初期値として MCM_UM_KIHON_BRAND.HARDHOSYU_KIN を使用する。
    // ===================================================================

    public void mergeBrand(Mcm2005uForm form, String loginUser) {
        jdbc.update(
            "MERGE INTO MCM.MCM_UM_BRAND AS target " +
            "USING (SELECT ? AS UM_KIHON_BRAND_ID, ? AS UM_MITSUMORI_ID) AS source " +
            "ON target.UM_KIHON_BRAND_ID = source.UM_KIHON_BRAND_ID " +
            "   AND target.UM_MITSUMORI_ID = source.UM_MITSUMORI_ID " +
            "WHEN MATCHED THEN " +
            "    UPDATE SET MITSUMORICHUKI=?, CHOSEI_KIN=?, HOSHU_KIN=?, BIKO=?, " +
            "               LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? " +
            "WHEN NOT MATCHED THEN " +
            "    INSERT (UM_KIHON_BRAND_ID, UM_MITSUMORI_ID, MITSUMORICHUKI, " +
            "            HARDHOSYU_KIN, CHOSEI_KIN, HOSHU_KIN, BIKO, " +
            "            CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY) " +
            "    VALUES (?, ?, ?, ?, ?, ?, ?, GETDATE(), ?, GETDATE(), ?);",
            // USING 句
            form.getUmKihonBrandId(), form.getUmMitsumoriId(),
            // WHEN MATCHED UPDATE
            form.getMitsumorichuki(), form.getChoseiKin(), form.getHoshuKin(), form.getBiko(),
            loginUser,
            // WHEN NOT MATCHED INSERT
            form.getUmKihonBrandId(), form.getUmMitsumoriId(), form.getMitsumorichuki(),
            form.getHardhosyuKinKihon(),  // 初期値は KIHON_BRAND のハード保守費
            form.getChoseiKin(), form.getHoshuKin(), form.getBiko(),
            loginUser, loginUser);
    }

    // ===================================================================
    // 機器単価更新 (MCM_UM_TANKA)
    // 【変換元】UpdateButtonTabNaiyo → MyBase.UpdateAll(MCM_UM_TANKA)
    //
    // HYOJUNSHOKEI_KIN はサービス層で計算 (HYOJUN_KIN × SURYO_NM の整数部)
    // ===================================================================

    public void updateTanka(Mcm2005uForm.TankaRowForm row, String loginUser) {
        jdbc.update(
            "UPDATE MCM.MCM_UM_TANKA " +
            "SET HYOJUN_KIN=?, HYOJUNSHOKEI_KIN=?, SIKIRISHOKEI_KIN=?, " +
            "    LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? " +
            "WHERE UM_TANKA_ID=?",
            row.getHyojunKin(), row.getHyojunshokeiKin(), row.getSikirishokeiKin(),
            loginUser, row.getUmTankaId());
    }
}
