package com.daifuku.mcm.repository;

import com.daifuku.mcm.dto.Mcm2001uRowDto;
import com.daifuku.mcm.form.Mcm2001uForm;
import org.springframework.jdbc.core.namedparam.MapSqlParameterSource;
import org.springframework.jdbc.core.namedparam.NamedParameterJdbcTemplate;
import org.springframework.stereotype.Repository;

import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.List;
import java.util.Map;

/**
 * 【変換元】Mcm2001uDataSet.Designer.vb（TableAdapter群）
 *            + Mcm2001u_MCM_MA_NONYUSAKITableAdapter.xml
 *            + Mcm2001u_MCM_MA_PLANTTableAdapter.xml
 *            + Mcm2001u_MCM_TM_MITSUMORITableAdapter.xml
 *            + Mcm2001uScreen.vb - getUmKeiyakujikanId() / SENTEIButton_Click()内SQL
 *   MCM2001U カスタマー見積検索 JdbcRepository
 *   VB.NET → Java変換
 *
 *   動的SQL（5条件）× 3テーブル + ユーティリティクエリ2件
 *   Oracle → SQL Server変換済み
 */
@Repository
public class Mcm2001uJdbcRepository {

    private final NamedParameterJdbcTemplate jdbc;

    public Mcm2001uJdbcRepository(NamedParameterJdbcTemplate jdbc) {
        this.jdbc = jdbc;
    }

    public List<Map<String,Object>> findNonyusakiById(BigDecimal id) {
        return jdbc.queryForList("SELECT NONYUSAKI_ID, NONYUSAKI_CD, NONYUSAKI_NK, NONYUSAKIKANA_KN, NONYUSAKIEIMEI_EN, NONYUSAKIKOJO_NK, KYUNONYUSAKI_NK FROM MCM.MCM_MA_NONYUSAKI WHERE NONYUSAKI_ID=:id", new MapSqlParameterSource("id",id));
    }
    public List<Map<String,Object>> findPlantById(BigDecimal id) {
        return jdbc.queryForList("SELECT PLANT_ID, NONYUSAKI_ID, SUPPORT_ID, PLANT_NK, NONYU_DT, HOSYUSYUSOKU_DT, SYSTEMKADOYOUBI, SYSTEMKADONISSU, SYSTEMKADOJIKAN FROM MCM.MCM_MA_PLANT WHERE PLANT_ID=:id", new MapSqlParameterSource("id",id));
    }

    // ================================================================
    //  納入先検索
    // ================================================================

    /**
     * 【変換元】Mcm2001u_MCM_MA_NONYUSAKITableAdapter.xml
     *   納入先マスタ動的検索（5条件）
     *
     *   Oracle → SQL Server変換:
     *     INSTR(UPPER(A || ' ' || B), UPPER(:param)) > 0
     *     → CHARINDEX(UPPER(@param), UPPER(A + ' ' + B)) > 0
     *     || → +
     *     :param → @param
     */
    public List<Map<String, Object>> searchNonyusaki(Mcm2001uForm form) {
        StringBuilder sql = new StringBuilder();
        MapSqlParameterSource params = new MapSqlParameterSource();

        sql.append("SELECT ");
        sql.append("  MAA.NONYUSAKI_ID, ");
        sql.append("  MAA.NONYUSAKI_CD, ");
        sql.append("  MAA.NONYUSAKI_NK, ");
        sql.append("  MAA.NONYUSAKIKANA_KN, ");
        sql.append("  MAA.NONYUSAKIEIMEI_EN, ");
        sql.append("  MAA.NONYUSAKIKOJO_NK, ");
        sql.append("  MAA.KYUNONYUSAKI_NK ");
        sql.append("FROM MCM.MCM_MA_NONYUSAKI MAA ");
        sql.append("WHERE 1 = 1 ");

        /*
         * 【変換元】CONDITION[0] SUFFIX="%"
         *   元コード: MAA.NONYUSAKI_CD LIKE :{0}
         */
        if (isNotBlank(form.getNonyusakiCd())) {
            sql.append("AND MAA.NONYUSAKI_CD LIKE :nonyusakiCd ");
            params.addValue("nonyusakiCd", form.getNonyusakiCd() + "%");
        }

        /*
         * 【変換元】CONDITION[1]
         *   元コード: INSTR(UPPER(MAA.NONYUSAKI_NK || ' ' || MAA.KYUNONYUSAKI_NK
         *             || ' ' || MAA.NONYUSAKIKOJO_NK || ' ' || MAA.NONYUSAKIKANA_KN
         *             || ' ' || MAA.NONYUSAKIEIMEI_EN), UPPER(:{1})) > 0
         */
        if (isNotBlank(form.getNonyusakiNk())) {
            sql.append("AND CHARINDEX(UPPER(:nonyusakiNk), ");
            sql.append("  UPPER(ISNULL(MAA.NONYUSAKI_NK,'') + ' ' + ISNULL(MAA.KYUNONYUSAKI_NK,'') ");
            sql.append("  + ' ' + ISNULL(MAA.NONYUSAKIKOJO_NK,'') + ' ' + ISNULL(MAA.NONYUSAKIKANA_KN,'') ");
            sql.append("  + ' ' + ISNULL(MAA.NONYUSAKIEIMEI_EN,''))) > 0 ");
            params.addValue("nonyusakiNk", form.getNonyusakiNk());
        }

        /*
         * 【変換元】CONDITION[2]
         *   元コード: INSTR(UPPER(MAA.YUBIN_NO || ' ' || MAA.JUSYO1_NK || ... || MAA.BIKO),
         *             UPPER(:{2})) > 0
         */
        if (isNotBlank(form.getEtc())) {
            sql.append("AND CHARINDEX(UPPER(:etc), ");
            sql.append("  UPPER(ISNULL(MAA.YUBIN_NO,'') + ' ' + ISNULL(MAA.JUSYO1_NK,'') ");
            sql.append("  + ' ' + ISNULL(MAA.JUSYO2_NK,'') + ' ' + ISNULL(MAA.KUNI_NK,'') ");
            sql.append("  + ' ' + ISNULL(MAA.TEL_NO,'') + ' ' + ISNULL(MAA.FAX_NO,'') ");
            sql.append("  + ' ' + ISNULL(MAA.BIKO,''))) > 0 ");
            params.addValue("etc", form.getEtc());
        }

        /*
         * 【変換元】CONDITION[3] SUFFIX="%"
         *   元コード: EXISTS (SELECT * FROM MCM_MA_PLANT MAB
         *             WHERE MAB.NONYUSAKI_ID = MAA.NONYUSAKI_ID AND MAB.SUPPORT_ID LIKE :{3})
         */
        if (isNotBlank(form.getSupportId())) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_PLANT MAB ");
            sql.append("  WHERE MAB.NONYUSAKI_ID = MAA.NONYUSAKI_ID ");
            sql.append("  AND MAB.SUPPORT_ID LIKE :supportId) ");
            params.addValue("supportId", form.getSupportId() + "%");
        }

        /*
         * 【変換元】CONDITION[4] PREFIX="%" SUFFIX="%"
         *   元コード: EXISTS (SELECT * FROM MCM_MA_PLANT MAB
         *             WHERE MAB.NONYUSAKI_ID = MAA.NONYUSAKI_ID
         *             AND UPPER(MAB.PLANT_NK) LIKE UPPER(:{4}))
         */
        if (isNotBlank(form.getPlantNk())) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_PLANT MAB ");
            sql.append("  WHERE MAB.NONYUSAKI_ID = MAA.NONYUSAKI_ID ");
            sql.append("  AND UPPER(MAB.PLANT_NK) LIKE UPPER(:plantNk)) ");
            params.addValue("plantNk", "%" + form.getPlantNk() + "%");
        }

        sql.append("ORDER BY MAA.NONYUSAKI_ID ");

        return jdbc.queryForList(sql.toString(), params);
    }

    // ================================================================
    //  プラント検索
    // ================================================================

    /**
     * 【変換元】Mcm2001u_MCM_MA_PLANTTableAdapter.xml
     *   プラントマスタ動的検索（5条件・EXISTS副問合せ）
     *
     *   Oracle → SQL Server変換:
     *     INSTR → CHARINDEX、|| → +、:param → :param (Named)
     */
    public List<Map<String, Object>> searchPlant(Mcm2001uForm form) {
        StringBuilder sql = new StringBuilder();
        MapSqlParameterSource params = new MapSqlParameterSource();

        sql.append("SELECT ");
        sql.append("  MAB.PLANT_ID, MAB.NONYUSAKI_ID, MAB.SUPPORT_ID, MAB.PLANT_NK, ");
        sql.append("  MAB.NONYUBUSYO_NK, MAB.NONYUTANTOSYA_NK, MAB.NONYUTEL_NO, ");
        sql.append("  MAB.NONYUFAX_NO, MAB.SYSTEMKADOJIKAN, MAB.SYSTEMKADONISSU, ");
        sql.append("  MAB.SYSTEMKADOYOUBI, MAB.NONYU_DT, MAB.TEKKYO_DT, ");
        sql.append("  MAB.HOSYUSYUSOKU_DT, MAB.BIKO, MAB.DTSRENKEI_FLG, ");
        sql.append("  MAB.CREATED_DT, MAB.CREATED_BY, MAB.LASTUPDATE_DT, MAB.LASTUPDATE_BY ");
        sql.append("FROM MCM.MCM_MA_PLANT MAB ");
        sql.append("WHERE 1 = 1 ");

        /*
         * 【変換元】CONDITION[0] SUFFIX="%"
         *   元コード: EXISTS (SELECT MAA.NONYUSAKI_ID FROM MCM_MA_NONYUSAKI MAA
         *             WHERE MAA.NONYUSAKI_ID = MAB.NONYUSAKI_ID AND MAA.NONYUSAKI_CD LIKE :{0})
         */
        if (isNotBlank(form.getNonyusakiCd())) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ");
            sql.append("  WHERE MAA.NONYUSAKI_ID = MAB.NONYUSAKI_ID ");
            sql.append("  AND MAA.NONYUSAKI_CD LIKE :nonyusakiCd) ");
            params.addValue("nonyusakiCd", form.getNonyusakiCd() + "%");
        }

        /*
         * 【変換元】CONDITION[1]
         *   元コード: EXISTS (SELECT * FROM MCM_MA_NONYUSAKI MAA WHERE ...
         *             AND INSTR(UPPER(...), UPPER(:{1})) > 0)
         */
        if (isNotBlank(form.getNonyusakiNk())) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ");
            sql.append("  WHERE MAA.NONYUSAKI_ID = MAB.NONYUSAKI_ID ");
            sql.append("  AND CHARINDEX(UPPER(:nonyusakiNk), ");
            sql.append("    UPPER(ISNULL(MAA.NONYUSAKI_NK,'') + ' ' + ISNULL(MAA.KYUNONYUSAKI_NK,'') ");
            sql.append("    + ' ' + ISNULL(MAA.NONYUSAKIKOJO_NK,'') + ' ' + ISNULL(MAA.NONYUSAKIKANA_KN,'') ");
            sql.append("    + ' ' + ISNULL(MAA.NONYUSAKIEIMEI_EN,''))) > 0) ");
            params.addValue("nonyusakiNk", form.getNonyusakiNk());
        }

        /*
         * 【変換元】CONDITION[2]
         */
        if (isNotBlank(form.getEtc())) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ");
            sql.append("  WHERE MAA.NONYUSAKI_ID = MAB.NONYUSAKI_ID ");
            sql.append("  AND CHARINDEX(UPPER(:etc), ");
            sql.append("    UPPER(ISNULL(MAA.YUBIN_NO,'') + ' ' + ISNULL(MAA.JUSYO1_NK,'') ");
            sql.append("    + ' ' + ISNULL(MAA.JUSYO2_NK,'') + ' ' + ISNULL(MAA.KUNI_NK,'') ");
            sql.append("    + ' ' + ISNULL(MAA.TEL_NO,'') + ' ' + ISNULL(MAA.FAX_NO,'') ");
            sql.append("    + ' ' + ISNULL(MAA.BIKO,''))) > 0) ");
            params.addValue("etc", form.getEtc());
        }

        /*
         * 【変換元】CONDITION[3] SUFFIX="%"
         *   元コード: MAB.SUPPORT_ID LIKE :{3}
         */
        if (isNotBlank(form.getSupportId())) {
            sql.append("AND MAB.SUPPORT_ID LIKE :supportId ");
            params.addValue("supportId", form.getSupportId() + "%");
        }

        /*
         * 【変換元】CONDITION[4] PREFIX="%" SUFFIX="%"
         *   元コード: UPPER(MAB.PLANT_NK) LIKE UPPER(:{4})
         */
        if (isNotBlank(form.getPlantNk())) {
            sql.append("AND UPPER(MAB.PLANT_NK) LIKE UPPER(:plantNk) ");
            params.addValue("plantNk", "%" + form.getPlantNk() + "%");
        }

        sql.append("ORDER BY MAB.SUPPORT_ID ASC ");

        return jdbc.queryForList(sql.toString(), params);
    }

    // ================================================================
    //  取引先見積検索（メイングリッド）
    // ================================================================

    /**
     * 【変換元】Mcm2001u_MCM_TM_MITSUMORITableAdapter.xml
     *   取引先見積 動的検索（6テーブル結合 + DISTINCT + 動的WHERE）
     *
     *   Oracle → SQL Server変換:
     *     NVL(a, b) → ISNULL(a, b)
     *     (+) → LEFT OUTER JOIN（XML元コードで変換済み）
     *     INSTR → CHARINDEX / || → +
     */
    public List<Mcm2001uRowDto> searchMitsumori(Mcm2001uForm form) {
        StringBuilder sql = new StringBuilder();
        MapSqlParameterSource params = new MapSqlParameterSource();

        sql.append("SELECT DISTINCT ");
        sql.append("  0 AS CHECKBOX, ");
        sql.append("  CASE ISNULL(TMV.TKA_JOTAI, '0') WHEN '0' THEN NULL ");
        sql.append("    ELSE ISNULL(TMV.TKA_KEIYAKU_NO, '------') END AS TKA_KEIYAKU_NO, ");
        sql.append("  TMV.TKA_TK_KEIYAKU_ID, ");
        sql.append("  TMB.TM_KEIYAKUJIKAN_ID, TMB.KEIYAKUJIKANTAI, TMB.JOTAI, ");
        sql.append("  TMA.TORIHIKISAKI_NK, TMA.MITSUMORI_DT, TMA.TM_IRAI_ID, ");
        sql.append("  TMA.TM_IRAI_NO, TMA.HOSYUHOHO, TMA.TENKENUMU, ");
        sql.append("  TMA.TENKENKANOYOBI, TMA.YAKANTAIOUMU, MAE.PLANT_ID, ");
        sql.append("  TMA.NONYUSAKI_ID, TMA.NONYUSAKI_CD, TMA.NONYUSAKI_NK, ");
        sql.append("  TMA.SUPPORT_ID, TMA.PLANT_NK ");
        sql.append("FROM MCM.MCM_TM_MITSUMORI TMA ");
        sql.append("  INNER JOIN MCM.MCM_TM_KEIYAKUJIKAN TMB ON TMA.TM_IRAI_ID = TMB.TM_IRAI_ID ");
        sql.append("  INNER JOIN MCM.MCM_TM_KIKIKOSEI TMC ON TMA.TM_IRAI_ID = TMC.TM_IRAI_ID ");
        sql.append("  INNER JOIN MCM.MCM_MA_KIKIKOSEI MAE ON TMC.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID ");
        sql.append("  INNER JOIN MCM.MCM_MA_PLANT MAB ON MAE.PLANT_ID = MAB.PLANT_ID ");
        sql.append("  LEFT OUTER JOIN MCM.MCM_TM_KEIYAKU_NO_V TMV ");
        sql.append("    ON TMV.TMB_TM_KEIYAKUJIKAN_ID = TMB.TM_KEIYAKUJIKAN_ID ");
        sql.append("WHERE TMB.JOTAI IN ('1', '2') ");
        sql.append("  AND (TMV.TKA_JOTAI IS NULL OR TMV.TKA_JOTAI IN ('1','2')) ");

        // ----- 動的条件 -----

        if (isNotBlank(form.getNonyusakiCd())) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ");
            sql.append("  WHERE MAA.NONYUSAKI_ID = TMA.NONYUSAKI_ID ");
            sql.append("  AND MAA.NONYUSAKI_CD LIKE :nonyusakiCd) ");
            params.addValue("nonyusakiCd", form.getNonyusakiCd() + "%");
        }

        if (isNotBlank(form.getNonyusakiNk())) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ");
            sql.append("  WHERE MAA.NONYUSAKI_ID = TMA.NONYUSAKI_ID ");
            sql.append("  AND CHARINDEX(UPPER(:nonyusakiNk), ");
            sql.append("    UPPER(ISNULL(MAA.NONYUSAKI_NK,'') + ' ' + ISNULL(MAA.KYUNONYUSAKI_NK,'') ");
            sql.append("    + ' ' + ISNULL(MAA.NONYUSAKIKOJO_NK,'') + ' ' + ISNULL(MAA.NONYUSAKIKANA_KN,'') ");
            sql.append("    + ' ' + ISNULL(MAA.NONYUSAKIEIMEI_EN,''))) > 0) ");
            params.addValue("nonyusakiNk", form.getNonyusakiNk());
        }

        if (isNotBlank(form.getEtc())) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ");
            sql.append("  WHERE MAA.NONYUSAKI_ID = TMA.NONYUSAKI_ID ");
            sql.append("  AND CHARINDEX(UPPER(:etc), ");
            sql.append("    UPPER(ISNULL(MAA.YUBIN_NO,'') + ' ' + ISNULL(MAA.JUSYO1_NK,'') ");
            sql.append("    + ' ' + ISNULL(MAA.JUSYO2_NK,'') + ' ' + ISNULL(MAA.KUNI_NK,'') ");
            sql.append("    + ' ' + ISNULL(MAA.TEL_NO,'') + ' ' + ISNULL(MAA.FAX_NO,'') ");
            sql.append("    + ' ' + ISNULL(MAA.BIKO,''))) > 0) ");
            params.addValue("etc", form.getEtc());
        }

        if (isNotBlank(form.getSupportId())) {
            sql.append("AND MAB.SUPPORT_ID LIKE :supportId ");
            params.addValue("supportId", form.getSupportId() + "%");
        }

        if (isNotBlank(form.getPlantNk())) {
            sql.append("AND UPPER(MAB.PLANT_NK) LIKE UPPER(:plantNk) ");
            params.addValue("plantNk", "%" + form.getPlantNk() + "%");
        }

        /*
         * 【変換元】CONDITION[5] - プラントID整合性チェック（常時適用）
         *   元コード: (SELECT MIN(CASE SMA.PLANT_ID WHEN NVL(SAE.PLANT_ID, 0) THEN 1 ELSE 0 END)
         *             FROM MCM_TM_MITSUMORI SMA INNER JOIN ... WHERE SMC.TM_IRAI_ID = TMC.TM_IRAI_ID) = 1
         */
        sql.append("AND ( ");
        sql.append("  SELECT MIN(CASE SMA.PLANT_ID WHEN ISNULL(SAE.PLANT_ID, 0) THEN 1 ELSE 0 END) ");
        sql.append("  FROM MCM.MCM_TM_MITSUMORI SMA ");
        sql.append("    INNER JOIN MCM.MCM_TM_KIKIKOSEI SMC ON SMA.TM_IRAI_ID = SMC.TM_IRAI_ID ");
        sql.append("    INNER JOIN MCM.MCM_TM_KIKIMEISAI SMD ON SMD.TM_KIKIKOSEI_ID = SMC.TM_KIKIKOSEI_ID ");
        sql.append("    INNER JOIN MCM.MCM_MA_KIKIMEISAI SAF ON SAF.KIKIMEISAI_ID = SMD.KIKIMEISAI_ID ");
        sql.append("    LEFT OUTER JOIN MCM.MCM_MA_KIKIKOSEI SAE ON SAE.KIKIKOSEI_ID = SAF.KIKIKOSEI_ID ");
        sql.append("  WHERE SMC.TM_IRAI_ID = TMC.TM_IRAI_ID) = 1 ");

        sql.append("ORDER BY TMA.TM_IRAI_NO ASC ");

        return jdbc.query(sql.toString(), params, (rs, rowNum) -> mapMitsumoriRow(rs));
    }

    /**
     * ResultSet → Mcm2001uRowDto マッピング
     */
    private Mcm2001uRowDto mapMitsumoriRow(ResultSet rs) throws SQLException {
        Mcm2001uRowDto row = new Mcm2001uRowDto();
        row.setCheckbox(rs.getInt("CHECKBOX"));
        row.setTorihikisakiNk(rs.getString("TORIHIKISAKI_NK"));
        java.sql.Date dt = rs.getDate("MITSUMORI_DT");
        row.setMitsumoriDt(dt != null ? dt.toLocalDate() : null);
        row.setTmIraiNo(rs.getString("TM_IRAI_NO"));
        row.setTkaKeiyakuNo(rs.getString("TKA_KEIYAKU_NO"));
        row.setKeiyakujikantai(rs.getString("KEIYAKUJIKANTAI"));
        row.setHosyuhoho(rs.getString("HOSYUHOHO"));
        row.setTenkenumu(rs.getString("TENKENUMU"));
        row.setTenkenkanoyobi(rs.getString("TENKENKANOYOBI"));
        row.setYakantaioumu(rs.getString("YAKANTAIOUMU"));
        row.setJotai(rs.getString("JOTAI"));
        row.setPlantId(rs.getBigDecimal("PLANT_ID"));
        row.setNonyusakiId(rs.getBigDecimal("NONYUSAKI_ID"));
        row.setNonyusakiCd(rs.getString("NONYUSAKI_CD"));
        row.setNonyusakiNk(rs.getString("NONYUSAKI_NK"));
        row.setSupportId(rs.getString("SUPPORT_ID"));
        row.setPlantNk(rs.getString("PLANT_NK"));
        row.setTmIraiId(rs.getBigDecimal("TM_IRAI_ID"));
        row.setTmKeiyakujikanId(rs.getBigDecimal("TM_KEIYAKUJIKAN_ID"));
        row.setTkaTkKeiyakuId(rs.getBigDecimal("TKA_TK_KEIYAKU_ID"));
        return row;
    }

    // ================================================================
    //  プラントID指定検索（delivery経由の場合）
    // ================================================================

    /**
     * 【変換元】Mcm2001uScreen.vb - Mcm2001uScreen_Load()内
     *   画面間遷移時にプラントID指定で取引先見積を検索
     *
     *   元コード: Me.DaoContainer.Fill(Me.Mcm2001uDataSet, "MCM_TM_MITSUMORI", ... plantId)
     */
    public List<Mcm2001uRowDto> searchMitsumoriByPlantId(BigDecimal plantId) {
        StringBuilder sql = new StringBuilder();
        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("plantId", plantId);

        sql.append("SELECT DISTINCT ");
        sql.append("  0 AS CHECKBOX, ");
        sql.append("  CASE ISNULL(TMV.TKA_JOTAI, '0') WHEN '0' THEN NULL ");
        sql.append("    ELSE ISNULL(TMV.TKA_KEIYAKU_NO, '------') END AS TKA_KEIYAKU_NO, ");
        sql.append("  TMV.TKA_TK_KEIYAKU_ID, ");
        sql.append("  TMB.TM_KEIYAKUJIKAN_ID, TMB.KEIYAKUJIKANTAI, TMB.JOTAI, ");
        sql.append("  TMA.TORIHIKISAKI_NK, TMA.MITSUMORI_DT, TMA.TM_IRAI_ID, ");
        sql.append("  TMA.TM_IRAI_NO, TMA.HOSYUHOHO, TMA.TENKENUMU, ");
        sql.append("  TMA.TENKENKANOYOBI, TMA.YAKANTAIOUMU, MAE.PLANT_ID, ");
        sql.append("  TMA.NONYUSAKI_ID, TMA.NONYUSAKI_CD, TMA.NONYUSAKI_NK, ");
        sql.append("  TMA.SUPPORT_ID, TMA.PLANT_NK ");
        sql.append("FROM MCM.MCM_TM_MITSUMORI TMA ");
        sql.append("  INNER JOIN MCM.MCM_TM_KEIYAKUJIKAN TMB ON TMA.TM_IRAI_ID = TMB.TM_IRAI_ID ");
        sql.append("  INNER JOIN MCM.MCM_TM_KIKIKOSEI TMC ON TMA.TM_IRAI_ID = TMC.TM_IRAI_ID ");
        sql.append("  INNER JOIN MCM.MCM_MA_KIKIKOSEI MAE ON TMC.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID ");
        sql.append("  INNER JOIN MCM.MCM_MA_PLANT MAB ON MAE.PLANT_ID = MAB.PLANT_ID ");
        sql.append("  LEFT OUTER JOIN MCM.MCM_TM_KEIYAKU_NO_V TMV ");
        sql.append("    ON TMV.TMB_TM_KEIYAKUJIKAN_ID = TMB.TM_KEIYAKUJIKAN_ID ");
        sql.append("WHERE TMB.JOTAI IN ('1', '2', '4') ");
        sql.append("  AND MAE.PLANT_ID = :plantId ");
        sql.append("  AND ( ");
        sql.append("    SELECT MIN(CASE SMA.PLANT_ID WHEN ISNULL(SAE.PLANT_ID, 0) THEN 1 ELSE 0 END) ");
        sql.append("    FROM MCM.MCM_TM_MITSUMORI SMA ");
        sql.append("      INNER JOIN MCM.MCM_TM_KIKIKOSEI SMC ON SMA.TM_IRAI_ID = SMC.TM_IRAI_ID ");
        sql.append("      INNER JOIN MCM.MCM_TM_KIKIMEISAI SMD ON SMD.TM_KIKIKOSEI_ID = SMC.TM_KIKIKOSEI_ID ");
        sql.append("      INNER JOIN MCM.MCM_MA_KIKIMEISAI SAF ON SAF.KIKIMEISAI_ID = SMD.KIKIMEISAI_ID ");
        sql.append("      LEFT OUTER JOIN MCM.MCM_MA_KIKIKOSEI SAE ON SAE.KIKIKOSEI_ID = SAF.KIKIKOSEI_ID ");
        sql.append("    WHERE SMC.TM_IRAI_ID = TMC.TM_IRAI_ID) = 1 ");
        sql.append("ORDER BY TMA.TM_IRAI_NO ASC ");

        return jdbc.query(sql.toString(), params, (rs, rowNum) -> mapMitsumoriRow(rs));
    }

    // ================================================================
    //  店舗基本見積ID → 取引先見積契約時間ID取得
    // ================================================================

    /**
     * 【変換元】Mcm2001uScreen.vb - getUmKeiyakujikanId()
     *   店舗基本見積IDから取引先見積契約時間IDリストを取得
     *
     *   元コード: SELECT DISTINCT TMG.TM_KEIYAKUJIKAN_ID
     *             FROM MCM_UM_MITSUMORI UMG
     *               INNER JOIN MCM_UM_TANKA UMF ON UMG.UM_MITSUMORI_ID = UMF.UM_MITSUMORI_ID
     *               INNER JOIN MCM_TM_TANKA TMF ON UMF.TM_TANKA_ID = TMF.TM_TANKA_ID
     *               INNER JOIN MCM_TM_KIKAN TMG ON TMF.TM_KIKAN_ID = TMG.TM_KIKAN_ID
     *             WHERE UMG.UM_KIHON_MITSUMORI_ID = :umKihonMitsumoriId
     */
    public List<String> findUmKeiyakujikanIds(int umKihonMitsumoriId) {
        String sql =
            "SELECT DISTINCT TMG.TM_KEIYAKUJIKAN_ID " +
            "FROM MCM.MCM_UM_MITSUMORI UMG " +
            "  INNER JOIN MCM.MCM_UM_TANKA UMF ON UMG.UM_MITSUMORI_ID = UMF.UM_MITSUMORI_ID " +
            "  INNER JOIN MCM.MCM_TM_TANKA TMF ON UMF.TM_TANKA_ID = TMF.TM_TANKA_ID " +
            "  INNER JOIN MCM.MCM_TM_KIKAN TMG ON TMF.TM_KIKAN_ID = TMG.TM_KIKAN_ID " +
            "WHERE UMG.UM_KIHON_MITSUMORI_ID = :umKihonMitsumoriId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("umKihonMitsumoriId", umKihonMitsumoriId);

        return jdbc.queryForList(sql, params, String.class);
    }

    // ================================================================
    //  機器明細重複チェック
    // ================================================================

    /**
     * 【変換元】Mcm2001uScreen.vb - SENTEIButton_Click()内SQL
     *   取引先見積内の機器明細重複チェック
     *   選定時に複数のTM_KEIYAKUJIKAN_IDで同一機器明細が重複していないか確認
     *
     *   元コード: SELECT TMD.KIKIMEISAI_ID, TME.KOTAIKANRI_ID, COUNT(*)
     *             FROM MCM_TM_KIKIKOSEI TMC, MCM_TM_KIKIMEISAI TMD,
     *                  MCM_TM_KOTAIMEISAI TME, MCM_TM_MITSUMORI TMA, MCM_TM_KEIYAKUJIKAN TMB
     *             WHERE ... AND TMB.TM_KEIYAKUJIKAN_ID IN (...)
     *             GROUP BY TMD.KIKIMEISAI_ID, TME.KOTAIKANRI_ID
     *             HAVING COUNT(*) > 1
     */
    public boolean hasDuplicateKikimeisai(List<String> tmKeiyakujikanIds) {
        if (tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty()) {
            return false;
        }

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("ids", tmKeiyakujikanIds);

        String sql =
            "SELECT TMD.KIKIMEISAI_ID, TME.KOTAIKANRI_ID, COUNT(*) AS CNT " +
            "FROM MCM.MCM_TM_KIKIKOSEI TMC " +
            "  INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD ON TMC.TM_KIKIKOSEI_ID = TMD.TM_KIKIKOSEI_ID " +
            "  INNER JOIN MCM.MCM_TM_KOTAIMEISAI TME ON TMD.TM_KIKIMEISAI_ID = TME.TM_KIKIMEISAI_ID " +
            "  INNER JOIN MCM.MCM_TM_MITSUMORI TMA ON TMC.TM_IRAI_ID = TMA.TM_IRAI_ID " +
            "  INNER JOIN MCM.MCM_TM_KEIYAKUJIKAN TMB ON TMA.TM_IRAI_ID = TMB.TM_IRAI_ID " +
            "WHERE TMB.TM_KEIYAKUJIKAN_ID IN (:ids) " +
            "GROUP BY TMD.KIKIMEISAI_ID, TME.KOTAIKANRI_ID " +
            "HAVING COUNT(*) > 1";

        List<Map<String, Object>> result = jdbc.queryForList(sql, params);
        return !result.isEmpty();
    }

    // ================================================================
    //  ユーティリティ
    // ================================================================

    private boolean isNotBlank(String s) {
        return s != null && !s.trim().isEmpty();
    }
}
