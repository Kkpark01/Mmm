package com.daifuku.mcm.repository;

import com.daifuku.mcm.dto.Mcm2002uKoseiRowDto;
import com.daifuku.mcm.dto.Mcm2002uMeisaiRowDto;
import com.daifuku.mcm.dto.Mcm2002uKotaiRowDto;
import com.daifuku.mcm.dto.Mcm2002uTankaRowDto;
import org.springframework.jdbc.core.namedparam.MapSqlParameterSource;
import org.springframework.jdbc.core.namedparam.NamedParameterJdbcTemplate;
import org.springframework.stereotype.Repository;

import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * 【変換元】Mcm2002uScreen.vb / Mcm2002uDataSet.Designer.vb / XML SQL定義8ファイル
 * MCM2002U（カスタマー見積機器選定）JdbcTemplate Repository
 *
 * <p>Oracle → SQL Server 変換済み（NVL→ISNULL, DECODE→CASE WHEN, (+)→LEFT JOIN等）</p>
 *
 * @author MCM Migration Tool
 */
@Repository
public class Mcm2002uJdbcRepository {

    private final NamedParameterJdbcTemplate jdbc;

    public Mcm2002uJdbcRepository(NamedParameterJdbcTemplate jdbc) {
        this.jdbc = jdbc;
    }

    // ================================================================
    // RowMapper ヘルパー
    // ================================================================

    private Mcm2002uKoseiRowDto mapKoseiRow(ResultSet rs, int rowNum) throws SQLException {
        Mcm2002uKoseiRowDto dto = new Mcm2002uKoseiRowDto();
        dto.setCheckFlg(getBigDecimal(rs, "CHECK_FLG"));
        dto.setKikikoseiId(getBigDecimal(rs, "KIKIKOSEI_ID"));
        dto.setKikikoseiNk(rs.getString("KIKIKOSEI_NK"));
        dto.setSetNm(getBigDecimal(rs, "SET_NM"));
        dto.setTani(rs.getString("TANI"));
        dto.setTehaiseiban(rs.getString("TEHAISEIBAN"));
        dto.setBiko(rs.getString("BIKO"));
        dto.setControllerFlg(getBigDecimal(rs, "CONTROLLER_FLG"));
        dto.setTmvTmIraiNo(rs.getString("TMV_TM_IRAI_NO"));
        dto.setTmvTmKikikoseiId(getBigDecimal(rs, "TMV_TM_KIKIKOSEI_ID"));
        dto.setUmvUmKikikoseiId(getBigDecimal(rs, "UMV_UM_KIKIKOSEI_ID"));
        dto.setUmvUmKihonBrandId(getBigDecimal(rs, "UMV_UM_KIHON_BRAND_ID"));
        dto.setMitsumoriCheckFlg(getBigDecimal(rs, "MITSUMORICHECK_FLG"));
        dto.setHyojijun(getBigDecimal(rs, "HYOJIJUN"));
        return dto;
    }

    private Mcm2002uMeisaiRowDto mapMeisaiRow(ResultSet rs, int rowNum) throws SQLException {
        Mcm2002uMeisaiRowDto dto = new Mcm2002uMeisaiRowDto();
        dto.setCheckFlg(getBigDecimal(rs, "CHECK_FLG"));
        dto.setKikikoseiId(getBigDecimal(rs, "KIKIKOSEI_ID"));
        dto.setKikimeisaiId(getBigDecimal(rs, "KIKIMEISAI_ID"));
        dto.setKikihinmeiNk(rs.getString("KIKIHINMEI_NK"));
        dto.setKikikatashiki(rs.getString("KIKIKATASHIKI"));
        dto.setSuryoNm(getBigDecimal(rs, "SURYO_NM"));
        dto.setAtsukaikikiId(getBigDecimal(rs, "ATSUKAIKIKI_ID"));
        dto.setHyojijun(getBigDecimal(rs, "HYOJIJUN"));
        dto.setBiko(rs.getString("BIKO"));
        dto.setSeizomakerNk(rs.getString("SEIZOMAKER_NK"));
        dto.setSeizomakerId(getBigDecimal(rs, "SEIZOMAKER_ID"));
        dto.setControllerFlg(getBigDecimal(rs, "CONTROLLER_FLG"));
        dto.setTmvTmIraiNo(rs.getString("TMV_TM_IRAI_NO"));
        dto.setTmvTmKikikoseiId(getBigDecimal(rs, "TMV_TM_KIKIKOSEI_ID"));
        dto.setTmvTmKikimeisaiId(getBigDecimal(rs, "TMV_TM_KIKIMEISAI_ID"));
        dto.setUmvUmKikikoseiId(getBigDecimal(rs, "UMV_UM_KIKIKOSEI_ID"));
        dto.setUmvUmKikimeisaiId(getBigDecimal(rs, "UMV_UM_KIKIMEISAI_ID"));
        dto.setMitsumoriCheckFlg(getBigDecimal(rs, "MITSUMORICHECK_FLG"));
        dto.setSosuNm(getBigDecimal(rs, "SOSU_NM"));
        dto.setMaeHyojijun(getBigDecimal(rs, "MAE_HYOJIJUN"));
        dto.setMafHyojijun(getBigDecimal(rs, "MAF_HYOJIJUN"));
        dto.setKotaikanriFlg(getBigDecimal(rs, "KOTAIKANRI_FLG"));
        return dto;
    }

    private Mcm2002uKotaiRowDto mapKotaiRow(ResultSet rs, int rowNum) throws SQLException {
        Mcm2002uKotaiRowDto dto = new Mcm2002uKotaiRowDto();
        dto.setCheckFlg(getBigDecimal(rs, "CHECK_FLG"));
        dto.setMaeKikikoseiId(getBigDecimal(rs, "MAE_KIKIKOSEI_ID"));
        dto.setMafKikimeisaiId(getBigDecimal(rs, "MAF_KIKIMEISAI_ID"));
        dto.setMagKotaikanriId(getBigDecimal(rs, "MAG_KOTAIKANRI_ID"));
        dto.setMahAtsukaikikiId(getBigDecimal(rs, "MAH_ATSUKAIKIKI_ID"));
        dto.setMagKotaiNk(rs.getString("MAG_KOTAI_NK"));
        dto.setMagSerialNo(rs.getString("MAG_SERIAL_NO"));
        dto.setMagItijinonyuDt(getLocalDate(rs, "MAG_ITIJINONYU_DT"));
        dto.setMagSetchibasyo(rs.getString("MAG_SETCHIBASYO"));
        dto.setMafKikihinmeiNk(rs.getString("MAF_KIKIHINMEI_NK"));
        dto.setMafKikikatashiki(rs.getString("MAF_KIKIKATASHIKI"));
        dto.setMafSuryoNm(getBigDecimal(rs, "MAF_SURYO_NM"));
        dto.setMaeSetNm(getBigDecimal(rs, "MAE_SET_NM"));
        dto.setMakTorihikisakiNk(rs.getString("MAK_TORIHIKISAKI_NK"));
        dto.setMadBrandkoseiId(getBigDecimal(rs, "MAD_BRANDKOSEI_ID"));
        dto.setMadBrandsyosaiNk(rs.getString("MAD_BRANDSYOSAI_NK"));
        dto.setMacBrandId(getBigDecimal(rs, "MAC_BRAND_ID"));
        dto.setMacBrandNk(rs.getString("MAC_BRAND_NK"));
        dto.setMagTekkyoDt(getLocalDate(rs, "MAG_TEKKYO_DT"));
        dto.setMagKeiyakukigenDt(getLocalDate(rs, "MAG_KEIYAKUKIGEN_DT"));
        dto.setMagEnchokeiyakukigenDt(getLocalDate(rs, "MAG_ENCHOKEIYAKUKIGEN_DT"));
        dto.setTmvTmIraiNo(rs.getString("TMV_TM_IRAI_NO"));
        dto.setTmvTmKikikoseiId(getBigDecimal(rs, "TMV_TM_KIKIKOSEI_ID"));
        dto.setTmvTmKikimeisaiId(getBigDecimal(rs, "TMV_TM_KIKIMEISAI_ID"));
        dto.setTmvTmKotaimeisaiId(getBigDecimal(rs, "TMV_TM_KOTAIMEISAI_ID"));
        dto.setUmvUmKikikoseiId(getBigDecimal(rs, "UMV_UM_KIKIKOSEI_ID"));
        dto.setUmvUmKikimeisaiId(getBigDecimal(rs, "UMV_UM_KIKIMEISAI_ID"));
        dto.setUmvUmKotaimeisaiId(getBigDecimal(rs, "UMV_UM_KOTAIMEISAI_ID"));
        dto.setUmbUmKihonBrandId(getBigDecimal(rs, "UMV_UM_KIHON_BRAND_ID"));
        dto.setMitsumoriCheckFlg(getBigDecimal(rs, "MITSUMORICHECK_FLG"));
        dto.setMaeHyojijun(getBigDecimal(rs, "MAE_HYOJIJUN"));
        dto.setMafHyojijun(getBigDecimal(rs, "MAF_HYOJIJUN"));
        dto.setMagHyojijun(getBigDecimal(rs, "MAG_HYOJIJUN"));
        dto.setMahControllerFlg(getBigDecimal(rs, "MAH_CONTROLLER_FLG"));
        dto.setKotaikanriFlg(getBigDecimal(rs, "KOTAIKANRI_FLG"));
        return dto;
    }

    private Mcm2002uTankaRowDto mapTankaRow(ResultSet rs, int rowNum) throws SQLException {
        Mcm2002uTankaRowDto dto = new Mcm2002uTankaRowDto();
        dto.setCheckFlg(getBigDecimal(rs, "CHECK_FLG"));
        dto.setKikikoseiId(getBigDecimal(rs, "KIKIKOSEI_ID"));
        dto.setKikimeisaiId(getBigDecimal(rs, "KIKIMEISAI_ID"));
        dto.setKikihinmeiNk(rs.getString("KIKIHINMEI_NK"));
        dto.setKikikatashiki(rs.getString("KIKIKATASHIKI"));
        dto.setSuryoNm(getBigDecimal(rs, "SURYO_NM"));
        dto.setKeiyakunaiyo(rs.getString("KEIYAKUNAIYO"));
        dto.setKeiyakuNo(rs.getString("KEIYAKU_NO"));
        dto.setTorihosyujikanId(getBigDecimal(rs, "TORIHOSYUJIKAN_ID"));
        dto.setHosyuhoho(rs.getString("HOSYUHOHO"));
        dto.setServicekeitai(rs.getString("SERVICEKEITAI"));
        dto.setKaisiDt(getLocalDate(rs, "KAISI_DT"));
        dto.setSyuryoDt(getLocalDate(rs, "SYURYO_DT"));
        dto.setTenkenkaisu(getBigDecimal(rs, "TENKENKAISU"));
        dto.setTenkenkanoyobi(rs.getString("TENKENKANOYOBI"));
        dto.setPackFlg(getBigDecimal(rs, "PACK_FLG"));
        dto.setHyojunKin(getBigDecimal(rs, "HYOJUN_KIN"));
        dto.setSikiriKin(getBigDecimal(rs, "SIKIRI_KIN"));
        dto.setBrandkoseiId(getBigDecimal(rs, "BRANDKOSEI_ID"));
        dto.setControllerFlg(getBigDecimal(rs, "CONTROLLER_FLG"));
        dto.setTmvTmKikikoseiId(getBigDecimal(rs, "TMV_TM_KIKIKOSEI_ID"));
        dto.setTmvTmKikimeisaiId(getBigDecimal(rs, "TMV_TM_KIKIMEISAI_ID"));
        dto.setTmvTmKikanId(getBigDecimal(rs, "TMV_TM_KIKAN_ID"));
        dto.setTmvTmTankaId(getBigDecimal(rs, "TMV_TM_TANKA_ID"));
        dto.setUmvUmKikikoseiId(getBigDecimal(rs, "UMV_UM_KIKIKOSEI_ID"));
        dto.setUmvUmKikimeisaiId(getBigDecimal(rs, "UMV_UM_KIKIMEISAI_ID"));
        dto.setUmvUmTankaId(getBigDecimal(rs, "UMV_UM_TANKA_ID"));
        dto.setUmvUmMitsumoriId(getBigDecimal(rs, "UMV_UM_MITSUMORI_ID"));
        dto.setMitsumoriCheckFlg(getBigDecimal(rs, "MITSUMORICHECK_FLG"));
        dto.setUmvKaisiDt(getLocalDate(rs, "UMV_KAISI_DT"));
        dto.setUmvSyuryoDt(getLocalDate(rs, "UMV_SYURYO_DT"));
        return dto;
    }

    private BigDecimal getBigDecimal(ResultSet rs, String col) throws SQLException {
        BigDecimal val = rs.getBigDecimal(col);
        return val != null ? val : BigDecimal.ZERO;
    }

    private LocalDate getLocalDate(ResultSet rs, String col) throws SQLException {
        java.sql.Date d = rs.getDate(col);
        return d != null ? d.toLocalDate() : null;
    }

    // ================================================================
    // 1. 機器構成パターン検索（新規：_Ins.xml）
    // 【変換元】Mcm2002u_MCM_MA_KIKIKOSEITableAdapter_Ins.xml
    //   Oracle: DECODE(NVL(LEAST(...)),0,0,1) → SQL Server: CASE WHEN
    //   Oracle: (+) → SQL Server: LEFT JOIN
    // ================================================================
    public List<Mcm2002uKoseiRowDto> searchKikokoseiInsert(
            List<BigDecimal> tmKeiyakujikanIds, BigDecimal plantId) {

        String sql =
            "SELECT DISTINCT " +
            "  CASE WHEN MAE.KIKIKOSEI_ID IS NOT NULL " +
            "        AND (SELECT COUNT(*) FROM MCM.MCM_TM_KIKIKOSEI_V " +
            "             WHERE TMB_TM_KEIYAKUJIKAN_ID IN (:tmKeiyakujikanIds) " +
            "               AND TMC_KIKIKOSEI_ID = MAE.KIKIKOSEI_ID) > 0 " +
            "       THEN 1 ELSE 0 END AS CHECK_FLG " +
            "  ,MAE.KIKIKOSEI_ID " +
            "  ,MAE.KIKIKOSEI_NK " +
            "  ,MAE.SET_NM " +
            "  ,MAE.TANI " +
            "  ,MAE.TEHAISEIBAN " +
            "  ,MAE.BIKO " +
            "  ,MAE.CONTROLLER_FLG " +
            "  ,'' AS TMV_TM_IRAI_NO " +
            "  ,0 AS TMV_TM_KIKIKOSEI_ID " +
            "  ,0 AS UMV_UM_KIKIKOSEI_ID " +
            "  ,0 AS UMV_UM_KIHON_BRAND_ID " +
            "  ,0 AS MITSUMORICHECK_FLG " +
            "  ,MAE.HYOJIJUN " +
            "FROM MCM.MCM_MA_KIKIKOSEI MAE " +
            "  LEFT JOIN ( " +
            "    SELECT TMA.TM_IRAI_NO, TMC.TM_KIKIKOSEI_ID, TMC.KIKIKOSEI_ID " +
            "    FROM MCM.MCM_TM_MITSUMORI TMA " +
            "      INNER JOIN MCM.MCM_TM_KEIYAKUJIKAN TMB ON TMA.TM_IRAI_ID = TMB.TM_IRAI_ID " +
            "      INNER JOIN MCM.MCM_TM_KIKIKOSEI TMC    ON TMA.TM_IRAI_ID = TMC.TM_IRAI_ID " +
            "    WHERE TMB.TM_KEIYAKUJIKAN_ID IN (:tmKeiyakujikanIds) " +
            "  ) TMV ON MAE.KIKIKOSEI_ID = TMV.KIKIKOSEI_ID " +
            "WHERE MAE.PLANT_ID = :plantId " +
            "ORDER BY MAE.HYOJIJUN";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("tmKeiyakujikanIds", tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty() ? List.of(BigDecimal.ZERO) : tmKeiyakujikanIds);
        params.addValue("plantId", plantId);
        return jdbc.query(sql, params, this::mapKoseiRow);
    }

    // ================================================================
    // 2. 機器構成パターン検索（複製：base XML）
    // 【変換元】Mcm2002u_MCM_MA_KIKIKOSEITableAdapter.xml
    //   Oracle: DECODE(NVL(LEAST(MAE.KIKIKOSEI_ID,UMV.UMC_KIKIKOSEI_ID),0),0,0,1)
    //   → SQL Server: CASE WHEN ... IS NOT NULL THEN 1 ELSE 0 END
    // ================================================================
    public List<Mcm2002uKoseiRowDto> searchKikokoseiByUm(
            List<BigDecimal> umKihonMitsumoriIds, BigDecimal plantId) {

        String sql =
            "SELECT DISTINCT " +
            "  CASE WHEN MAE.KIKIKOSEI_ID IS NOT NULL " +
            "        AND UMV.UMC_KIKIKOSEI_ID IS NOT NULL " +
            "       THEN 1 ELSE 0 END AS CHECK_FLG " +
            "  ,MAE.KIKIKOSEI_ID " +
            "  ,MAE.KIKIKOSEI_NK " +
            "  ,MAE.SET_NM " +
            "  ,MAE.TANI " +
            "  ,MAE.TEHAISEIBAN " +
            "  ,MAE.BIKO " +
            "  ,MAE.CONTROLLER_FLG " +
            "  ,'' AS TMV_TM_IRAI_NO " +
            "  ,0 AS TMV_TM_KIKIKOSEI_ID " +
            "  ,ISNULL(UMV.UMC_UM_KIKIKOSEI_ID,0) AS UMV_UM_KIKIKOSEI_ID " +
            "  ,ISNULL(UMV.UMB_UM_KIHON_BRAND_ID,0) AS UMV_UM_KIHON_BRAND_ID " +
            "  ,0 AS MITSUMORICHECK_FLG " +
            "  ,MAE.HYOJIJUN " +
            "FROM MCM.MCM_MA_KIKIKOSEI MAE " +
            "  LEFT JOIN MCM.MCM_UM_KIKIKOSEI_V UMV " +
            "    ON MAE.KIKIKOSEI_ID = UMV.UMC_KIKIKOSEI_ID " +
            "   AND UMV.UMA_UM_KIHON_MITSUMORI_ID IN (:umKihonMitsumoriIds) " +
            "WHERE MAE.PLANT_ID = :plantId " +
            "ORDER BY MAE.HYOJIJUN";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("umKihonMitsumoriIds", umKihonMitsumoriIds);
        params.addValue("plantId", plantId);
        return jdbc.query(sql, params, this::mapKoseiRow);
    }

    // ================================================================
    // 3. 機器構成パターン検索（追加/修正：_Upd.xml）
    // 【変換元】Mcm2002u_MCM_MA_KIKIKOSEITableAdapter_Upd.xml
    //   カスタマー見積修正時に使用。CHECK_FLGはカスタマー側テーブル、
    //   MITSUMORICHECK_FLGは取引先側テーブルで判定。
    // ================================================================
    public List<Mcm2002uKoseiRowDto> searchKikokoseiUpdate(
            List<BigDecimal> umKihonMitsumoriIds,
            List<BigDecimal> tmKeiyakujikanIds,
            BigDecimal plantId) {

        String sql =
            "SELECT DISTINCT " +
            "  CASE WHEN MAE.KIKIKOSEI_ID IS NOT NULL " +
            "        AND (SELECT COUNT(*) FROM MCM.MCM_UM_KIKIMEISAI_V " +
            "             WHERE UMA_UM_KIHON_MITSUMORI_ID IN (:umKihonMitsumoriIds) " +
            "               AND UMC_KIKIKOSEI_ID = MAE.KIKIKOSEI_ID) > 0 " +
            "       THEN 1 ELSE 0 END AS CHECK_FLG " +
            "  ,MAE.KIKIKOSEI_ID " +
            "  ,MAE.KIKIKOSEI_NK " +
            "  ,MAE.SET_NM " +
            "  ,MAE.TANI " +
            "  ,MAE.TEHAISEIBAN " +
            "  ,MAE.BIKO " +
            "  ,MAE.CONTROLLER_FLG " +
            "  ,'' AS TMV_TM_IRAI_NO " +
            "  ,0 AS TMV_TM_KIKIKOSEI_ID " +
            "  ,0 AS UMV_UM_KIKIKOSEI_ID " +
            "  ,0 AS UMV_UM_KIHON_BRAND_ID " +
            "  ,CASE WHEN TMV.KIKIKOSEI_ID IS NOT NULL THEN 1 ELSE 0 END AS MITSUMORICHECK_FLG " +
            "  ,MAE.HYOJIJUN " +
            "FROM MCM.MCM_MA_KIKIKOSEI MAE " +
            "  LEFT JOIN ( " +
            "    SELECT TMA.TM_IRAI_NO, TMC.TM_KIKIKOSEI_ID, TMC.KIKIKOSEI_ID " +
            "    FROM MCM.MCM_TM_MITSUMORI TMA " +
            "      INNER JOIN MCM.MCM_TM_KEIYAKUJIKAN TMB ON TMA.TM_IRAI_ID = TMB.TM_IRAI_ID " +
            "      INNER JOIN MCM.MCM_TM_KIKIKOSEI TMC    ON TMA.TM_IRAI_ID = TMC.TM_IRAI_ID " +
            "    WHERE TMB.TM_KEIYAKUJIKAN_ID IN (:tmKeiyakujikanIds) " +
            "  ) TMV ON MAE.KIKIKOSEI_ID = TMV.KIKIKOSEI_ID " +
            "  LEFT JOIN ( " +
            "    SELECT DISTINCT UMC.KIKIKOSEI_ID, UMB.BRANDKOSEI_ID " +
            "    FROM MCM.MCM_UM_KIHON_MITSUMORI UMA " +
            "      INNER JOIN MCM.MCM_UM_KIHON_BRAND UMB ON UMA.UM_KIHON_MITSUMORI_ID = UMB.UM_KIHON_MITSUMORI_ID " +
            "      INNER JOIN MCM.MCM_UM_KIKIKOSEI UMC   ON UMB.UM_KIHON_BRAND_ID = UMC.UM_KIHON_BRAND_ID " +
            "    WHERE UMA.UM_KIHON_MITSUMORI_ID IN (:umKihonMitsumoriIds) " +
            "  ) UMV ON MAE.KIKIKOSEI_ID = UMV.KIKIKOSEI_ID " +
            "WHERE MAE.PLANT_ID = :plantId " +
            "ORDER BY MAE.HYOJIJUN";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("umKihonMitsumoriIds", umKihonMitsumoriIds);
        params.addValue("tmKeiyakujikanIds", tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty() ? List.of(BigDecimal.ZERO) : tmKeiyakujikanIds);
        params.addValue("plantId", plantId);
        return jdbc.query(sql, params, this::mapKoseiRow);
    }

    // ================================================================
    // 4. 機器明細検索（新規：_Ins.xml）
    // 【変換元】Mcm2002u_MCM_MA_KIKIMEISAITableAdapter_Ins.xml
    //   5テーブル結合 + TMV サブクエリ LEFT JOIN
    // ================================================================
    public List<Mcm2002uMeisaiRowDto> searchKikimeisaiInsert(
            List<BigDecimal> tmKeiyakujikanIds, BigDecimal plantId) {

        String sql =
            "SELECT DISTINCT " +
            "  CASE WHEN MAF.KIKIMEISAI_ID IS NOT NULL " +
            "        AND (SELECT COUNT(*) FROM MCM.MCM_TM_KIKIMEISAI_V " +
            "             WHERE TMB_TM_KEIYAKUJIKAN_ID IN (:tmKeiyakujikanIds) " +
            "               AND TMD_KIKIMEISAI_ID = MAF.KIKIMEISAI_ID) > 0 " +
            "       THEN 1 ELSE 0 END AS CHECK_FLG " +
            "  ,MAF.KIKIKOSEI_ID " +
            "  ,MAF.KIKIMEISAI_ID " +
            "  ,MAF.KIKIHINMEI_NK " +
            "  ,MAF.KIKIKATASHIKI " +
            "  ,MAF.SURYO_NM " +
            "  ,MAF.ATSUKAIKIKI_ID " +
            "  ,MAF.HYOJIJUN " +
            "  ,MAF.BIKO " +
            "  ,MAJ.SEIZOMAKER_NK " +
            "  ,MAJ.SEIZOMAKER_ID " +
            "  ,MAH.CONTROLLER_FLG " +
            "  ,'' AS TMV_TM_IRAI_NO " +
            "  ,0  AS TMV_TM_KIKIKOSEI_ID " +
            "  ,0  AS TMV_TM_KIKIMEISAI_ID " +
            "  ,0 AS UMV_UM_KIKIKOSEI_ID " +
            "  ,0 AS UMV_UM_KIKIMEISAI_ID " +
            "  ,0 AS MITSUMORICHECK_FLG " +
            "  ,ISNULL(MAF.SURYO_NM, 0) * ISNULL(MAE.SET_NM, 0) AS SOSU_NM " +
            "  ,MAE.HYOJIJUN AS MAE_HYOJIJUN " +
            "  ,MAF.HYOJIJUN AS MAF_HYOJIJUN " +
            "  ,MAM.KOTAIKANRI_FLG " +
            "FROM MCM.MCM_MA_KIKIMEISAI MAF " +
            "  INNER JOIN MCM.MCM_MA_ATSUKAIKIKI MAH ON MAF.ATSUKAIKIKI_ID = MAH.ATSUKAIKIKI_ID " +
            "  INNER JOIN MCM.MCM_MA_SEIZOMAKER MAJ  ON MAH.SEIZOMAKER_ID = MAJ.SEIZOMAKER_ID " +
            "  INNER JOIN MCM.MCM_MA_KIKIKOSEI MAE   ON MAF.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID " +
            "  INNER JOIN MCM.MCM_MA_KIKIBUNRUI MAM  ON MAH.KIKIBUNRUI_ID = MAM.KIKIBUNRUI_ID " +
            "  LEFT JOIN ( " +
            "    SELECT TMA.TM_IRAI_NO, TMC.TM_KIKIKOSEI_ID, TMC.KIKIKOSEI_ID, " +
            "           TMD.TM_KIKIMEISAI_ID, TMD.KIKIMEISAI_ID " +
            "    FROM MCM.MCM_TM_MITSUMORI TMA " +
            "      INNER JOIN MCM.MCM_TM_KEIYAKUJIKAN TMB ON TMA.TM_IRAI_ID = TMB.TM_IRAI_ID " +
            "      INNER JOIN MCM.MCM_TM_KIKIKOSEI TMC    ON TMA.TM_IRAI_ID = TMC.TM_IRAI_ID " +
            "      INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD   ON TMC.TM_KIKIKOSEI_ID = TMD.TM_KIKIKOSEI_ID " +
            "    WHERE TMB.TM_KEIYAKUJIKAN_ID IN (:tmKeiyakujikanIds) " +
            "  ) TMV ON MAF.KIKIMEISAI_ID = TMV.KIKIMEISAI_ID " +
            "WHERE MAH.DUMMY_FLG = 0 " +
            "  AND MAE.PLANT_ID = :plantId " +
            "ORDER BY MAE.HYOJIJUN, MAF.HYOJIJUN";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("tmKeiyakujikanIds", tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty() ? List.of(BigDecimal.ZERO) : tmKeiyakujikanIds);
        params.addValue("plantId", plantId);
        return jdbc.query(sql, params, this::mapMeisaiRow);
    }

    // ================================================================
    // 5. 機器明細検索（複製：base XML）
    // 【変換元】Mcm2002u_MCM_MA_KIKIMEISAITableAdapter.xml
    // ================================================================
    public List<Mcm2002uMeisaiRowDto> searchKikimeisaiByUm(
            List<BigDecimal> umKihonMitsumoriIds, BigDecimal plantId) {

        String sql =
            "SELECT DISTINCT " +
            "  CASE WHEN MAF.KIKIMEISAI_ID IS NOT NULL " +
            "        AND UMV.UMD_KIKIMEISAI_ID IS NOT NULL " +
            "       THEN 1 ELSE 0 END AS CHECK_FLG " +
            "  ,MAF.KIKIKOSEI_ID " +
            "  ,MAF.KIKIMEISAI_ID " +
            "  ,MAF.KIKIHINMEI_NK " +
            "  ,MAF.KIKIKATASHIKI " +
            "  ,MAF.SURYO_NM " +
            "  ,MAF.ATSUKAIKIKI_ID " +
            "  ,MAF.HYOJIJUN " +
            "  ,MAF.BIKO " +
            "  ,MAJ.SEIZOMAKER_NK " +
            "  ,MAJ.SEIZOMAKER_ID " +
            "  ,MAH.CONTROLLER_FLG " +
            "  ,'' AS TMV_TM_IRAI_NO " +
            "  ,0  AS TMV_TM_KIKIKOSEI_ID " +
            "  ,0  AS TMV_TM_KIKIMEISAI_ID " +
            "  ,ISNULL(UMV.UMC_UM_KIKIKOSEI_ID,0) AS UMV_UM_KIKIKOSEI_ID " +
            "  ,ISNULL(UMV.UMD_UM_KIKIMEISAI_ID,0) AS UMV_UM_KIKIMEISAI_ID " +
            "  ,0 AS MITSUMORICHECK_FLG " +
            "  ,ISNULL(MAF.SURYO_NM, 0) * ISNULL(MAE.SET_NM, 0) AS SOSU_NM " +
            "  ,MAE.HYOJIJUN AS MAE_HYOJIJUN " +
            "  ,MAF.HYOJIJUN AS MAF_HYOJIJUN " +
            "  ,MAM.KOTAIKANRI_FLG " +
            "FROM MCM.MCM_MA_KIKIKOSEI MAE " +
            "  INNER JOIN MCM.MCM_MA_KIKIMEISAI MAF   ON MAF.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID " +
            "  INNER JOIN MCM.MCM_MA_ATSUKAIKIKI MAH  ON MAF.ATSUKAIKIKI_ID = MAH.ATSUKAIKIKI_ID " +
            "  INNER JOIN MCM.MCM_MA_SEIZOMAKER MAJ   ON MAH.SEIZOMAKER_ID = MAJ.SEIZOMAKER_ID " +
            "  INNER JOIN MCM.MCM_MA_KIKIBUNRUI MAM   ON MAH.KIKIBUNRUI_ID = MAM.KIKIBUNRUI_ID " +
            "  LEFT JOIN MCM.MCM_UM_KIKIMEISAI_V UMV " +
            "    ON MAE.KIKIKOSEI_ID = UMV.UMC_KIKIKOSEI_ID " +
            "   AND MAF.KIKIMEISAI_ID = UMV.UMD_KIKIMEISAI_ID AND UMV.UMA_UM_KIHON_MITSUMORI_ID IN (:umKihonMitsumoriIds) " +
            "WHERE MAH.DUMMY_FLG = 0 " +
            "  AND MAE.PLANT_ID = :plantId " +
            "ORDER BY MAE.HYOJIJUN, MAF.HYOJIJUN";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("umKihonMitsumoriIds", umKihonMitsumoriIds);
        params.addValue("plantId", plantId);
        return jdbc.query(sql, params, this::mapMeisaiRow);
    }

    // ================================================================
    // 6. 機器明細検索（追加/修正：_Upd.xml）
    // 【変換元】Mcm2002u_MCM_MA_KIKIMEISAITableAdapter_Upd.xml
    // TODO: 完全実装 — _Upd.xml の3重サブクエリ対応
    // ================================================================
    public List<Mcm2002uMeisaiRowDto> searchKikimeisaiUpdate(
            List<BigDecimal> umKihonMitsumoriIds,
            List<BigDecimal> tmKeiyakujikanIds,
            BigDecimal plantId) {

        String sql =
            "SELECT DISTINCT " +
            "  CASE WHEN MAF.KIKIMEISAI_ID IS NOT NULL " +
            "        AND (SELECT COUNT(*) FROM MCM.MCM_UM_KIKIMEISAI_V " +
            "             WHERE UMA_UM_KIHON_MITSUMORI_ID IN (:umKihonMitsumoriIds) " +
            "               AND UMD_KIKIMEISAI_ID = MAF.KIKIMEISAI_ID) > 0 " +
            "       THEN 1 ELSE 0 END AS CHECK_FLG " +
            "  ,MAF.KIKIKOSEI_ID " +
            "  ,MAF.KIKIMEISAI_ID " +
            "  ,MAF.KIKIHINMEI_NK " +
            "  ,MAF.KIKIKATASHIKI " +
            "  ,MAF.SURYO_NM " +
            "  ,MAF.ATSUKAIKIKI_ID " +
            "  ,MAF.HYOJIJUN " +
            "  ,MAF.BIKO " +
            "  ,MAJ.SEIZOMAKER_NK " +
            "  ,MAJ.SEIZOMAKER_ID " +
            "  ,MAH.CONTROLLER_FLG " +
            "  ,'' AS TMV_TM_IRAI_NO " +
            "  ,0  AS TMV_TM_KIKIKOSEI_ID " +
            "  ,0  AS TMV_TM_KIKIMEISAI_ID " +
            "  ,0 AS UMV_UM_KIKIKOSEI_ID " +
            "  ,0 AS UMV_UM_KIKIMEISAI_ID " +
            "  ,CASE WHEN TMV.KIKIMEISAI_ID IS NOT NULL THEN 1 ELSE 0 END AS MITSUMORICHECK_FLG " +
            "  ,ISNULL(MAF.SURYO_NM, 0) * ISNULL(MAE.SET_NM, 0) AS SOSU_NM " +
            "  ,MAE.HYOJIJUN AS MAE_HYOJIJUN " +
            "  ,MAF.HYOJIJUN AS MAF_HYOJIJUN " +
            "  ,MAM.KOTAIKANRI_FLG " +
            "FROM MCM.MCM_MA_KIKIMEISAI MAF " +
            "  INNER JOIN MCM.MCM_MA_ATSUKAIKIKI MAH ON MAF.ATSUKAIKIKI_ID = MAH.ATSUKAIKIKI_ID " +
            "  INNER JOIN MCM.MCM_MA_SEIZOMAKER MAJ  ON MAH.SEIZOMAKER_ID = MAJ.SEIZOMAKER_ID " +
            "  INNER JOIN MCM.MCM_MA_KIKIKOSEI MAE   ON MAF.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID " +
            "  INNER JOIN MCM.MCM_MA_KIKIBUNRUI MAM  ON MAH.KIKIBUNRUI_ID = MAM.KIKIBUNRUI_ID " +
            "  LEFT JOIN ( " +
            "    SELECT TMC.KIKIKOSEI_ID, TMD.KIKIMEISAI_ID " +
            "    FROM MCM.MCM_TM_MITSUMORI TMA " +
            "      INNER JOIN MCM.MCM_TM_KEIYAKUJIKAN TMB ON TMA.TM_IRAI_ID = TMB.TM_IRAI_ID " +
            "      INNER JOIN MCM.MCM_TM_KIKIKOSEI TMC    ON TMA.TM_IRAI_ID = TMC.TM_IRAI_ID " +
            "      INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD   ON TMC.TM_KIKIKOSEI_ID = TMD.TM_KIKIKOSEI_ID " +
            "    WHERE TMB.TM_KEIYAKUJIKAN_ID IN (:tmKeiyakujikanIds) " +
            "  ) TMV ON MAF.KIKIMEISAI_ID = TMV.KIKIMEISAI_ID " +
            "WHERE MAH.DUMMY_FLG = 0 " +
            "  AND MAE.PLANT_ID = :plantId " +
            "ORDER BY MAE.HYOJIJUN, MAF.HYOJIJUN";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("umKihonMitsumoriIds", umKihonMitsumoriIds);
        params.addValue("tmKeiyakujikanIds", tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty() ? List.of(BigDecimal.ZERO) : tmKeiyakujikanIds);
        params.addValue("plantId", plantId);
        return jdbc.query(sql, params, this::mapMeisaiRow);
    }

    // ================================================================
    // 7. 機器個体明細検索（新規：_Ins.xml）
    // 【変換元】Mcm2002u_MCM_MA_KOTAIMEISAI_VTableAdapter_Ins.xml
    //   複雑な5テーブル結合TMVサブクエリ + CASE WHEN条件付き表示
    // ================================================================
    public List<Mcm2002uKotaiRowDto> searchKotaimeisaiInsert(
            List<BigDecimal> tmKeiyakujikanIds, BigDecimal plantId) {

        String sql =
            "SELECT " +
            "  CASE WHEN MAV.MAG_KOTAIKANRI_ID IS NOT NULL " +
            "        AND TMV.TME_KOTAIKANRI_ID IS NOT NULL " +
            "       THEN 1 ELSE 0 END AS CHECK_FLG " +
            "  ,MAV.MAE_KIKIKOSEI_ID " +
            "  ,MAV.MAF_KIKIMEISAI_ID " +
            "  ,MAV.MAG_KOTAIKANRI_ID " +
            "  ,MAV.MAH_ATSUKAIKIKI_ID " +
            "  ,MAV.MAG_KOTAI_NK " +
            "  ,MAV.MAG_SERIAL_NO " +
            "  ,MAV.MAG_ITIJINONYU_DT " +
            "  ,MAV.MAG_SETCHIBASYO " +
            "  ,MAV.MAF_KIKIHINMEI_NK " +
            "  ,MAV.MAF_KIKIKATASHIKI " +
            "  ,MAV.MAF_SURYO_NM " +
            "  ,MAV.MAE_SET_NM " +
            "  ,CASE WHEN TMV.TMA_TM_IRAI_NO IS NOT NULL OR MAV.MAH_CONTROLLER_FLG = 1 " +
            "        THEN MAV.MAK_TORIHIKISAKI_NK ELSE '' END AS MAK_TORIHIKISAKI_NK " +
            "  ,MAV.MAD_BRANDKOSEI_ID " +
            "  ,MAV.MAD_BRANDSYOSAI_NK " +
            "  ,MAV.MAC_BRAND_ID " +
            "  ,MAV.MAC_BRAND_NK " +
            "  ,MAV.MAG_TEKKYO_DT " +
            "  ,MAV.MAG_KEIYAKUKIGEN_DT " +
            "  ,MAV.MAG_ENCHOKEIYAKUKIGEN_DT " +
            "  ,ISNULL(TMV.TMA_TM_IRAI_NO, '')          AS TMV_TM_IRAI_NO " +
            "  ,ISNULL(TMV.TMC_TM_KIKIKOSEI_ID, 0)      AS TMV_TM_KIKIKOSEI_ID " +
            "  ,ISNULL(TMV.TMD_TM_KIKIMEISAI_ID, 0)     AS TMV_TM_KIKIMEISAI_ID " +
            "  ,ISNULL(TMV.TME_TM_KOTAIMEISAI_ID, 0)    AS TMV_TM_KOTAIMEISAI_ID " +
            "  ,0 AS UMV_UM_KIKIKOSEI_ID " +
            "  ,0 AS UMV_UM_KIKIMEISAI_ID " +
            "  ,0 AS UMV_UM_KOTAIMEISAI_ID " +
            "  ,0 AS UMV_UM_KIHON_BRAND_ID " +
            "  ,0 AS MITSUMORICHECK_FLG " +
            "  ,MAV.MAE_HYOJIJUN " +
            "  ,MAV.MAF_HYOJIJUN " +
            "  ,MAV.MAG_HYOJIJUN " +
            "  ,MAV.MAH_CONTROLLER_FLG " +
            "  ,MAM.KOTAIKANRI_FLG " +
            "FROM MCM.MCM_MA_KOTAIMEISAI_V MAV " +
            "  INNER JOIN MCM.MCM_MA_KIKIBUNRUI MAM ON MAV.MAH_KIKIBUNRUI_ID = MAM.KIKIBUNRUI_ID " +
            "  LEFT JOIN ( " +
            "    SELECT TMA.TM_IRAI_NO         AS TMA_TM_IRAI_NO, " +
            "           TMC.TM_KIKIKOSEI_ID    AS TMC_TM_KIKIKOSEI_ID, " +
            "           TMD.TM_KIKIMEISAI_ID   AS TMD_TM_KIKIMEISAI_ID, " +
            "           TME.TM_KOTAIMEISAI_ID  AS TME_TM_KOTAIMEISAI_ID, " +
            "           TME.KOTAIKANRI_ID      AS TME_KOTAIKANRI_ID, " +
            "           TMC.KIKIKOSEI_ID, TMD.KIKIMEISAI_ID " +
            "    FROM MCM.MCM_TM_MITSUMORI TMA " +
            "      INNER JOIN MCM.MCM_TM_KEIYAKUJIKAN TMB  ON TMA.TM_IRAI_ID = TMB.TM_IRAI_ID " +
            "      INNER JOIN MCM.MCM_TM_KIKIKOSEI TMC     ON TMA.TM_IRAI_ID = TMC.TM_IRAI_ID " +
            "      INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD    ON TMC.TM_KIKIKOSEI_ID = TMD.TM_KIKIKOSEI_ID " +
            "      INNER JOIN MCM.MCM_TM_KOTAIMEISAI TME   ON TMD.TM_KIKIMEISAI_ID = TME.TM_KIKIMEISAI_ID " +
            "    WHERE TMB.TM_KEIYAKUJIKAN_ID IN (:tmKeiyakujikanIds) " +
            "  ) TMV ON MAV.MAG_KOTAIKANRI_ID = TMV.TME_KOTAIKANRI_ID " +
            "       AND MAV.MAF_KIKIMEISAI_ID = TMV.KIKIMEISAI_ID " +
            "       AND MAV.MAE_KIKIKOSEI_ID  = TMV.KIKIKOSEI_ID " +
            "WHERE MAV.MAI_HOSYUKEIYAKUTAISYO_FLG = 0 " +
            "  AND MAV.MAH_DUMMY_FLG = 0 " +
            "  AND MAV.MAE_PLANT_ID = :plantId " +
            "ORDER BY MAV.MAE_HYOJIJUN, MAV.MAF_HYOJIJUN, MAV.MAG_HYOJIJUN";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("tmKeiyakujikanIds", tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty() ? List.of(BigDecimal.ZERO) : tmKeiyakujikanIds);
        params.addValue("plantId", plantId);
        return jdbc.query(sql, params, this::mapKotaiRow);
    }

    // ================================================================
    // 8. 機器個体明細検索（複製：base XML）
    // 【変換元】Mcm2002u_MCM_MA_KOTAIMEISAI_VTableAdapter.xml
    // ================================================================
    public List<Mcm2002uKotaiRowDto> searchKotaimeisaiByUm(
            List<BigDecimal> umKihonMitsumoriIds, BigDecimal plantId) {

        String sql =
            "SELECT DISTINCT " +
            "  CASE WHEN MAV.MAG_KOTAIKANRI_ID IS NOT NULL " +
            "        AND UMV.UME_UM_KOTAIMEISAI_ID IS NOT NULL " +
            "       THEN 1 ELSE 0 END AS CHECK_FLG " +
            "  ,MAV.MAE_KIKIKOSEI_ID " +
            "  ,MAV.MAF_KIKIMEISAI_ID " +
            "  ,MAV.MAG_KOTAIKANRI_ID " +
            "  ,MAV.MAH_ATSUKAIKIKI_ID " +
            "  ,MAV.MAG_KOTAI_NK " +
            "  ,MAV.MAG_SERIAL_NO " +
            "  ,MAV.MAG_ITIJINONYU_DT " +
            "  ,MAV.MAG_SETCHIBASYO " +
            "  ,MAV.MAF_KIKIHINMEI_NK " +
            "  ,MAV.MAF_KIKIKATASHIKI " +
            "  ,MAV.MAF_SURYO_NM " +
            "  ,MAV.MAE_SET_NM " +
            "  ,MAV.MAK_TORIHIKISAKI_NK " +
            "  ,MAV.MAD_BRANDKOSEI_ID " +
            "  ,MAV.MAD_BRANDSYOSAI_NK " +
            "  ,MAV.MAC_BRAND_ID " +
            "  ,MAV.MAC_BRAND_NK " +
            "  ,MAV.MAG_TEKKYO_DT " +
            "  ,MAV.MAG_KEIYAKUKIGEN_DT " +
            "  ,MAV.MAG_ENCHOKEIYAKUKIGEN_DT " +
            "  ,'' AS TMV_TM_IRAI_NO " +
            "  ,0 AS TMV_TM_KIKIKOSEI_ID " +
            "  ,0 AS TMV_TM_KIKIMEISAI_ID " +
            "  ,0 AS TMV_TM_KOTAIMEISAI_ID " +
            "  ,ISNULL(UMV.UMC_UM_KIKIKOSEI_ID, 0)   AS UMV_UM_KIKIKOSEI_ID " +
            "  ,ISNULL(UMV.UMD_UM_KIKIMEISAI_ID, 0)  AS UMV_UM_KIKIMEISAI_ID " +
            "  ,ISNULL(UMV.UME_UM_KOTAIMEISAI_ID, 0) AS UMV_UM_KOTAIMEISAI_ID " +
            "  ,ISNULL(UMV.UMB_UM_KIHON_BRAND_ID, 0) AS UMV_UM_KIHON_BRAND_ID " +
            "  ,0 AS MITSUMORICHECK_FLG " +
            "  ,MAV.MAE_HYOJIJUN " +
            "  ,MAV.MAF_HYOJIJUN " +
            "  ,MAV.MAG_HYOJIJUN " +
            "  ,MAV.MAH_CONTROLLER_FLG " +
            "  ,MAM.KOTAIKANRI_FLG " +
            "FROM MCM.MCM_MA_KOTAIMEISAI_V MAV " +
            "  INNER JOIN MCM.MCM_MA_KIKIBUNRUI MAM ON MAV.MAH_KIKIBUNRUI_ID = MAM.KIKIBUNRUI_ID " +
            "  LEFT JOIN ( " +
            "    SELECT UME.UM_KOTAIMEISAI_ID AS UME_UM_KOTAIMEISAI_ID, UME.KOTAIKANRI_ID, " +
            "           UMC.KIKIKOSEI_ID      AS UMC_KIKIKOSEI_ID, " +
            "           UMD.KIKIMEISAI_ID     AS UMD_KIKIMEISAI_ID, " +
            "           UMB.BRANDKOSEI_ID     AS UMB_BRANDKOSEI_ID, " +
            "           UMB.UM_KIHON_BRAND_ID AS UMB_UM_KIHON_BRAND_ID, " +
            "           UMC.UM_KIKIKOSEI_ID   AS UMC_UM_KIKIKOSEI_ID, " +
            "           UMD.UM_KIKIMEISAI_ID  AS UMD_UM_KIKIMEISAI_ID " +
            "    FROM MCM.MCM_UM_KIHON_BRAND UMB " +
            "      INNER JOIN MCM.MCM_UM_KIKIKOSEI UMC   ON UMC.UM_KIHON_BRAND_ID = UMB.UM_KIHON_BRAND_ID " +
            "      INNER JOIN MCM.MCM_UM_KIKIMEISAI UMD  ON UMD.UM_KIKIKOSEI_ID = UMC.UM_KIKIKOSEI_ID " +
            "      INNER JOIN MCM.MCM_UM_KOTAIMEISAI UME ON UME.UM_KIKIMEISAI_ID = UMD.UM_KIKIMEISAI_ID " +
            "    WHERE UMB.UM_KIHON_MITSUMORI_ID IN (:umKihonMitsumoriIds) " +
            "  ) UMV ON MAV.MAG_KOTAIKANRI_ID = UMV.KOTAIKANRI_ID AND MAV.MAF_KIKIMEISAI_ID = UMV.UMD_KIKIMEISAI_ID " +
            "       AND MAV.MAE_KIKIKOSEI_ID  = UMV.UMC_KIKIKOSEI_ID " +
            "       AND MAV.MAD_BRANDKOSEI_ID = UMV.UMB_BRANDKOSEI_ID " +
            "WHERE MAV.MAI_HOSYUKEIYAKUTAISYO_FLG = 0 " +
            "  AND MAV.MAH_DUMMY_FLG = 0 " +
            "  AND MAV.MAE_PLANT_ID = :plantId " +
            "ORDER BY MAV.MAE_HYOJIJUN, MAV.MAF_HYOJIJUN, MAV.MAG_HYOJIJUN";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("umKihonMitsumoriIds", umKihonMitsumoriIds);
        params.addValue("plantId", plantId);
        return jdbc.query(sql, params, this::mapKotaiRow);
    }

    // ================================================================
    // 9. 機器個体明細検索（追加/修正：_Upd.xml）
    // 【変換元】Mcm2002u_MCM_MA_KOTAIMEISAI_VTableAdapter_Upd.xml
    // TODO: _Upd.xml の UMV + TMV 二重サブクエリ対応
    //   UMV側でカスタマー見積チェック、TMV側で見積チェック（MITSUMORICHECK_FLG）を判定
    // ================================================================
    public List<Mcm2002uKotaiRowDto> searchKotaimeisaiUpdate(
            List<BigDecimal> umKihonMitsumoriIds,
            List<BigDecimal> tmKeiyakujikanIds,
            BigDecimal plantId) {

        String sql =
            "SELECT DISTINCT " +
            "  CASE WHEN UMV.UME_UM_KOTAIMEISAI_ID IS NOT NULL THEN 1 ELSE 0 END AS CHECK_FLG " +
            "  ,MAV.MAE_KIKIKOSEI_ID " +
            "  ,MAV.MAF_KIKIMEISAI_ID " +
            "  ,MAV.MAG_KOTAIKANRI_ID " +
            "  ,MAV.MAH_ATSUKAIKIKI_ID " +
            "  ,MAV.MAG_KOTAI_NK " +
            "  ,MAV.MAG_SERIAL_NO " +
            "  ,MAV.MAG_ITIJINONYU_DT " +
            "  ,MAV.MAG_SETCHIBASYO " +
            "  ,MAV.MAF_KIKIHINMEI_NK " +
            "  ,MAV.MAF_KIKIKATASHIKI " +
            "  ,MAV.MAF_SURYO_NM " +
            "  ,MAV.MAE_SET_NM " +
            "  ,MAV.MAK_TORIHIKISAKI_NK " +
            "  ,MAV.MAD_BRANDKOSEI_ID " +
            "  ,MAV.MAD_BRANDSYOSAI_NK " +
            "  ,MAV.MAC_BRAND_ID " +
            "  ,MAV.MAC_BRAND_NK " +
            "  ,MAV.MAG_TEKKYO_DT " +
            "  ,MAV.MAG_KEIYAKUKIGEN_DT " +
            "  ,MAV.MAG_ENCHOKEIYAKUKIGEN_DT " +
            "  ,ISNULL(TMV.TMA_TM_IRAI_NO, '')          AS TMV_TM_IRAI_NO " +
            "  ,ISNULL(TMV.TMC_TM_KIKIKOSEI_ID, 0)      AS TMV_TM_KIKIKOSEI_ID " +
            "  ,ISNULL(TMV.TMD_TM_KIKIMEISAI_ID, 0)     AS TMV_TM_KIKIMEISAI_ID " +
            "  ,ISNULL(TMV.TME_TM_KOTAIMEISAI_ID, 0)    AS TMV_TM_KOTAIMEISAI_ID " +
            "  ,ISNULL(UMV.UMC_UM_KIKIKOSEI_ID, 0)      AS UMV_UM_KIKIKOSEI_ID " +
            "  ,ISNULL(UMV.UMD_UM_KIKIMEISAI_ID, 0)     AS UMV_UM_KIKIMEISAI_ID " +
            "  ,ISNULL(UMV.UME_UM_KOTAIMEISAI_ID, 0)    AS UMV_UM_KOTAIMEISAI_ID " +
            "  ,ISNULL(UMV.UMB_UM_KIHON_BRAND_ID, 0)    AS UMV_UM_KIHON_BRAND_ID " +
            "  ,CASE WHEN TMV.TME_KOTAIKANRI_ID IS NOT NULL THEN 1 ELSE 0 END AS MITSUMORICHECK_FLG " +
            "  ,MAV.MAE_HYOJIJUN " +
            "  ,MAV.MAF_HYOJIJUN " +
            "  ,MAV.MAG_HYOJIJUN " +
            "  ,MAV.MAH_CONTROLLER_FLG " +
            "  ,MAM.KOTAIKANRI_FLG " +
            "FROM MCM.MCM_MA_KOTAIMEISAI_V MAV " +
            "  INNER JOIN MCM.MCM_MA_KIKIBUNRUI MAM ON MAV.MAH_KIKIBUNRUI_ID = MAM.KIKIBUNRUI_ID " +
            "  LEFT JOIN ( " +
            "    SELECT UME.UM_KOTAIMEISAI_ID AS UME_UM_KOTAIMEISAI_ID, UME.KOTAIKANRI_ID, " +
            "           UMC.KIKIKOSEI_ID, UMD.KIKIMEISAI_ID, " +
            "           UMB.BRANDKOSEI_ID     AS UMB_BRANDKOSEI_ID, " +
            "           UMB.UM_KIHON_BRAND_ID AS UMB_UM_KIHON_BRAND_ID, " +
            "           UMC.UM_KIKIKOSEI_ID   AS UMC_UM_KIKIKOSEI_ID, " +
            "           UMD.UM_KIKIMEISAI_ID  AS UMD_UM_KIKIMEISAI_ID " +
            "    FROM MCM.MCM_UM_KIHON_BRAND UMB " +
            "      INNER JOIN MCM.MCM_UM_KIKIKOSEI UMC   ON UMC.UM_KIHON_BRAND_ID = UMB.UM_KIHON_BRAND_ID " +
            "      INNER JOIN MCM.MCM_UM_KIKIMEISAI UMD  ON UMD.UM_KIKIKOSEI_ID = UMC.UM_KIKIKOSEI_ID " +
            "      INNER JOIN MCM.MCM_UM_KOTAIMEISAI UME ON UME.UM_KIKIMEISAI_ID = UMD.UM_KIKIMEISAI_ID " +
            "    WHERE UMB.UM_KIHON_MITSUMORI_ID IN (:umKihonMitsumoriIds) " +
            "  ) UMV ON MAV.MAG_KOTAIKANRI_ID = UMV.KOTAIKANRI_ID " +
            "       AND MAV.MAF_KIKIMEISAI_ID = UMV.KIKIMEISAI_ID " +
            "       AND MAV.MAE_KIKIKOSEI_ID  = UMV.KIKIKOSEI_ID " +
            "       AND MAV.MAD_BRANDKOSEI_ID = UMV.UMB_BRANDKOSEI_ID " +
            "  LEFT JOIN ( " +
            "    SELECT TMA.TM_IRAI_NO         AS TMA_TM_IRAI_NO, " +
            "           TMC.TM_KIKIKOSEI_ID    AS TMC_TM_KIKIKOSEI_ID, " +
            "           TMD.TM_KIKIMEISAI_ID   AS TMD_TM_KIKIMEISAI_ID, " +
            "           TME.TM_KOTAIMEISAI_ID  AS TME_TM_KOTAIMEISAI_ID, " +
            "           TME.KOTAIKANRI_ID      AS TME_KOTAIKANRI_ID, " +
            "           TMC.KIKIKOSEI_ID, TMD.KIKIMEISAI_ID " +
            "    FROM MCM.MCM_TM_MITSUMORI TMA " +
            "      INNER JOIN MCM.MCM_TM_KEIYAKUJIKAN TMB  ON TMA.TM_IRAI_ID = TMB.TM_IRAI_ID " +
            "      INNER JOIN MCM.MCM_TM_KIKIKOSEI TMC     ON TMA.TM_IRAI_ID = TMC.TM_IRAI_ID " +
            "      INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD    ON TMC.TM_KIKIKOSEI_ID = TMD.TM_KIKIKOSEI_ID " +
            "      INNER JOIN MCM.MCM_TM_KOTAIMEISAI TME   ON TMD.TM_KIKIMEISAI_ID = TME.TM_KIKIMEISAI_ID " +
            "    WHERE TMB.TM_KEIYAKUJIKAN_ID IN (:tmKeiyakujikanIds) " +
            "  ) TMV ON MAV.MAG_KOTAIKANRI_ID = TMV.TME_KOTAIKANRI_ID " +
            "       AND MAV.MAF_KIKIMEISAI_ID = TMV.KIKIMEISAI_ID " +
            "       AND MAV.MAE_KIKIKOSEI_ID  = TMV.KIKIKOSEI_ID " +
            "WHERE MAV.MAI_HOSYUKEIYAKUTAISYO_FLG = 0 " +
            "  AND MAV.MAH_DUMMY_FLG = 0 " +
            "  AND MAV.MAE_PLANT_ID = :plantId " +
            "ORDER BY MAV.MAE_HYOJIJUN, MAV.MAF_HYOJIJUN, MAV.MAG_HYOJIJUN";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("umKihonMitsumoriIds", umKihonMitsumoriIds);
        params.addValue("tmKeiyakujikanIds", tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty() ? List.of(BigDecimal.ZERO) : tmKeiyakujikanIds);
        params.addValue("plantId", plantId);
        return jdbc.query(sql, params, this::mapKotaiRow);
    }

    // ================================================================
    // 10. 見積単価検索（新規：_Ins.xml）
    // 【変換元】Mcm2002u_MCM_TM_TANKATableAdapter_Ins.xml
    // TODO: 完全な _Ins.xml 変換 — 現在は簡易版
    // ================================================================
    public List<Mcm2002uTankaRowDto> searchTankaInsert(
            List<BigDecimal> tmKeiyakujikanIds, BigDecimal plantId) {

        String sql =
            "SELECT " +
            "  1 AS CHECK_FLG " +
            "  ,TMC.KIKIKOSEI_ID " +
            "  ,TMD.KIKIMEISAI_ID " +
            "  ,MAF.KIKIHINMEI_NK " +
            "  ,MAF.KIKIKATASHIKI " +
            "  ,0 AS SURYO_NM " +
            "  ,TMF.KEIYAKUNAIYO " +
            "  ,TMF.KEIYAKU_NO " +
            "  ,TMF.TORIHOSYUJIKAN_ID " +
            "  ,TMF.HOSYUHOHO " +
            "  ,TMF.SERVICEKEITAI " +
            "  ,TMG.KAISI_DT " +
            "  ,TMG.SYURYO_DT " +
            "  ,TMH.TENKENKAISU " +
            "  ,TMH.TENKENKANOYOBI " +
            "  ,TMF.PACK_FLG " +
            "  ,TMF.HYOJUN_KIN " +
            "  ,TMF.SIKIRI_KIN " +
            "  ,BID.BRANDKOSEI_ID " +
            "  ,MAH.CONTROLLER_FLG " +
            "  ,TMC.TM_KIKIKOSEI_ID  AS TMV_TM_KIKIKOSEI_ID " +
            "  ,TMD.TM_KIKIMEISAI_ID AS TMV_TM_KIKIMEISAI_ID " +
            "  ,TMG.TM_KIKAN_ID      AS TMV_TM_KIKAN_ID " +
            "  ,TMF.TM_TANKA_ID      AS TMV_TM_TANKA_ID " +
            "  ,0 AS UMV_UM_KIKIKOSEI_ID " +
            "  ,0 AS UMV_UM_KIKIMEISAI_ID " +
            "  ,0 AS UMV_UM_TANKA_ID " +
            "  ,0 AS UMV_UM_MITSUMORI_ID " +
            "  ,0 AS MITSUMORICHECK_FLG " +
            "  ,TMG.KAISI_DT AS UMV_KAISI_DT " +
            "  ,TMG.SYURYO_DT AS UMV_SYURYO_DT " +
            "FROM MCM.MCM_TM_MITSUMORI TMA " +
            "  INNER JOIN MCM.MCM_TM_KEIYAKUJIKAN TMB ON TMA.TM_IRAI_ID = TMB.TM_IRAI_ID " +
            "  INNER JOIN MCM.MCM_TM_KIKIKOSEI TMC    ON TMA.TM_IRAI_ID = TMC.TM_IRAI_ID " +
            "  INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD   ON TMC.TM_KIKIKOSEI_ID = TMD.TM_KIKIKOSEI_ID " +
            "  INNER JOIN MCM.MCM_TM_KIKAN TMG        ON TMB.TM_KEIYAKUJIKAN_ID = TMG.TM_KEIYAKUJIKAN_ID " +
            "  INNER JOIN MCM.MCM_TM_TANKA TMF        ON TMD.TM_KIKIMEISAI_ID = TMF.TM_KIKIMEISAI_ID " +
            "                                      AND TMG.TM_KIKAN_ID = TMF.TM_KIKAN_ID " +
            "  INNER JOIN (SELECT DISTINCT TM_KIKIMEISAI_ID,BRANDKOSEI_ID FROM MCM.MCM_TM_KOTAIMEISAI) BID ON BID.TM_KIKIMEISAI_ID=TMD.TM_KIKIMEISAI_ID " +
            "  INNER JOIN MCM.MCM_MA_KIKIMEISAI MAF   ON TMD.KIKIMEISAI_ID = MAF.KIKIMEISAI_ID " +
            "  INNER JOIN MCM.MCM_MA_ATSUKAIKIKI MAH  ON MAF.ATSUKAIKIKI_ID = MAH.ATSUKAIKIKI_ID " +
            "  LEFT JOIN MCM.MCM_MA_ATSUKAIKIKI TAH ON TAH.ATSUKAIKIKI_ID=TMD.ATSUKAIKIKI_ID " +
            "  LEFT JOIN MCM.MCM_MA_KIKIBUNRUI TB ON TB.KIKIBUNRUI_ID=TAH.KIKIBUNRUI_ID " +
            "  LEFT JOIN MCM.MCM_TM_TENKEN TMH ON TMH.TM_KIKIKOSEI_ID=TMC.TM_KIKIKOSEI_ID AND TMH.TM_KIKAN_ID=TMG.TM_KIKAN_ID AND TMH.OYAKIKIBUNRUI_CD=TB.OYAKIKIBUNRUI_CD " +
            "WHERE TMB.TM_KEIYAKUJIKAN_ID IN (:tmKeiyakujikanIds) " +
            "  AND TMC.KIKIKOSEI_ID IN ( " +
            "    SELECT MAE2.KIKIKOSEI_ID FROM MCM.MCM_MA_KIKIKOSEI MAE2 " +
            "    WHERE MAE2.PLANT_ID = :plantId ) " +
            "ORDER BY TMC.KIKIKOSEI_ID, TMD.KIKIMEISAI_ID, TMG.KAISI_DT";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("tmKeiyakujikanIds", tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty() ? List.of(BigDecimal.ZERO) : tmKeiyakujikanIds);
        params.addValue("plantId", plantId);
        return jdbc.query(sql, params, this::mapTankaRow);
    }

    // ================================================================
    // 11. 見積単価検索（修正：_Upd.xml）
    // 【変換元】Mcm2002u_MCM_TM_TANKATableAdapter_Upd.xml
    // TODO: 完全な _Upd.xml 変換 — カスタマー見積ID付加ロジック
    // ================================================================
    public List<Mcm2002uTankaRowDto> searchTankaUpdate(List<BigDecimal> umIds,List<BigDecimal> tmIds,BigDecimal plantId) {
        String sql="""
            SELECT 1 CHECK_FLG,C.KIKIKOSEI_ID,D.KIKIMEISAI_ID,D.KIKIHINMEI_NK,D.KIKIKATASHIKI,D.SURYO_NM,
              F.KEIYAKUNAIYO,F.KEIYAKU_NO,F.TORIHOSYUJIKAN_ID,F.HOSYUHOHO,F.SERVICEKEITAI,
              P.KAISI_DT,P.SYURYO_DT,F.TENKENKAISU,F.TENKENYOBI TENKENKANOYOBI,F.PACK_FLG,F.HYOJUN_KIN,F.SIKIRI_KIN,
              B.BRANDKOSEI_ID,A.CONTROLLER_FLG,0 TMV_TM_KIKIKOSEI_ID,COALESCE(T.TM_KIKIMEISAI_ID,0) TMV_TM_KIKIMEISAI_ID,
              COALESCE(T.TM_KIKAN_ID,0) TMV_TM_KIKAN_ID,COALESCE(F.TM_TANKA_ID,0) TMV_TM_TANKA_ID,
              C.UM_KIKIKOSEI_ID UMV_UM_KIKIKOSEI_ID,D.UM_KIKIMEISAI_ID UMV_UM_KIKIMEISAI_ID,F.UM_TANKA_ID UMV_UM_TANKA_ID,
              P.UM_MITSUMORI_ID UMV_UM_MITSUMORI_ID,1 MITSUMORICHECK_FLG,P.KAISI_DT UMV_KAISI_DT,P.SYURYO_DT UMV_SYURYO_DT
            FROM MCM.MCM_UM_KIHON_MITSUMORI H JOIN MCM.MCM_UM_KIHON_BRAND B ON B.UM_KIHON_MITSUMORI_ID=H.UM_KIHON_MITSUMORI_ID
              JOIN MCM.MCM_UM_KIKIKOSEI C ON C.UM_KIHON_BRAND_ID=B.UM_KIHON_BRAND_ID
              JOIN MCM.MCM_UM_KIKIMEISAI D ON D.UM_KIKIKOSEI_ID=C.UM_KIKIKOSEI_ID
              JOIN MCM.MCM_UM_TANKA F ON F.UM_KIKIMEISAI_ID=D.UM_KIKIMEISAI_ID
              JOIN MCM.MCM_UM_MITSUMORI P ON P.UM_MITSUMORI_ID=F.UM_MITSUMORI_ID
              LEFT JOIN MCM.MCM_MA_ATSUKAIKIKI A ON A.ATSUKAIKIKI_ID=D.ATSUKAIKIKI_ID
              LEFT JOIN MCM.MCM_TM_TANKA T ON T.TM_TANKA_ID=F.TM_TANKA_ID
            WHERE H.UM_KIHON_MITSUMORI_ID IN (:umIds) AND H.PLANT_ID=:plantId ORDER BY P.KAISI_DT,D.KIKIMEISAI_ID
            """;
        var result=new ArrayList<>(jdbc.query(sql,new MapSqlParameterSource("umIds",umIds).addValue("plantId",plantId),this::mapTankaRow));
        if(tmIds!=null&&!tmIds.isEmpty())for(var row:searchTankaInsert(tmIds,plantId))if(result.stream().noneMatch(old->old.getTmvTmTankaId().compareTo(row.getTmvTmTankaId())==0&&old.getBrandkoseiId().compareTo(row.getBrandkoseiId())==0&&java.util.Objects.equals(old.getKaisiDt(),row.getKaisiDt())))result.add(row);
        return result;
    }

    // ================================================================
    // 12. カスタマー契約時間ID取得
    // 【変換元】Mcm2002uScreen.vb - getUmKeiyakujikanId()
    //   元コード: SELECT DISTINCT TMG.TM_KEIYAKUJIKAN_ID
    //             FROM MCM_UM_MITSUMORI UMG
    //             INNER JOIN MCM_UM_TANKA UMF ON ...
    //             INNER JOIN MCM_TM_TANKA TMF ON ...
    //             INNER JOIN MCM_TM_KIKAN TMG ON ...
    // ================================================================
    public List<String> findUmKeiyakujikanIds(BigDecimal umKihonMitsumoriId) {

        String sql =
            "SELECT DISTINCT TMG.TM_KEIYAKUJIKAN_ID " +
            "FROM MCM.MCM_UM_MITSUMORI UMG " +
            "  INNER JOIN MCM.MCM_UM_TANKA UMF " +
            "    ON UMG.UM_MITSUMORI_ID = UMF.UM_MITSUMORI_ID " +
            "   AND UMG.UM_KIHON_MITSUMORI_ID = :umKihonMitsumoriId " +
            "  INNER JOIN MCM.MCM_TM_TANKA TMF ON UMF.TM_TANKA_ID = TMF.TM_TANKA_ID " +
            "  INNER JOIN MCM.MCM_TM_KIKAN TMG ON TMF.TM_KIKAN_ID = TMG.TM_KIKAN_ID";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("umKihonMitsumoriId", umKihonMitsumoriId);
        return jdbc.queryForList(sql, params, String.class);
    }

    // ================================================================
    // 13. 状態取得（JOTAI / SYOUNIN_JOTAI）
    // 【変換元】Mcm2002uScreen.vb - getJotai()
    //   元コード: SELECT UMA.JOTAI, UMA.SYOUNIN_JOTAI
    //             FROM MCM_UM_KIHON_MITSUMORI UMA
    //             WHERE UMA.UM_KIHON_MITSUMORI_ID = ...
    // ================================================================
    public List<Object> findJotai(BigDecimal umKihonMitsumoriId) {

        String sql =
            "SELECT UMA.JOTAI, UMA.SYOUNIN_JOTAI " +
            "FROM MCM.MCM_UM_KIHON_MITSUMORI UMA " +
            "WHERE UMA.UM_KIHON_MITSUMORI_ID = :umKihonMitsumoriId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("umKihonMitsumoriId", umKihonMitsumoriId);
        List<List<Object>> results = jdbc.query(sql, params, (rs, rowNum) -> {
            List<Object> row = new ArrayList<>();
            row.add(rs.getString("JOTAI"));
            row.add(rs.getString("SYOUNIN_JOTAI"));
            return row;
        });
        return results.isEmpty() ? List.of("", "") : results.get(0);
    }

    // ================================================================
    // 14. ブランド構成ID取得
    // 【変換元】Mcm2002uScreen.vb - getBrandKoseiId()
    //   元コード: SELECT DISTINCT MAD.BRANDKOSEI_ID
    //             FROM MCM_MA_BRANDKOSEI MAD, MCM_MA_KIKIKOSEI MAE
    //             WHERE MAD.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID
    //               AND MAE.PLANT_ID = ...
    // ================================================================
    public List<String> findBrandKoseiIds(BigDecimal plantId) {

        String sql =
            "SELECT DISTINCT MAD.BRANDKOSEI_ID " +
            "FROM MCM.MCM_MA_BRANDKOSEI MAD " +
            "  INNER JOIN MCM.MCM_MA_KIKIKOSEI MAE ON MAD.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID " +
            "WHERE MAE.PLANT_ID = :plantId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("plantId", plantId);
        return jdbc.queryForList(sql, params, String.class);
    }

    // ================================================================
    // 15. ブランド構成取得
    // 【変換元】Mcm2002uScreen.vb - getBrandKosei()
    //   元コード: SELECT MAC.BRAND_ID, MAC.BRAND_NK,
    //             MAD.BRANDKOSEI_ID, MAD.BRANDSYOSAI_NK
    //             FROM MCM_MA_BRAND MAC, MCM_MA_BRANDKOSEI MAD
    //             WHERE MAC.BRAND_ID = MAD.BRAND_ID
    //               AND MAD.BRANDKOSEI_ID = ...
    // ================================================================
    public List<Object> findBrandKosei(BigDecimal brandKoseiId) {

        String sql =
            "SELECT MAC.BRAND_ID, MAC.BRAND_NK, " +
            "       MAD.BRANDKOSEI_ID, MAD.BRANDSYOSAI_NK " +
            "FROM MCM.MCM_MA_BRAND MAC " +
            "  INNER JOIN MCM.MCM_MA_BRANDKOSEI MAD ON MAC.BRAND_ID = MAD.BRAND_ID " +
            "WHERE MAD.BRANDKOSEI_ID = :brandKoseiId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("brandKoseiId", brandKoseiId);
        List<List<Object>> results = jdbc.query(sql, params, (rs, rowNum) -> {
            List<Object> row = new ArrayList<>();
            row.add(rs.getBigDecimal("BRAND_ID"));
            row.add(rs.getString("BRAND_NK"));
            row.add(rs.getBigDecimal("BRANDKOSEI_ID"));
            row.add(rs.getString("BRANDSYOSAI_NK"));
            return row;
        });
        return results.isEmpty() ? Collections.emptyList() : results.get(0);
    }

    // ================================================================
    // 16. コントローラ金額取得
    // 【変換元】Mcm2002uScreen.vb - getControllerKin()
    //   元コード: SELECT NVL(SUM(TMF.SIKIRI_KIN),0) ...
    // ================================================================
    public BigDecimal findControllerKin(BigDecimal kikimeisaiId) {

        String sql =
            "SELECT ISNULL(SUM(TMF.SIKIRI_KIN), 0) " +
            "FROM MCM.MCM_TM_TANKA TMF " +
            "  INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD ON TMF.TM_KIKIMEISAI_ID = TMD.TM_KIKIMEISAI_ID " +
            "WHERE TMD.KIKIMEISAI_ID = :kikimeisaiId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("kikimeisaiId", kikimeisaiId);
        BigDecimal result = jdbc.queryForObject(sql, params, BigDecimal.class);
        return result != null ? result : BigDecimal.ZERO;
    }

    // ================================================================
    // 17. 取引先依頼NO取得（構成）
    // 【変換元】Mcm2002uScreen.vb - getTmIraiNoKosei()
    //   元コード: SELECT TMA.TM_IRAI_NO, TMC.TM_KIKIKOSEI_ID
    //             FROM MCM_TM_MITSUMORI TMA ...
    //             WHERE TMB.TM_KEIYAKUJIKAN_ID IN (...)
    //               AND TMC.KIKIKOSEI_ID = ...
    // ================================================================
    public List<List<Object>> findTmIraiNoByKosei(
            List<BigDecimal> tmKeiyakujikanIds, BigDecimal kikikoseiId) {

        String sql =
            "SELECT DISTINCT TMA.TM_IRAI_NO, TMC.TM_KIKIKOSEI_ID " +
            "FROM MCM.MCM_TM_MITSUMORI TMA " +
            "  INNER JOIN MCM.MCM_TM_KEIYAKUJIKAN TMB ON TMA.TM_IRAI_ID = TMB.TM_IRAI_ID " +
            "  INNER JOIN MCM.MCM_TM_KIKIKOSEI TMC    ON TMA.TM_IRAI_ID = TMC.TM_IRAI_ID " +
            "WHERE TMB.TM_KEIYAKUJIKAN_ID IN (:tmKeiyakujikanIds) " +
            "  AND TMC.KIKIKOSEI_ID = :kikikoseiId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("tmKeiyakujikanIds", tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty() ? List.of(BigDecimal.ZERO) : tmKeiyakujikanIds);
        params.addValue("kikikoseiId", kikikoseiId);
        return jdbc.query(sql, params, (rs, rowNum) -> {
            List<Object> row = new ArrayList<>();
            row.add(rs.getString("TM_IRAI_NO"));
            row.add(rs.getBigDecimal("TM_KIKIKOSEI_ID"));
            return row;
        });
    }

    // ================================================================
    // 18. 取引先依頼NO取得（明細）
    // 【変換元】Mcm2002uScreen.vb - getTmIraiNoMeisai()
    //   元コード: SELECT TMA.TM_IRAI_NO, TMD.TM_KIKIMEISAI_ID
    //             WHERE TMD.KIKIMEISAI_ID = ...
    // ================================================================
    public List<List<Object>> findTmIraiNoByMeisai(
            List<BigDecimal> tmKeiyakujikanIds, BigDecimal kikimeisaiId) {

        String sql =
            "SELECT DISTINCT TMA.TM_IRAI_NO, TMD.TM_KIKIMEISAI_ID " +
            "FROM MCM.MCM_TM_MITSUMORI TMA " +
            "  INNER JOIN MCM.MCM_TM_KEIYAKUJIKAN TMB ON TMA.TM_IRAI_ID = TMB.TM_IRAI_ID " +
            "  INNER JOIN MCM.MCM_TM_KIKIKOSEI TMC    ON TMA.TM_IRAI_ID = TMC.TM_IRAI_ID " +
            "  INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD   ON TMC.TM_KIKIKOSEI_ID = TMD.TM_KIKIKOSEI_ID " +
            "WHERE TMB.TM_KEIYAKUJIKAN_ID IN (:tmKeiyakujikanIds) " +
            "  AND TMD.KIKIMEISAI_ID = :kikimeisaiId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("tmKeiyakujikanIds", tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty() ? List.of(BigDecimal.ZERO) : tmKeiyakujikanIds);
        params.addValue("kikimeisaiId", kikimeisaiId);
        return jdbc.query(sql, params, (rs, rowNum) -> {
            List<Object> row = new ArrayList<>();
            row.add(rs.getString("TM_IRAI_NO"));
            row.add(rs.getBigDecimal("TM_KIKIMEISAI_ID"));
            return row;
        });
    }

    // ================================================================
    // 19. カスタマー見積ID取得
    // 【変換元】Mcm2002uScreen.vb - getMitsumoriId()
    //   元コード: SELECT UMF.UM_MITSUMORI_ID FROM MCM_UM_MITSUMORI UMF
    //             WHERE UMF.UM_KIHON_MITSUMORI_ID = ...
    //               AND UMF.KAISI_DT = ...
    // ================================================================
    public BigDecimal findMitsumoriId(BigDecimal umKihonMitsumoriId, LocalDate kaisiDt) {

        String sql =
            "SELECT UMF.UM_MITSUMORI_ID " +
            "FROM MCM.MCM_UM_MITSUMORI UMF " +
            "WHERE UMF.UM_KIHON_MITSUMORI_ID = :umKihonMitsumoriId " +
            "  AND UMF.KAISI_DT = :kaisiDt";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("umKihonMitsumoriId", umKihonMitsumoriId);
        params.addValue("kaisiDt", kaisiDt);
        List<BigDecimal> results = jdbc.queryForList(sql, params, BigDecimal.class);
        return results.isEmpty() ? BigDecimal.ZERO : results.get(0);
    }

    // ================================================================
    // 20. 開始日取得
    // 【変換元】Mcm2002uScreen.vb - getKaisiDt()
    //   元コード: SELECT MIN(UMG.KAISI_DT) AS KAISI_DT
    //             FROM MCM_UM_MITSUMORI UMG
    //             WHERE UMG.UM_KIHON_MITSUMORI_ID = ...
    // ================================================================
    public LocalDate findKaisiDt(BigDecimal umKihonMitsumoriId) {

        String sql =
            "SELECT MIN(UMG.KAISI_DT) AS KAISI_DT " +
            "FROM MCM.MCM_UM_MITSUMORI UMG " +
            "WHERE UMG.UM_KIHON_MITSUMORI_ID = :umKihonMitsumoriId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("umKihonMitsumoriId", umKihonMitsumoriId);
        LocalDate result = jdbc.queryForObject(sql, params, LocalDate.class);
        return result;
    }

    // ================================================================
    // 21. 開始期間一覧取得
    // 【変換元】Mcm2002uScreen.vb - getKaisiKikan()
    //   Oracle: NVL(TMG.SYURYO_DT,'9999/12/30') + 1
    //   → SQL Server: DATEADD(DAY, 1, ISNULL(TMG.SYURYO_DT, '9999-12-30'))
    //   Oracle: UMG.KAISI_DT(+) = TMG.KAISI_DT
    //   → SQL Server: LEFT JOIN ... ON UMG.KAISI_DT = TMG.KAISI_DT
    // ================================================================
    public List<List<Object>> findKaisiKikan(
            List<BigDecimal> tmKeiyakujikanIds,
            BigDecimal umKihonMitsumoriId,
            boolean isFukusei) {

        if (tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty()) {
            return Collections.emptyList();
        }

        String umMitsumoriIdCol = isFukusei
            ? "0 AS UM_MITSUMORI_ID"
            : "UMG.UM_MITSUMORI_ID";
        String umKihonMitsumoriIdCol = isFukusei
            ? "0 AS UM_KIHON_MITSUMORI_ID"
            : "UMG.UM_KIHON_MITSUMORI_ID";

        String sql =
            "SELECT TMG.KAISI_DT, " + umMitsumoriIdCol + ", " + umKihonMitsumoriIdCol + ", " +
            "       0 AS UMU_FLG " +
            "FROM ( " +
            "    SELECT UM_MITSUMORI_ID, UM_KIHON_MITSUMORI_ID, KAISI_DT " +
            "    FROM MCM.MCM_UM_MITSUMORI " +
            "    WHERE UM_KIHON_MITSUMORI_ID = ISNULL(:umKihonMitsumoriId, 0) " +
            ") UMG " +
            "RIGHT JOIN ( " +
            "    SELECT DISTINCT TMG2.KAISI_DT AS KAISI_DT " +
            "    FROM MCM.MCM_TM_KIKAN TMG2 " +
            "    WHERE TMG2.TM_KEIYAKUJIKAN_ID IN (:tmKeiyakujikanIds) " +
            "    UNION " +
            "    SELECT DISTINCT DATEADD(DAY, 1, ISNULL(TMG3.SYURYO_DT, '9999-12-30')) AS KAISI_DT " +
            "    FROM MCM.MCM_TM_KIKAN TMG3 " +
            "    WHERE TMG3.TM_KEIYAKUJIKAN_ID IN (:tmKeiyakujikanIds) " +
            ") TMG ON UMG.KAISI_DT = TMG.KAISI_DT " +
            "ORDER BY TMG.KAISI_DT";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("tmKeiyakujikanIds", tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty() ? List.of(BigDecimal.ZERO) : tmKeiyakujikanIds);
        params.addValue("umKihonMitsumoriId", umKihonMitsumoriId);
        return jdbc.query(sql, params, (rs, rowNum) -> {
            List<Object> row = new ArrayList<>();
            row.add(getLocalDate(rs, "KAISI_DT"));
            row.add(getBigDecimal(rs, "UM_MITSUMORI_ID"));
            row.add(getBigDecimal(rs, "UM_KIHON_MITSUMORI_ID"));
            row.add(getBigDecimal(rs, "UMU_FLG"));
            return row;
        });
    }
}
