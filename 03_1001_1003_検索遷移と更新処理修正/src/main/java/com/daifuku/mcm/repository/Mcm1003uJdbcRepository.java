package com.daifuku.mcm.repository;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.namedparam.MapSqlParameterSource;
import org.springframework.jdbc.core.namedparam.NamedParameterJdbcTemplate;
import org.springframework.stereotype.Repository;

import java.math.BigDecimal;
import java.util.List;

/**
 * 【変換元】Mcm1003uDataSet.Designer.vb - 各TableAdapter（PLANT_ID / TK_KEIYAKU / TM_KEIYAKUJIKAN / TK_SIHARAIMEISAI / UM_TANKA）
 * 【変換元行数】約672,000行（DataSet.Designer.vb全体）
 * 【作成者】T.Kajimura (2008/02/14)
 *
 * MCM1003U 画面のビジネスロジックで使用する補助クエリ群。
 * 各メソッドは元VBのTableAdapterのFill/GetData相当。
 *
 * 【SQL変換】Oracle → SQL Server
 *   ROWNUM <= n → TOP n
 *   :param      → @param（NamedParameterJdbcTemplate使用のため :param のまま記述可）
 */
@Repository
public class Mcm1003uJdbcRepository {

    private static final Logger log = LoggerFactory.getLogger(Mcm1003uJdbcRepository.class);

    private final NamedParameterJdbcTemplate jdbcTemplate;

    public Mcm1003uJdbcRepository(NamedParameterJdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    /**
     * 契約時間IDに紐づくプラントIDを取得する。
     *
     * 【変換元】Mcm1003uDataSet.Designer.vb - PLANT_IDTableAdapter.InitCommandCollection()
     *   元コード:
     *     SELECT DISTINCT MAE.PLANT_ID
     *     FROM MCM_MA_KIKIKOSEI MAE, MCM_MA_KIKIMEISAI MAF,
     *          MCM_TM_KIKIMEISAI TMD, MCM_TM_KIKIKOSEI TMC,
     *          MCM_TM_MITSUMORI TMA, MCM_TM_KEIYAKUJIKAN TMB
     *     WHERE MAE.KIKIKOSEI_ID = MAF.KIKIKOSEI_ID
     *       AND MAF.KIKIMEISAI_ID = TMD.KIKIMEISAI_ID
     *       AND TMD.TM_KIKIKOSEI_ID = TMC.TM_KIKIKOSEI_ID
     *       AND TMC.TM_MITSUMORI_ID = TMA.TM_MITSUMORI_ID
     *       AND TMA.TM_KEIYAKUJIKAN_ID = TMB.TM_KEIYAKUJIKAN_ID
     *       AND TMB.TM_KEIYAKUJIKAN_ID = :TM_KEIYAKUJIKAN_ID
     *
     * @param tmKeiyakujikanId 契約時間ID
     * @return プラントIDリスト
     */
    public List<BigDecimal> findPlantIdsByKeiyakujikanId(BigDecimal tmKeiyakujikanId) {
        String sql =
                "SELECT DISTINCT MAE.PLANT_ID " +
                "FROM MCM.MCM_MA_KIKIKOSEI MAE " +
                "INNER JOIN MCM.MCM_MA_KIKIMEISAI MAF ON MAE.KIKIKOSEI_ID = MAF.KIKIKOSEI_ID " +
                "INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD ON MAF.KIKIMEISAI_ID = TMD.KIKIMEISAI_ID " +
                "INNER JOIN MCM.MCM_TM_KIKIKOSEI TMC ON TMD.TM_KIKIKOSEI_ID = TMC.TM_KIKIKOSEI_ID " +
                "INNER JOIN MCM.MCM_TM_MITSUMORI TMA ON TMC.TM_MITSUMORI_ID = TMA.TM_MITSUMORI_ID " +
                "INNER JOIN MCM.MCM_TM_KEIYAKUJIKAN TMB ON TMA.TM_KEIYAKUJIKAN_ID = TMB.TM_KEIYAKUJIKAN_ID " +
                "WHERE TMB.TM_KEIYAKUJIKAN_ID = :tmKeiyakujikanId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("tmKeiyakujikanId", tmKeiyakujikanId);

        return jdbcTemplate.queryForList(sql, params, BigDecimal.class);
    }

    /**
     * 契約IDの存在チェック。
     *
     * 【変換元】Mcm1003uDataSet.Designer.vb - TK_KEIYAKUTableAdapter.InitCommandCollection()
     *   元コード:
     *     SELECT COUNT(TK_KEIYAKU_ID)
     *     FROM MCM_TK_KEIYAKU
     *     WHERE TK_KEIYAKU_ID = :TK_KEIYAKU_ID AND ROWNUM <= 1
     *   → SQL Server: ROWNUM → TOP 1
     *
     * @param tkKeiyakuId 契約ID
     * @return 件数（0 or 1）
     */
    public int countTkKeiyakuById(BigDecimal tkKeiyakuId) {
        String sql =
                "SELECT TOP 1 COUNT(TK_KEIYAKU_ID) " +
                "FROM MCM.MCM_TK_KEIYAKU " +
                "WHERE TK_KEIYAKU_ID = :tkKeiyakuId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("tkKeiyakuId", tkKeiyakuId);

        Integer count = jdbcTemplate.queryForObject(sql, params, Integer.class);
        return count != null ? count : 0;
    }

    /**
     * 契約IDに紐づく支払明細の最小月を取得する。
     *
     * 【変換元】Mcm1003uDataSet.Designer.vb - TK_SIHARAIMEISAITableAdapter.InitCommandCollection()
     *   元コード:
     *     SELECT TKK.TSUKI
     *     FROM MCM_TK_SIHARAIMEISAI TKK
     *     JOIN MCM_TK_SIHARAI TKJ ON TKK.TK_SIHARAI_ID = TKJ.TK_SIHARAI_ID
     *     JOIN MCM_TK_KIKAN TKB ON TKJ.TK_KIKAN_ID = TKB.TK_KIKAN_ID
     *     JOIN MCM_TK_KEIYAKU TKA ON TKB.TK_KEIYAKU_ID = TKA.TK_KEIYAKU_ID
     *     WHERE TKA.TK_KEIYAKU_ID = :TK_KEIYAKU_ID
     *
     * @param tkKeiyakuId 契約ID
     * @return 月リスト（String）
     */
    public List<String> findTsukiByKeiyakuId(BigDecimal tkKeiyakuId) {
        String sql =
                "SELECT TKK.TSUKI " +
                "FROM MCM.MCM_TK_SIHARAIMEISAI TKK " +
                "INNER JOIN MCM.MCM_TK_SIHARAI TKJ ON TKK.TK_SIHARAI_ID = TKJ.TK_SIHARAI_ID " +
                "INNER JOIN MCM.MCM_TK_KIKAN TKB ON TKJ.TK_KIKAN_ID = TKB.TK_KIKAN_ID " +
                "INNER JOIN MCM.MCM_TK_KEIYAKU TKA ON TKB.TK_KEIYAKU_ID = TKA.TK_KEIYAKU_ID " +
                "WHERE TKA.TK_KEIYAKU_ID = :tkKeiyakuId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("tkKeiyakuId", tkKeiyakuId);

        return jdbcTemplate.queryForList(sql, params, String.class);
    }

    /**
     * 依頼IDに紐づくユーザ見積単価の件数を取得する（行削除チェック用）。
     *
     * 【変換元】Mcm1003uDataSet.Designer.vb - UM_TANKATableAdapter.InitCommandCollection()
     *   元コード:
     *     SELECT COUNT(*)
     *     FROM MCM_TM_KIKIMEISAI TMD, MCM_TM_KIKIKOSEI TMC,
     *          MCM_TM_TANKA TMF, MCM_UM_TANKA UMF
     *     WHERE TMD.TM_KIKIKOSEI_ID = TMC.TM_KIKIKOSEI_ID
     *       AND TMD.TM_KIKIMEISAI_ID = TMF.TM_KIKIMEISAI_ID
     *       AND TMF.TM_TANKA_ID = UMF.TM_TANKA_ID
     *       AND TMC.TM_IRAI_ID = :TM_IRAI_ID
     *
     * @param tmIraiId 依頼ID
     * @return ユーザ見積単価件数
     */
    public int countUmTankaByIraiId(BigDecimal tmIraiId) {
        String sql =
                "SELECT COUNT(*) " +
                "FROM MCM.MCM_TM_KIKIMEISAI TMD " +
                "INNER JOIN MCM.MCM_TM_KIKIKOSEI TMC ON TMD.TM_KIKIKOSEI_ID = TMC.TM_KIKIKOSEI_ID " +
                "INNER JOIN MCM.MCM_TM_TANKA TMF ON TMD.TM_KIKIMEISAI_ID = TMF.TM_KIKIMEISAI_ID " +
                "INNER JOIN MCM.MCM_UM_TANKA UMF ON TMF.TM_TANKA_ID = UMF.TM_TANKA_ID " +
                "WHERE TMC.TM_IRAI_ID = :tmIraiId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("tmIraiId", tmIraiId);

        Integer count = jdbcTemplate.queryForObject(sql, params, Integer.class);
        return count != null ? count : 0;
    }

    /**
     * 依頼担当者の重複なしリストを取得する。
     *
     * 【変換元】McmIraitantosyaDataTable_Sql.xml
     *   元コード:
     *     SELECT DISTINCT TMA.IRAITANTOSYA
     *     FROM MCM_TM_MITSUMORI TMA
     *     WHERE TMA.IRAITANTOSYA IS NOT NULL
     *
     * @return 依頼担当者リスト
     */
    public List<String> findDistinctIraitantosya() {
        String sql =
                "SELECT DISTINCT IRAITANTOSYA " +
                "FROM MCM.MCM_TM_MITSUMORI " +
                "WHERE IRAITANTOSYA IS NOT NULL " +
                "ORDER BY IRAITANTOSYA";
        return jdbcTemplate.queryForList(sql, new MapSqlParameterSource(), String.class);
    }
    /** VB McmLogin_AuthorityInfo.xml: function-specific authority, not the user's maximum. */
    public String findAuthority(String loginId,String functionId) {
        String sql="SELECT MAX(C.RIYOKENGEN_KBN) FROM MCM.MCM_MO_TANTO A " +
            "JOIN MCM.MCM_MO_TANTOKENGEN B ON B.TANTO_ID=A.TANTO_ID " +
            "JOIN MCM.MCM_MO_KENGENKOSEI C ON C.KENGENBUNRUI_ID=B.KENGENBUNRUI_ID AND C.RIYOKENGEN_KBN=B.RIYOKENGEN_KBN " +
            "JOIN MCM.MCM_MO_KENGENBUNRUI D ON D.KENGENBUNRUI_ID=C.KENGENBUNRUI_ID " +
            "JOIN MCM.MCM_MO_KINO E ON E.KINO_ID=C.KINO_ID WHERE A.LOGIN_ID=:loginId AND E.KINO_ID=:functionId";
        return jdbcTemplate.queryForObject(sql,new MapSqlParameterSource("loginId",loginId).addValue("functionId",functionId),String.class);
    }
    public boolean findMaintenance(String loginId) {
        var flags=jdbcTemplate.queryForList("SELECT MAINTENANCE_FLG FROM MCM.MCM_MO_TANTO WHERE LOGIN_ID=:loginId",new MapSqlParameterSource("loginId",loginId),String.class);
        return !flags.isEmpty()&&"1".equals(flags.get(0));
    }


    /** VB MCM_DEL_PAC.SP_TM。Serviceのトランザクション内でのみ実行する。 */
    public int deleteTorihikisaki(BigDecimal tmIraiId) {
        var params = new MapSqlParameterSource("tmIraiId", tmIraiId);
        var contractIds = jdbcTemplate.queryForList("""
            SELECT DISTINCT TKB.TK_KEIYAKU_ID
            FROM MCM.MCM_TK_KIKAN      TKB                                                INNER JOIN
            MCM.MCM_TK_KIKIKOSEI  TKC ON TKC.TK_KIKAN_ID      = TKB.TK_KIKAN_ID      INNER JOIN
            MCM.MCM_TK_KIKIMEISAI TKD ON TKD.TK_KIKIKOSEI_ID  = TKC.TK_KIKIKOSEI_ID  INNER JOIN
            MCM.MCM_TK_TANKA      TKF ON TKF.TK_KIKIMEISAI_ID = TKD.TK_KIKIMEISAI_ID INNER JOIN
            MCM.MCM_TM_TANKA      TMF ON TMF.TM_TANKA_ID      = TKF.TM_TANKA_ID      INNER JOIN
            MCM.MCM_TM_KIKIMEISAI TMD ON TMD.TM_KIKIMEISAI_ID = TMF.TM_KIKIMEISAI_ID INNER JOIN
            MCM.MCM_TM_KIKIKOSEI  TMC ON TMC.TM_KIKIKOSEI_ID  = TMD.TM_KIKIKOSEI_ID
            WHERE TMC.TM_IRAI_ID  = :tmIraiId
            """, params, BigDecimal.class);
        for (BigDecimal contractId : contractIds) {
            params.addValue("tkKeiyakuId", contractId);
            jdbcTemplate.update("""
                DELETE FROM MCM.MCM_TK_SIHARAIMEISAI
                WHERE EXISTS (SELECT *
                FROM MCM.MCM_TK_KIKAN   TKB                                      INNER JOIN MCM.MCM_TK_SIHARAI TKJ ON TKJ.TK_KIKAN_ID = TKB.TK_KIKAN_ID
                WHERE TKJ.TK_SIHARAI_ID = MCM_TK_SIHARAIMEISAI.TK_SIHARAI_ID
                AND TKB.TK_KEIYAKU_ID = :tkKeiyakuId)
                """, params);
            jdbcTemplate.update("""
                DELETE FROM MCM.MCM_TK_SIHARAI
                WHERE EXISTS (SELECT *
                FROM MCM.MCM_TK_KIKAN TKB
                WHERE TKB.TK_KIKAN_ID   = MCM_TK_SIHARAI.TK_KIKAN_ID
                AND TKB.TK_KEIYAKU_ID = :tkKeiyakuId)
                """, params);
            jdbcTemplate.update("""
                DELETE FROM MCM.MCM_TK_TENKENMEISAI
                WHERE EXISTS (SELECT *
                FROM MCM.MCM_TK_KIKAN     TKB                                              INNER JOIN MCM.MCM_TK_KIKIKOSEI TKC ON TKC.TK_KIKAN_ID     = TKB.TK_KIKAN_ID     INNER JOIN MCM.MCM_TK_TENKEN    TKH ON TKH.TK_KIKIKOSEI_ID = TKC.TK_KIKIKOSEI_ID
                WHERE TKH.TK_TENKEN_ID  = MCM_TK_TENKENMEISAI.TK_TENKEN_ID
                AND TKB.TK_KEIYAKU_ID = :tkKeiyakuId)
                """, params);
            jdbcTemplate.update("""
                DELETE FROM MCM.MCM_TK_TENKEN
                WHERE EXISTS (SELECT *
                FROM MCM.MCM_TK_KIKAN     TKB                                              INNER JOIN MCM.MCM_TK_KIKIKOSEI TKC ON TKC.TK_KIKAN_ID     = TKB.TK_KIKAN_ID
                WHERE TKC.TK_KIKIKOSEI_ID = MCM_TK_TENKEN.TK_KIKIKOSEI_ID
                AND TKB.TK_KEIYAKU_ID   = :tkKeiyakuId)
                """, params);
            jdbcTemplate.update("""
                DELETE FROM MCM.MCM_TK_TENPU
                WHERE EXISTS (SELECT *
                FROM MCM.MCM_TK_KIKAN TKB
                WHERE TKB.TK_KIKAN_ID   = MCM_TK_TENPU.TK_KIKAN_ID
                AND TKB.TK_KEIYAKU_ID = :tkKeiyakuId)
                """, params);
            jdbcTemplate.update("""
                DELETE FROM MCM.MCM_TK_TANKA
                WHERE EXISTS (SELECT *
                FROM MCM.MCM_TK_KIKAN      TKB                                              INNER JOIN MCM.MCM_TK_KIKIKOSEI  TKC ON TKC.TK_KIKAN_ID     = TKB.TK_KIKAN_ID     INNER JOIN MCM.MCM_TK_KIKIMEISAI TKD ON TKD.TK_KIKIKOSEI_ID = TKC.TK_KIKIKOSEI_ID
                WHERE TKD.TK_KIKIMEISAI_ID = MCM_TK_TANKA.TK_KIKIMEISAI_ID
                AND TKB.TK_KEIYAKU_ID    = :tkKeiyakuId)
                """, params);
            jdbcTemplate.update("""
                DELETE FROM MCM.MCM_TK_KOTAIMEISAI
                WHERE EXISTS (SELECT *
                FROM MCM.MCM_TK_KIKAN      TKB                                              INNER JOIN MCM.MCM_TK_KIKIKOSEI  TKC ON TKC.TK_KIKAN_ID     = TKB.TK_KIKAN_ID     INNER JOIN MCM.MCM_TK_KIKIMEISAI TKD ON TKD.TK_KIKIKOSEI_ID = TKC.TK_KIKIKOSEI_ID
                WHERE TKD.TK_KIKIMEISAI_ID = MCM_TK_KOTAIMEISAI.TK_KIKIMEISAI_ID
                AND TKB.TK_KEIYAKU_ID    = :tkKeiyakuId)
                """, params);
            jdbcTemplate.update("""
                DELETE FROM MCM.MCM_TK_KIKIMEISAI
                WHERE EXISTS (SELECT *
                FROM MCM.MCM_TK_KIKAN      TKB                                              INNER JOIN MCM.MCM_TK_KIKIKOSEI  TKC ON TKC.TK_KIKAN_ID     = TKB.TK_KIKAN_ID
                WHERE TKC.TK_KIKIKOSEI_ID = MCM_TK_KIKIMEISAI.TK_KIKIKOSEI_ID
                AND TKB.TK_KEIYAKU_ID   = :tkKeiyakuId)
                """, params);
            jdbcTemplate.update("""
                DELETE FROM MCM.MCM_TK_KIKIKOSEI
                WHERE EXISTS (SELECT *
                FROM MCM.MCM_TK_KIKAN TKB
                WHERE TKB.TK_KIKAN_ID   = MCM_TK_KIKIKOSEI.TK_KIKAN_ID
                AND TKB.TK_KEIYAKU_ID = :tkKeiyakuId)
                """, params);
            jdbcTemplate.update("""
                DELETE FROM MCM.MCM_TK_KIKAN
                WHERE MCM_TK_KIKAN.TK_KEIYAKU_ID = :tkKeiyakuId
                """, params);
            jdbcTemplate.update("""
                DELETE FROM MCM.MCM_TK_KEIYAKU
                WHERE MCM_TK_KEIYAKU.TK_KEIYAKU_ID = :tkKeiyakuId
                """, params);
        }
        jdbcTemplate.update("""
            DELETE FROM MCM.MCM_TM_TENKEN
            WHERE EXISTS (SELECT *
            FROM MCM.MCM_TM_KIKIKOSEI TMC
            WHERE TMC.TM_KIKIKOSEI_ID = MCM_TM_TENKEN.TM_KIKIKOSEI_ID
            AND TMC.TM_IRAI_ID      = :tmIraiId)
            """, params);
        jdbcTemplate.update("""
            DELETE FROM MCM.MCM_TM_KIKAN
            WHERE EXISTS (SELECT *
            FROM MCM.MCM_TM_KEIYAKUJIKAN TMB
            WHERE TMB.TM_KEIYAKUJIKAN_ID = MCM_TM_KIKAN.TM_KEIYAKUJIKAN_ID
            AND TMB.TM_IRAI_ID         = :tmIraiId)
            """, params);
        jdbcTemplate.update("""
            DELETE FROM MCM.MCM_TM_TANKA
            WHERE EXISTS (SELECT *
            FROM MCM.MCM_TM_KIKIKOSEI  TMC                                              INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD ON TMD.TM_KIKIKOSEI_ID = TMC.TM_KIKIKOSEI_ID
            WHERE TMD.TM_KIKIMEISAI_ID = MCM_TM_TANKA.TM_KIKIMEISAI_ID
            AND TMC.TM_IRAI_ID       = :tmIraiId)
            """, params);
        jdbcTemplate.update("""
            DELETE FROM MCM.MCM_TM_KOTAIMEISAI
            WHERE EXISTS (SELECT *
            FROM MCM.MCM_TM_KIKIKOSEI  TMC                                              INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD ON TMD.TM_KIKIKOSEI_ID = TMC.TM_KIKIKOSEI_ID
            WHERE TMD.TM_KIKIMEISAI_ID = MCM_TM_KOTAIMEISAI.TM_KIKIMEISAI_ID
            AND TMC.TM_IRAI_ID       = :tmIraiId)
            """, params);
        jdbcTemplate.update("""
            DELETE FROM MCM.MCM_TM_KIKIMEISAI
            WHERE EXISTS (SELECT *
            FROM MCM.MCM_TM_KIKIKOSEI TMC
            WHERE TMC.TM_KIKIKOSEI_ID = MCM_TM_KIKIMEISAI.TM_KIKIKOSEI_ID
            AND TMC.TM_IRAI_ID      = :tmIraiId)
            """, params);
        jdbcTemplate.update("""
            DELETE FROM MCM.MCM_TM_KIKIKOSEI
            WHERE MCM_TM_KIKIKOSEI.TM_IRAI_ID = :tmIraiId
            """, params);
        jdbcTemplate.update("""
            DELETE FROM MCM.MCM_TM_KEIYAKUJIKAN
            WHERE MCM_TM_KEIYAKUJIKAN.TM_IRAI_ID = :tmIraiId
            """, params);
        return jdbcTemplate.update("""
            DELETE FROM MCM.MCM_TM_MITSUMORI
            WHERE MCM_TM_MITSUMORI.TM_IRAI_ID = :tmIraiId
            """, params);
    }
}
