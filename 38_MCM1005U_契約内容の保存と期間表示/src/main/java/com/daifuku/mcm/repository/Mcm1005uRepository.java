/**
 * 【変換元】Mcm1005uDataSet.xsd（各TableAdapter）
 * 【説明】取引先契約内容（MCM1005U）のDBアクセスクラス
 *
 * MCM_TK_KEIYAKU / MCM_TK_KIKAN / MCM_TK_TENKEN / MCM_TK_SIHARAI /
 * MCM_TK_SIHARAIMEISAI / MCM_TK_TENPU の CRUD + MAX値取得
 * Oracle構文はSQL Server構文へ変換済み。
 */
package com.daifuku.mcm.repository;

import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.RowMapper;
import org.springframework.stereotype.Repository;

import com.daifuku.mcm.entity.McmTkKeiyakuEntity;
import com.daifuku.mcm.entity.McmTkKikanEntity;
import com.daifuku.mcm.entity.McmTkSiharaiEntity;
import com.daifuku.mcm.entity.McmTkTenkenEntity;
import com.daifuku.mcm.entity.McmTkTenpuEntity;

@Repository
public class Mcm1005uRepository {
    /** VB MCM_TK_KIKIMEISAI adapter, preserving left joins for historical master rows. */
    public List<Map<String, Object>> findReportDetails(BigDecimal periodId) {
        return jdbcTemplate.queryForList("""
            SELECT D.KIKIMEISAI_ID, C.KIKIKOSEI_ID, D.TK_KIKIKOSEI_ID,
                   D.KIKIHINMEI_NK, D.KIKIKATASHIKI, D.SURYO_NM, B.OYAKIKIBUNRUI_CD,
                   MC.HYOJIJUN AS MAE_HYOJIJUN, MD.HYOJIJUN
            FROM MCM.MCM_TK_KIKIMEISAI D
            JOIN MCM.MCM_TK_KIKIKOSEI C ON D.TK_KIKIKOSEI_ID=C.TK_KIKIKOSEI_ID
            JOIN MCM.MCM_MA_ATSUKAIKIKI A ON D.ATSUKAIKIKI_ID=A.ATSUKAIKIKI_ID
            JOIN MCM.MCM_MA_KIKIBUNRUI B ON A.KIKIBUNRUI_ID=B.KIKIBUNRUI_ID
            LEFT JOIN MCM.MCM_MA_KIKIMEISAI MD ON D.KIKIMEISAI_ID=MD.KIKIMEISAI_ID
            LEFT JOIN MCM.MCM_MA_KIKIKOSEI MC ON C.KIKIKOSEI_ID=MC.KIKIKOSEI_ID
            WHERE C.TK_KIKAN_ID=? ORDER BY MC.HYOJIJUN,MD.HYOJIJUN,D.TK_KIKIMEISAI_ID
            """, periodId);
    }

    public Map<String, Object> findReportContact(BigDecimal supplierId) {
        var rows = jdbcTemplate.queryForList("""
            SELECT A.TORIHIKISAKI_NK,B.TORISYUTANTOSYA_NK FROM MCM.MCM_MA_TORIHIKISAKI A
            JOIN MCM.MCM_MA_TORIMADOGUCHI B ON A.TORIHIKISAKI_ID=B.TORIHIKISAKI_ID
            WHERE A.TORIHIKISAKI_ID=? AND B.MADOGUCHI_KBN=2 AND B.YUKO_FLG=0
            """, supplierId);
        return rows.size() == 1 ? rows.get(0) : Map.of();
    }

    public String findReportPostalCode(BigDecimal periodId) {
        return jdbcTemplate.queryForObject("SELECT YUBIN_NO FROM MCM.MCM_TK_KIKAN WHERE TK_KIKAN_ID=?", String.class, periodId);
    }

    public List<String> findReportQuotationNumbers(BigDecimal periodId) {
        return jdbcTemplate.queryForList("""
            SELECT DISTINCT P.TM_MITSUMORI_NO FROM MCM.MCM_TM_KIKAN P
            JOIN MCM.MCM_TM_TANKA M ON P.TM_KIKAN_ID=M.TM_KIKAN_ID
            JOIN MCM.MCM_TK_TANKA T ON M.TM_TANKA_ID=T.TM_TANKA_ID
            JOIN MCM.MCM_TK_KIKIMEISAI D ON T.TK_KIKIMEISAI_ID=D.TK_KIKIMEISAI_ID
            JOIN MCM.MCM_TK_KIKIKOSEI C ON D.TK_KIKIKOSEI_ID=C.TK_KIKIKOSEI_ID
            WHERE C.TK_KIKAN_ID=? ORDER BY P.TM_MITSUMORI_NO
            """, String.class, periodId);
    }

    public List<String> findReportOrderNumbers(BigDecimal periodId) {
        return jdbcTemplate.queryForList("SELECT DISTINCT LTRIM(RTRIM(TEHAISEIBAN)) FROM MCM.MCM_TK_KIKIKOSEI WHERE TK_KIKAN_ID=? AND LTRIM(RTRIM(TEHAISEIBAN))<>'' ORDER BY 1", String.class, periodId);
    }

    @Autowired
    private JdbcTemplate jdbcTemplate;

    // ===================================================================
    // MCM_TK_KEIYAKU
    // ===================================================================

    private static final String SQL_SELECT_KEIYAKU =
            "SELECT TK_KEIYAKU_ID, IRAIJIGYOSYO_NK, IRAITANTOSYA, KEIYAKU_NO," +
            " KEIYAKU_DT, KEIYAKUKAISI_DT, KEIYAKUSYURYO_DT, JIDOKOSIN_FLG," +
            " JIKAIKOSIN_DT, KEIYAKUMANRYO_DT, ENTYOKEIYAKUMANRYO_DT," +
            " KAIYAKU_DT, JOTAI, SHONINJOTAI, PACK_FLG, PACKKEIYAKUNAIYO," +
            " PACKKEIYAKU_NO, BIKO, KOSINNAIYO, SHOKAI_KEIYAKU_DT," +
            " CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY" +
            " FROM MCM.MCM_TK_KEIYAKU WHERE TK_KEIYAKU_ID = ?";

    private static final String SQL_MAX_KEIYAKU_ID =
            "SELECT ISNULL(MAX(TK_KEIYAKU_ID), 0) FROM MCM.MCM_TK_KEIYAKU";

    private static final String SQL_INSERT_KEIYAKU =
            "INSERT INTO MCM.MCM_TK_KEIYAKU" +
            " (TK_KEIYAKU_ID, IRAIJIGYOSYO_NK, IRAITANTOSYA, KEIYAKU_NO," +
            "  KEIYAKU_DT, KEIYAKUKAISI_DT, KEIYAKUSYURYO_DT, JIDOKOSIN_FLG," +
            "  JIKAIKOSIN_DT, KEIYAKUMANRYO_DT, ENTYOKEIYAKUMANRYO_DT," +
            "  KAIYAKU_DT, JOTAI, SHONINJOTAI, PACK_FLG, PACKKEIYAKUNAIYO," +
            "  PACKKEIYAKU_NO, BIKO, KOSINNAIYO, SHOKAI_KEIYAKU_DT," +
            "  CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY)" +
            " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,GETDATE(),?,GETDATE(),?)";

    private static final String SQL_UPDATE_KEIYAKU =
            "UPDATE MCM.MCM_TK_KEIYAKU SET" +
            " IRAIJIGYOSYO_NK=?, IRAITANTOSYA=?, KEIYAKU_NO=?," +
            " KEIYAKU_DT=?, KEIYAKUKAISI_DT=?, KEIYAKUSYURYO_DT=?, JIDOKOSIN_FLG=?," +
            " JIKAIKOSIN_DT=?, KEIYAKUMANRYO_DT=?, ENTYOKEIYAKUMANRYO_DT=?," +
            " KAIYAKU_DT=?, JOTAI=?, SHONINJOTAI=?, PACK_FLG=?, PACKKEIYAKUNAIYO=?," +
            " PACKKEIYAKU_NO=?, BIKO=?, KOSINNAIYO=?, SHOKAI_KEIYAKU_DT=?," +
            " LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=?" +
            " WHERE TK_KEIYAKU_ID=?";

    public McmTkKeiyakuEntity findKeiyakuById(BigDecimal tkKeiyakuId) {
        List<McmTkKeiyakuEntity> list = jdbcTemplate.query(SQL_SELECT_KEIYAKU,
                new KeiyakuRowMapper(), tkKeiyakuId);
        return list.isEmpty() ? null : list.get(0);
    }

    public McmTkKeiyakuEntity lockKeiyaku(BigDecimal id) {
        List<McmTkKeiyakuEntity> rows = jdbcTemplate.query(SQL_SELECT_KEIYAKU.replace(
                "FROM MCM.MCM_TK_KEIYAKU WHERE", "FROM MCM.MCM_TK_KEIYAKU WITH (UPDLOCK,HOLDLOCK) WHERE"), new KeiyakuRowMapper(), id);
        return rows.isEmpty() ? null : rows.get(0);
    }

    /** VB UpdateAll: preserve detail IDs/history and use actual contract years, never year 2000. */
    public void saveMonthChecks(boolean inspection, BigDecimal parentId, LocalDateTime start, LocalDateTime end, boolean[] checked, String user) {
        // Only these two internal constants are used as SQL identifiers.
        String table = inspection ? "MCM_TK_TENKENMEISAI" : "MCM_TK_SIHARAIMEISAI";
        String parent = inspection ? "TK_TENKEN_ID" : "TK_SIHARAI_ID";
        String id = inspection ? "TK_TENKENMEISAI_ID" : "TK_SIHARAIMEISAI_ID";
        java.time.YearMonth from = java.time.YearMonth.from(start);
        LocalDateTime finish = java.time.temporal.ChronoUnit.MONTHS.between(from, java.time.YearMonth.from(end.plusDays(1))) < 12
                ? start.plusYears(1).minusDays(1) : end;
        java.time.YearMonth until = java.time.YearMonth.from(finish);
        for (int month = 1; month <= 12; month++) {
            if (!checked[month - 1]) {
                jdbcTemplate.update("UPDATE MCM." + table + " SET OFF_FLG=1,OFF_DT=GETDATE(),OFF_BY=?,LASTUPDATE_DT=GETDATE(),LASTUPDATE_BY=? WHERE "
                        + parent + "=? AND MONTH(TSUKI)=? AND ISNULL(OFF_FLG,0)=0", user, user, parentId, month);
            }
        }
        for (java.time.YearMonth month = from; !month.isAfter(until); month = month.plusMonths(1)) {
            if (!checked[month.getMonthValue() - 1]) continue;
            java.sql.Date date = java.sql.Date.valueOf(month.atDay(1));
            Integer exists = jdbcTemplate.queryForObject("SELECT COUNT(*) FROM MCM." + table + " WITH (UPDLOCK,HOLDLOCK) WHERE " + parent + "=? AND TSUKI=?", Integer.class, parentId, date);
            if (exists != null && exists > 0) {
                jdbcTemplate.update("UPDATE MCM." + table + " SET ON_FLG=1,ON_DT=GETDATE(),ON_BY=?,OFF_FLG=0,OFF_DT=NULL,OFF_BY=NULL,LASTUPDATE_DT=GETDATE(),LASTUPDATE_BY=? WHERE "
                        + parent + "=? AND TSUKI=? AND ISNULL(OFF_FLG,0)<>0", user, user, parentId, date);
            } else {
                jdbcTemplate.update("INSERT INTO MCM." + table + " (" + id + "," + parent + ",TSUKI,ON_FLG,ON_DT,ON_BY,OFF_FLG,CREATED_DT,CREATED_BY,LASTUPDATE_DT,LASTUPDATE_BY)"
                        + " SELECT ISNULL(MAX(" + id + "),0)+1,?,?,1,GETDATE(),?,0,GETDATE(),?,GETDATE(),? FROM MCM." + table + " WITH (UPDLOCK,HOLDLOCK)", parentId, date, user, user, user);
            }
        }
    }

    /** Mcm1005uDataSet.xsd: MCM_TK_KIKIJOHO, Oracle outer joins translated to SQL Server. */
    public List<java.util.Map<String,Object>> findKikiList(BigDecimal periodId) {
        return jdbcTemplate.queryForList("""
SELECT          TKC.TK_KIKIKOSEI_ID, TKC.TK_KIKAN_ID, TKC.KIKIKOSEI_ID, 
                      TKC.KIKIKOSEI_NK, TKC.SET_NM, TKC.TEHAISEIBAN, TKD.SEIZOMAKER_NK, 
                      TKD.KIKIHINMEI_NK, TKD.KIKIKATASHIKI, TKD.SURYO_NM, 
                      TKZ.TORIHOSYUJIKAN_ID, TKZ.TENKENUMU, TKZ.HOSYUHOHO, 
                      TKZ.SERVICEKEITAI, SUM(TKF.HYOJUN_KIN) AS HYOJUN_KIN, 
                      ISNULL(TKD.SURYO_NM, 0) * ISNULL(SUM(TKF.HYOJUN_KIN), 0) 
                      AS HYOJUNKEI_KIN, SUM(TKF.SIKIRI_KIN) AS SIKIRI_KIN, 
                      ISNULL(TKD.SURYO_NM, 0) * ISNULL(SUM(TKF.SIKIRI_KIN), 0) AS SIKIRIKEI_KIN, 
                      TKD.TK_KIKIMEISAI_ID, TKB.SUPPORT_ID AS TKB_SUPPORT_ID_OLD, 
                      MAB.SUPPORT_ID AS MAB_SUPPORT_ID, MAE.HYOJIJUN AS MAE_HYOJIJUN, 
                      MAF.HYOJIJUN AS MAF_HYOJIJUN, TKD.KIKIMEISAI_ID
FROM MCM.MCM_TK_KIKIKOSEI TKC
INNER JOIN MCM.MCM_TK_KIKIMEISAI TKD ON TKC.TK_KIKIKOSEI_ID=TKD.TK_KIKIKOSEI_ID
INNER JOIN MCM.MCM_TK_TANKA TKF ON TKD.TK_KIKIMEISAI_ID=TKF.TK_KIKIMEISAI_ID
INNER JOIN MCM.MCM_TK_KIKAN TKB ON TKC.TK_KIKAN_ID=TKB.TK_KIKAN_ID
INNER JOIN MCM.MCM_TK_TANKA TKZ ON TKF.TK_KIKIMEISAI_ID=TKZ.TK_KIKIMEISAI_ID AND TKB.KAISI_DT=TKZ.KAISI_DT
LEFT JOIN MCM.MCM_MA_KIKIKOSEI MAE ON TKC.KIKIKOSEI_ID=MAE.KIKIKOSEI_ID
INNER JOIN MCM.MCM_MA_PLANT MAB ON MAE.PLANT_ID=MAB.PLANT_ID
LEFT JOIN MCM.MCM_MA_KIKIMEISAI MAF ON TKD.KIKIMEISAI_ID=MAF.KIKIMEISAI_ID
WHERE TKC.TK_KIKAN_ID=?
GROUP BY     TKC.TK_KIKIKOSEI_ID, TKC.TK_KIKAN_ID, TKC.KIKIKOSEI_ID, 
                      TKC.KIKIKOSEI_NK, TKC.SET_NM, TKC.TEHAISEIBAN, TKD.SEIZOMAKER_NK, 
                      TKD.KIKIHINMEI_NK, TKD.KIKIKATASHIKI, TKD.SURYO_NM, 
                      TKZ.TORIHOSYUJIKAN_ID, TKZ.TENKENUMU, TKZ.HOSYUHOHO, 
                      TKZ.SERVICEKEITAI, TKD.TK_KIKIMEISAI_ID, TKB.SUPPORT_ID, 
                      MAB.SUPPORT_ID, MAE.HYOJIJUN, MAF.HYOJIJUN, TKD.KIKIMEISAI_ID, 
                      TKC.HYOJIJUN, TKD.HYOJIJUN
ORDER BY     TKC.HYOJIJUN, TKD.HYOJIJUN, MAE_HYOJIJUN, MAF_HYOJIJUN
                """, periodId);
    }

    public BigDecimal getMaxKeiyakuId() {
        BigDecimal max = jdbcTemplate.queryForObject(SQL_MAX_KEIYAKU_ID, BigDecimal.class);
        return max == null ? BigDecimal.ZERO : max;
    }

    public void insertKeiyaku(McmTkKeiyakuEntity e, String loginUser) {
        jdbcTemplate.update(SQL_INSERT_KEIYAKU,
                e.getTkKeiyakuId(), e.getIraijigyosyoNk(), e.getIraitantosya(), e.getKeiyakuNo(),
                e.getKeiyakuDt(), e.getKeiyakukaisiDt(), e.getKeiyakusyuryoDt(), e.getJidokosinFlg(),
                e.getJikaikosinDt(), e.getKeiyakumanryoDt(), e.getEntyokeiyakumanryoDt(),
                e.getKaiyakuDt(), e.getJotai(), e.getShoninjotai(),
                e.getPackFlg(), e.getPackkeiyakunaiyo(), e.getPackkeiyakuNo(),
                e.getBiko(), e.getKosinnaiyo(), e.getShokaiKeiyakuDt(),
                loginUser, loginUser);
    }

    public void updateKeiyaku(McmTkKeiyakuEntity e, String loginUser) {
        jdbcTemplate.update(SQL_UPDATE_KEIYAKU,
                e.getIraijigyosyoNk(), e.getIraitantosya(), e.getKeiyakuNo(),
                e.getKeiyakuDt(), e.getKeiyakukaisiDt(), e.getKeiyakusyuryoDt(), e.getJidokosinFlg(),
                e.getJikaikosinDt(), e.getKeiyakumanryoDt(), e.getEntyokeiyakumanryoDt(),
                e.getKaiyakuDt(), e.getJotai(), e.getShoninjotai(),
                e.getPackFlg(), e.getPackkeiyakunaiyo(), e.getPackkeiyakuNo(),
                e.getBiko(), e.getKosinnaiyo(), e.getShokaiKeiyakuDt(),
                loginUser, e.getTkKeiyakuId());
    }

    public void updateKeiyakuShoninjotai(BigDecimal tkKeiyakuId, String shoninjotai, String loginUser) {
        jdbcTemplate.update(
                "UPDATE MCM.MCM_TK_KEIYAKU SET SHONINJOTAI=?, LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=?" +
                " WHERE TK_KEIYAKU_ID=?",
                shoninjotai, loginUser, tkKeiyakuId);
    }

    // ===================================================================
    // MCM_TK_KIKAN
    // ===================================================================

    private static final String SQL_SELECT_KIKAN =
            "SELECT TK_KIKAN_ID, TK_KEIYAKU_ID, KAISI_DT, SYURYO_DT," +
            " HYOJUNGOKEI_KIN, SIKIRISYOKEI_KIN, SYUSSEINEBIKI_KIN, SIKIRIGOKEI_KIN," +
            " NONYUSAKI_ID, NONYUSAKI_CD, NONYUSAKI_NK," +
            " NONYUSAKIJUSYO1_NK, NONYUSAKIJUSYO2_NK," +
            " PLANT_ID, SUPPORT_ID, PLANT_NK," +
            " NONYUBUSYO_NK, NONYUTANTOSYA_NK, NONYUTEL_NO, NONYUFAX_NO, BIKO," +
            " TORIHIKISAKI_ID, TORIHIKISAKI_CD, TORIHIKISAKI_NK," +
            " TORITEL_NO, TORIFAX_NO, TORIJIGYOSYO_NK," +
            " TORISYUTANTOSYA_NK, TORIASSISTANT_NK," +
            " KEIYAKUJIKANTAI, HOSYUHOHO, YUKO_FLG, KEIYAKU_NO," +
            " CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY" +
            " FROM MCM.MCM_TK_KIKAN WHERE TK_KEIYAKU_ID = ? ORDER BY KAISI_DT";

    private static final String SQL_MAX_KIKAN_ID =
            "SELECT ISNULL(MAX(TK_KIKAN_ID), 0) FROM MCM.MCM_TK_KIKAN";

    private static final String SQL_INSERT_KIKAN =
            "INSERT INTO MCM.MCM_TK_KIKAN" +
            " (TK_KIKAN_ID, TK_KEIYAKU_ID, KAISI_DT, SYURYO_DT," +
            "  HYOJUNGOKEI_KIN, SIKIRISYOKEI_KIN, SYUSSEINEBIKI_KIN, SIKIRIGOKEI_KIN," +
            "  NONYUSAKI_ID, NONYUSAKI_CD, NONYUSAKI_NK," +
            "  NONYUSAKIJUSYO1_NK, NONYUSAKIJUSYO2_NK," +
            "  PLANT_ID, SUPPORT_ID, PLANT_NK," +
            "  NONYUBUSYO_NK, NONYUTANTOSYA_NK, NONYUTEL_NO, NONYUFAX_NO, BIKO," +
            "  TORIHIKISAKI_ID, TORIHIKISAKI_CD, TORIHIKISAKI_NK," +
            "  TORITEL_NO, TORIFAX_NO, TORIJIGYOSYO_NK," +
            "  TORISYUTANTOSYA_NK, TORIASSISTANT_NK," +
            "  KEIYAKUJIKANTAI, HOSYUHOHO, YUKO_FLG, KEIYAKU_NO," +
            "  CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY)" +
            " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,GETDATE(),?,GETDATE(),?)";

    private static final String SQL_UPDATE_KIKAN =
            "UPDATE MCM.MCM_TK_KIKAN SET" +
            " KAISI_DT=?, SYURYO_DT=?," +
            " HYOJUNGOKEI_KIN=?, SIKIRISYOKEI_KIN=?, SYUSSEINEBIKI_KIN=?, SIKIRIGOKEI_KIN=?," +
            " NONYUSAKI_ID=?, NONYUSAKI_CD=?, NONYUSAKI_NK=?," +
            " NONYUSAKIJUSYO1_NK=?, NONYUSAKIJUSYO2_NK=?," +
            " PLANT_ID=?, SUPPORT_ID=?, PLANT_NK=?," +
            " NONYUBUSYO_NK=?, NONYUTANTOSYA_NK=?, NONYUTEL_NO=?, NONYUFAX_NO=?, BIKO=?," +
            " TORIHIKISAKI_ID=?, TORIHIKISAKI_CD=?, TORIHIKISAKI_NK=?," +
            " TORITEL_NO=?, TORIFAX_NO=?, TORIJIGYOSYO_NK=?," +
            " TORISYUTANTOSYA_NK=?, TORIASSISTANT_NK=?," +
            " KEIYAKUJIKANTAI=?, HOSYUHOHO=?, YUKO_FLG=?, KEIYAKU_NO=?," +
            " LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=?" +
            " WHERE TK_KIKAN_ID=?";

    public List<McmTkKikanEntity> findKikanByKeiyakuId(BigDecimal tkKeiyakuId) {
        return jdbcTemplate.query(SQL_SELECT_KIKAN, new KikanRowMapper(), tkKeiyakuId);
    }

    public List<McmTkKikanEntity> lockKikanList(BigDecimal contractId) {
        return jdbcTemplate.query(SQL_SELECT_KIKAN.replace("MCM.MCM_TK_KIKAN WHERE",
                "MCM.MCM_TK_KIKAN WITH (UPDLOCK,HOLDLOCK) WHERE"), new KikanRowMapper(), contractId);
    }

    public List<McmTkTenkenEntity> lockTenkenList(BigDecimal periodId) {
        return jdbcTemplate.query(SQL_SELECT_TENKEN.replace("MCM.MCM_TK_TENKEN TKA",
                "MCM.MCM_TK_TENKEN TKA WITH (UPDLOCK,HOLDLOCK)"), new TenkenRowMapper(), periodId);
    }

    public McmTkSiharaiEntity lockSiharai(BigDecimal periodId) {
        List<McmTkSiharaiEntity> rows = jdbcTemplate.query(SQL_SELECT_SIHARAI.replace("MCM.MCM_TK_SIHARAI WHERE",
                "MCM.MCM_TK_SIHARAI WITH (UPDLOCK,HOLDLOCK) WHERE"), new SiharaiRowMapper(), periodId);
        return rows.isEmpty() ? null : rows.get(0);
    }

    public BigDecimal getMaxKikanId() {
        BigDecimal max = jdbcTemplate.queryForObject(SQL_MAX_KIKAN_ID, BigDecimal.class);
        return max == null ? BigDecimal.ZERO : max;
    }

    public void insertKikan(McmTkKikanEntity e, String loginUser) {
        jdbcTemplate.update(SQL_INSERT_KIKAN,
                e.getTkKikanId(), e.getTkKeiyakuId(), e.getKaisiDt(), e.getSyuryoDt(),
                e.getHyojungokeiKin(), e.getSikirisyokeiKin(), e.getSyusseinebikiKin(), e.getSikirigokeiKin(),
                e.getNonyusakiId(), e.getNonyusakiCd(), e.getNonyusakiNk(),
                e.getNonyusakijusyo1Nk(), e.getNonyusakijusyo2Nk(),
                e.getPlantId(), e.getSupportId(), e.getPlantNk(),
                e.getNonyubusyoNk(), e.getNonyutantosyaNk(), e.getNonyutelNo(), e.getNonyufaxNo(), e.getBiko(),
                e.getTorihikisakiId(), e.getTorihikisakiCd(), e.getTorihikisakiNk(),
                e.getToritelNo(), e.getTorifaxNo(), e.getTorijigyosyoNk(),
                e.getTorisyutantosyaNk(), e.getToriassistantNk(),
                e.getKeiyakujikantai(), e.getHosyuhoho(), e.getYukoFlg(), e.getKeiyakuNo(),
                loginUser, loginUser);
    }

    public void updateKikan(McmTkKikanEntity e, String loginUser) {
        jdbcTemplate.update(SQL_UPDATE_KIKAN,
                e.getKaisiDt(), e.getSyuryoDt(),
                e.getHyojungokeiKin(), e.getSikirisyokeiKin(), e.getSyusseinebikiKin(), e.getSikirigokeiKin(),
                e.getNonyusakiId(), e.getNonyusakiCd(), e.getNonyusakiNk(),
                e.getNonyusakijusyo1Nk(), e.getNonyusakijusyo2Nk(),
                e.getPlantId(), e.getSupportId(), e.getPlantNk(),
                e.getNonyubusyoNk(), e.getNonyutantosyaNk(), e.getNonyutelNo(), e.getNonyufaxNo(), e.getBiko(),
                e.getTorihikisakiId(), e.getTorihikisakiCd(), e.getTorihikisakiNk(),
                e.getToritelNo(), e.getTorifaxNo(), e.getTorijigyosyoNk(),
                e.getTorisyutantosyaNk(), e.getToriassistantNk(),
                e.getKeiyakujikantai(), e.getHosyuhoho(), e.getYukoFlg(), e.getKeiyakuNo(),
                loginUser, e.getTkKikanId());
    }

    public void updateKikanYukoFlg(BigDecimal tkKikanId, BigDecimal yukoFlg, String loginUser) {
        jdbcTemplate.update(
                "UPDATE MCM.MCM_TK_KIKAN SET YUKO_FLG=?, LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=?" +
                " WHERE TK_KIKAN_ID=?",
                yukoFlg, loginUser, tkKikanId);
    }

    public void updateKikanShoninjotai(BigDecimal tkKikanId, BigDecimal yukoFlg, String loginUser) {
        jdbcTemplate.update(
                "UPDATE MCM.MCM_TK_KIKAN SET YUKO_FLG=?, LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=?" +
                " WHERE TK_KIKAN_ID=?",
                yukoFlg, loginUser, tkKikanId);
    }

    // ===================================================================
    // MCM_TK_TENKEN
    // ===================================================================

    private static final String SQL_SELECT_TENKEN =
            "SELECT TKA.TK_TENKEN_ID, TKA.TK_KIKIKOSEI_ID, TKA.TENKENKAISU," +
            " TKA.TENKENKANOYOBI, TKA.YAKANTAIOUMU, TKA.HOSHUGAISHA_ID," +
            " TKA.YAKANHOSHUGAISHA_ID, TKA.OYAKIKIBUNRUI_CD, TKA.BIKO," +
            " TKA.CREATED_DT, TKA.CREATED_BY, TKA.LASTUPDATE_DT, TKA.LASTUPDATE_BY," +
            " TKC.TK_KIKAN_ID, TKC.KIKIKOSEI_NK" +
            " FROM MCM.MCM_TK_TENKEN TKA" +
            " INNER JOIN MCM.MCM_TK_KIKIKOSEI TKC ON TKA.TK_KIKIKOSEI_ID = TKC.TK_KIKIKOSEI_ID" +
            " WHERE TKC.TK_KIKAN_ID = ?" +
            " ORDER BY TKA.TK_TENKEN_ID";

    private static final String SQL_MAX_TENKEN_ID =
            "SELECT ISNULL(MAX(TK_TENKEN_ID), 0) FROM MCM.MCM_TK_TENKEN";

    private static final String SQL_INSERT_TENKEN =
            "INSERT INTO MCM.MCM_TK_TENKEN" +
            " (TK_TENKEN_ID, TK_KIKIKOSEI_ID, TENKENKAISU, TENKENKANOYOBI," +
            "  YAKANTAIOUMU, HOSHUGAISHA_ID, YAKANHOSHUGAISHA_ID, OYAKIKIBUNRUI_CD," +
            "  BIKO, CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY)" +
            " VALUES (?,?,?,?,?,?,?,?,?,GETDATE(),?,GETDATE(),?)";

    private static final String SQL_UPDATE_TENKEN =
            "UPDATE MCM.MCM_TK_TENKEN SET" +
            " TENKENKAISU=?, TENKENKANOYOBI=?, YAKANTAIOUMU=?," +
            " HOSHUGAISHA_ID=?, YAKANHOSHUGAISHA_ID=?, OYAKIKIBUNRUI_CD=?," +
            " BIKO=?, LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=?" +
            " WHERE TK_TENKEN_ID=? AND ISNULL(OFF_FLG,0)=0";

    private static final String SQL_DELETE_TENKEN_BY_KIKAN =
            "DELETE FROM MCM.MCM_TK_TENKEN WHERE TK_KIKIKOSEI_ID IN" +
            " (SELECT TK_KIKIKOSEI_ID FROM MCM_TK_KIKIKOSEI WHERE TK_KIKAN_ID=?)";

    public List<McmTkTenkenEntity> findTenkenByKikanId(BigDecimal tkKikanId) {
        return jdbcTemplate.query(SQL_SELECT_TENKEN, new TenkenRowMapper(), tkKikanId);
    }

    public BigDecimal getMaxTenkenId() {
        BigDecimal max = jdbcTemplate.queryForObject(SQL_MAX_TENKEN_ID, BigDecimal.class);
        return max == null ? BigDecimal.ZERO : max;
    }

    public void insertTenken(McmTkTenkenEntity e, String loginUser) {
        jdbcTemplate.update(SQL_INSERT_TENKEN,
                e.getTkTenkenId(), e.getTkKikokoseiId(), e.getTenkenkaisu(), e.getTenkenkanoyobi(),
                e.getYakantaioumu(), e.getHoshugaishaId(), e.getYakanhoshugaishaId(), e.getOyakikibunruiCd(),
                e.getBiko(), loginUser, loginUser);
    }

    public void updateTenken(McmTkTenkenEntity e, String loginUser) {
        jdbcTemplate.update(SQL_UPDATE_TENKEN,
                e.getTenkenkaisu(), e.getTenkenkanoyobi(), e.getYakantaioumu(),
                e.getHoshugaishaId(), e.getYakanhoshugaishaId(), e.getOyakikibunruiCd(),
                e.getBiko(), loginUser, e.getTkTenkenId());
    }

    public void deleteTenkenByKikanId(BigDecimal tkKikanId) {
        jdbcTemplate.update(SQL_DELETE_TENKEN_BY_KIKAN, tkKikanId);
    }

    // ===================================================================
    // MCM_TK_TENKENMEISAI
    // ===================================================================

    private static final String SQL_SELECT_TSUKI_CHECK =
            "SELECT MONTH(TSUKI) AS TSUKI_MM" +
            " FROM MCM.MCM_TK_TENKENMEISAI WHERE TK_TENKEN_ID=? AND ISNULL(OFF_FLG,0)=0";

    /** 点検明細から月チェック済みの月番号(1〜12)リストを返す */
    public List<Integer> findTenkenMeisaiTsukiMonths(BigDecimal tkTenkenId) {
        return jdbcTemplate.queryForList(SQL_SELECT_TSUKI_CHECK, Integer.class, tkTenkenId);
    }

    // ===================================================================
    // MCM_TK_SIHARAI
    // ===================================================================

    private static final String SQL_SELECT_SIHARAI =
            "SELECT TK_SIHARAI_ID, TK_KIKAN_ID, KAISU, BIKO," +
            " CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY" +
            " FROM MCM.MCM_TK_SIHARAI WHERE TK_KIKAN_ID=?";

    private static final String SQL_MAX_SIHARAI_ID =
            "SELECT ISNULL(MAX(TK_SIHARAI_ID), 0) FROM MCM.MCM_TK_SIHARAI";

    private static final String SQL_INSERT_SIHARAI =
            "INSERT INTO MCM.MCM_TK_SIHARAI" +
            " (TK_SIHARAI_ID, TK_KIKAN_ID, KAISU, BIKO, CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY)" +
            " VALUES (?,?,?,?,GETDATE(),?,GETDATE(),?)";

    private static final String SQL_UPDATE_SIHARAI =
            "UPDATE MCM.MCM_TK_SIHARAI SET KAISU=?, BIKO=?, LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=?" +
            " WHERE TK_SIHARAI_ID=?";

    private static final String SQL_DELETE_SIHARAI_BY_KIKAN =
            "DELETE FROM MCM.MCM_TK_SIHARAI WHERE TK_KIKAN_ID=?";

    public McmTkSiharaiEntity findSiharaiByKikanId(BigDecimal tkKikanId) {
        List<McmTkSiharaiEntity> list = jdbcTemplate.query(SQL_SELECT_SIHARAI,
                new SiharaiRowMapper(), tkKikanId);
        return list.isEmpty() ? null : list.get(0);
    }

    public BigDecimal getMaxSiharaiId() {
        BigDecimal max = jdbcTemplate.queryForObject(SQL_MAX_SIHARAI_ID, BigDecimal.class);
        return max == null ? BigDecimal.ZERO : max;
    }

    public void insertSiharai(McmTkSiharaiEntity e, String loginUser) {
        jdbcTemplate.update(SQL_INSERT_SIHARAI,
                e.getTkSiharaiId(), e.getTkKikanId(), e.getKaisu(), e.getBiko(),
                loginUser, loginUser);
    }

    public void updateSiharai(McmTkSiharaiEntity e, String loginUser) {
        jdbcTemplate.update(SQL_UPDATE_SIHARAI,
                e.getKaisu(), e.getBiko(), loginUser, e.getTkSiharaiId());
    }

    public void deleteSiharaiByKikanId(BigDecimal tkKikanId) {
        jdbcTemplate.update(SQL_DELETE_SIHARAI_BY_KIKAN, tkKikanId);
    }

    /** MCM_TK_SIHARAIMEISAI から月チェック済みの月番号(1〜12)リストを返す */
    public List<Integer> findSiharaiMeisaiTsukiMonths(BigDecimal tkSiharaiId) {
        return jdbcTemplate.queryForList(
                "SELECT MONTH(TSUKI) AS TSUKI_MM FROM MCM.MCM_TK_SIHARAIMEISAI WHERE TK_SIHARAI_ID=? AND ISNULL(OFF_FLG,0)=0",
                Integer.class, tkSiharaiId);
    }

    /** MCM_TK_SIHARAIMEISAI を TK_SIHARAI_ID で全削除して再INSERT */
    public void deleteSiharaimeisaiBySiharaiId(BigDecimal tkSiharaiId) {
        jdbcTemplate.update("DELETE FROM MCM.MCM_TK_SIHARAIMEISAI WHERE TK_SIHARAI_ID=?", tkSiharaiId);
    }

    public BigDecimal getMaxSiharaimeisaiId() {
        BigDecimal max = jdbcTemplate.queryForObject(
                "SELECT ISNULL(MAX(TK_SIHARAIMEISAI_ID), 0) FROM MCM.MCM_TK_SIHARAIMEISAI", BigDecimal.class);
        return max == null ? BigDecimal.ZERO : max;
    }

    /** 支払月(1〜12)に対応する TSUKI 日付（月初）でSIHARAIMEISAI INSERT */
    public void insertSiharaimeisai(BigDecimal tkSiharaimeisaiId, BigDecimal tkSiharaiId,
                                    int month, String loginUser) {
        String tsukiVal = String.format("2000-%02d-01", month);
        jdbcTemplate.update(
                "INSERT INTO MCM.MCM_TK_SIHARAIMEISAI" +
                " (TK_SIHARAIMEISAI_ID, TK_SIHARAI_ID, TSUKI, ON_FLG, OFF_FLG," +
                "  CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY)" +
                " VALUES (?,?,?,1,0,GETDATE(),?,GETDATE(),?)",
                tkSiharaimeisaiId, tkSiharaiId, tsukiVal, loginUser, loginUser);
    }

    // ===================================================================
    // MCM_TK_TENPU
    // ===================================================================

    private static final String SQL_SELECT_TENPU =
            "SELECT TK_TENPU_ID, TK_KIKAN_ID, TENPUFILE_NK, DIRECTORY, SHONINJOTAI," +
            " BIKO, CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY" +
            " FROM MCM.MCM_TK_TENPU WHERE TK_KIKAN_ID=?" +
            " ORDER BY TK_TENPU_ID";

    private static final String SQL_MAX_TENPU_ID =
            "SELECT ISNULL(MAX(TK_TENPU_ID), 0) FROM MCM.MCM_TK_TENPU";

    private static final String SQL_INSERT_TENPU =
            "INSERT INTO MCM.MCM_TK_TENPU" +
            " (TK_TENPU_ID, TK_KIKAN_ID, TENPUFILE_NK, DIRECTORY, SHONINJOTAI, BIKO," +
            "  CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY)" +
            " VALUES (?,?,?,?,?,?,GETDATE(),?,GETDATE(),?)";

    private static final String SQL_UPDATE_TENPU =
            "UPDATE MCM.MCM_TK_TENPU SET TENPUFILE_NK=?, DIRECTORY=?, SHONINJOTAI=?, BIKO=?," +
            " LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? WHERE TK_TENPU_ID=?";

    private static final String SQL_DELETE_TENPU =
            "DELETE FROM MCM.MCM_TK_TENPU WHERE TK_TENPU_ID=?";

    private static final String SQL_UPDATE_TENPU_SHONINJOTAI_BY_KIKAN =
            "UPDATE MCM.MCM_TK_TENPU SET SHONINJOTAI=?, LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=?" +
            " WHERE TK_KIKAN_ID=? AND SHONINJOTAI IN (?,?)";

    public List<McmTkTenpuEntity> findTenpuByKikanId(BigDecimal tkKikanId) {
        return jdbcTemplate.query(SQL_SELECT_TENPU, new TenpuRowMapper(), tkKikanId);
    }

    public BigDecimal getMaxTenpuId() {
        BigDecimal max = jdbcTemplate.queryForObject(SQL_MAX_TENPU_ID, BigDecimal.class);
        return max == null ? BigDecimal.ZERO : max;
    }

    public void insertTenpu(McmTkTenpuEntity e, String loginUser) {
        jdbcTemplate.update(SQL_INSERT_TENPU,
                e.getTkTenpuId(), e.getTkKikanId(), e.getTenpufileNk(), e.getDirectory(),
                e.getShoninjotai(), e.getBiko(), loginUser, loginUser);
    }

    public BigDecimal nextTenpuId() {
        return jdbcTemplate.queryForObject("SELECT ISNULL(MAX(TK_TENPU_ID),0)+1 FROM MCM.MCM_TK_TENPU WITH (UPDLOCK,HOLDLOCK)", BigDecimal.class);
    }

    public void updateTenpu(McmTkTenpuEntity e, String loginUser) {
        jdbcTemplate.update(SQL_UPDATE_TENPU,
                e.getTenpufileNk(), e.getDirectory(), e.getShoninjotai(), e.getBiko(),
                loginUser, e.getTkTenpuId());
    }

    public void deleteTenpu(BigDecimal tkTenpuId) {
        jdbcTemplate.update(SQL_DELETE_TENPU, tkTenpuId);
    }

    /** 申請処理：作成中/差戻しの添付を「審査中」に更新 */
    public void updateTenpuShoninjotaiToShinsachu(BigDecimal tkKikanId, String loginUser) {
        jdbcTemplate.update(SQL_UPDATE_TENPU_SHONINJOTAI_BY_KIKAN,
                "1", loginUser, tkKikanId, "0", "4");
    }

    // ===================================================================
    // MCM_TK_KIKIKOSEI / MCM_TK_KIKIMEISAI / MCM_TK_KOTAIMEISAI / MCM_TK_TANKA
    // 【変換元】Mcm1005uTabControl.vb — setKikiKoseiDataTable/setKikiMeisaiDataTable 等
    // ===================================================================

    /** 期間に紐づく TM_KEIYAKUJIKAN_ID を取得（MCM1006U 変更ボタン用） */
    public List<BigDecimal> findTmKeiyakujikanIdsByKikanId(BigDecimal tkKikanId) {
        return jdbcTemplate.queryForList(
            "SELECT DISTINCT mt.TM_KEIYAKUJIKAN_ID" +
            " FROM MCM.MCM_TK_TANKA t" +
            " INNER JOIN MCM_TM_TANKA mt ON t.TM_TANKA_ID = mt.TM_TANKA_ID" +
            " INNER JOIN MCM_TK_KIKIKOSEI k ON t.TK_KIKIKOSEI_ID = k.TK_KIKIKOSEI_ID" +
            " WHERE k.TK_KIKAN_ID = ?",
            BigDecimal.class, tkKikanId);
    }

    /** 期間に紐づく機器構成を取得（MCM1006U 変更ボタン用） */
    public List<java.util.Map<String, Object>> findKoseiRawByKikanId(BigDecimal tkKikanId) {
        return jdbcTemplate.queryForList(
            "SELECT TK_KIKIKOSEI_ID, KIKIKOSEI_ID, KIKIKOSEI_NK, SET_NM, TEHAISEIBAN, HYOJIJUN" +
            " FROM MCM.MCM_TK_KIKIKOSEI WHERE TK_KIKAN_ID = ?", tkKikanId);
    }

    /** 期間に紐づく機器明細を取得（MCM1006U 変更ボタン用） */
    public List<java.util.Map<String, Object>> findMeisaiRawByKikanId(BigDecimal tkKikanId) {
        return jdbcTemplate.queryForList(
            "SELECT m.TK_KIKIMEISAI_ID, m.TK_KIKIKOSEI_ID, m.KIKIKOSEI_ID, m.KIKIMEISAI_ID," +
            " m.SEIZOMAKER_ID, m.SEIZOMAKER_NK, m.KIKIHINMEI_NK, m.KIKIKATASHIKI," +
            " m.SURYO_NM, m.ATSUKAIKIKI_ID, m.HYOJIJUN" +
            " FROM MCM.MCM_TK_KIKIMEISAI m" +
            " INNER JOIN MCM.MCM_TK_KIKIKOSEI k ON m.TK_KIKIKOSEI_ID = k.TK_KIKIKOSEI_ID" +
            " WHERE k.TK_KIKAN_ID = ?", tkKikanId);
    }

    /** 期間に紐づく個体明細を取得（MCM1006U 変更ボタン用） */
    public List<java.util.Map<String, Object>> findKotaiRawByKikanId(BigDecimal tkKikanId) {
        return jdbcTemplate.queryForList(
            "SELECT o.TK_KOTAIMEISAI_ID, o.TK_KIKIKOSEI_ID, o.TK_KIKIMEISAI_ID," +
            " o.KIKIKOSEI_ID, o.KIKIMEISAI_ID, o.KOTAIKANRI_ID, o.KOTAI_NK, o.SERIAL_NO" +
            " FROM MCM.MCM_TK_KOTAIMEISAI o" +
            " INNER JOIN MCM.MCM_TK_KIKIKOSEI k ON o.TK_KIKIKOSEI_ID = k.TK_KIKIKOSEI_ID" +
            " WHERE k.TK_KIKAN_ID = ?", tkKikanId);
    }

    /** 期間に紐づく単価を取得（MCM1006U 変更ボタン用） */
    public List<java.util.Map<String, Object>> findTankaRawByKikanId(BigDecimal tkKikanId) {
        return jdbcTemplate.queryForList(
            "SELECT t.TK_TANKA_ID, t.TK_KIKIKOSEI_ID, t.TK_KIKIMEISAI_ID," +
            " t.KIKIKOSEI_ID, t.KIKIMEISAI_ID, t.TM_TANKA_ID, t.TM_KIKIMEISAI_ID" +
            " FROM MCM.MCM_TK_TANKA t" +
            " INNER JOIN MCM.MCM_TK_KIKIKOSEI k ON t.TK_KIKIKOSEI_ID = k.TK_KIKIKOSEI_ID" +
            " WHERE k.TK_KIKAN_ID = ?", tkKikanId);
    }

    // --- MAX IDs ---

    public BigDecimal getMaxKikikoseiId() {
        BigDecimal v = jdbcTemplate.queryForObject(
            "SELECT ISNULL(MAX(TK_KIKIKOSEI_ID),0) FROM MCM.MCM_TK_KIKIKOSEI", BigDecimal.class);
        return v == null ? BigDecimal.ZERO : v;
    }

    public BigDecimal getMaxKikimeisaiId() {
        BigDecimal v = jdbcTemplate.queryForObject(
            "SELECT ISNULL(MAX(TK_KIKIMEISAI_ID),0) FROM MCM.MCM_TK_KIKIMEISAI", BigDecimal.class);
        return v == null ? BigDecimal.ZERO : v;
    }

    public BigDecimal getMaxKotaimeisaiId() {
        BigDecimal v = jdbcTemplate.queryForObject(
            "SELECT ISNULL(MAX(TK_KOTAIMEISAI_ID),0) FROM MCM.MCM_TK_KOTAIMEISAI", BigDecimal.class);
        return v == null ? BigDecimal.ZERO : v;
    }

    public BigDecimal getMaxTankaId() {
        BigDecimal v = jdbcTemplate.queryForObject(
            "SELECT ISNULL(MAX(TK_TANKA_ID),0) FROM MCM.MCM_TK_TANKA", BigDecimal.class);
        return v == null ? BigDecimal.ZERO : v;
    }

    // --- INSERT ---

    public void insertKikikosei(BigDecimal tkKikikoseiId, BigDecimal tkKikanId,
            BigDecimal kikikoseiId, String kikikoseiNk, String setNm, String tehaiseiban,
            BigDecimal hyojijun, String loginUser) {
        jdbcTemplate.update(
            "INSERT INTO MCM.MCM_TK_KIKIKOSEI" +
            " (TK_KIKIKOSEI_ID, TK_KIKAN_ID, KIKIKOSEI_ID, KIKIKOSEI_NK, SET_NM," +
            "  TEHAISEIBAN, HYOJIJUN, CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY)" +
            " VALUES (?,?,?,?,?,?,?,GETDATE(),?,GETDATE(),?)",
            tkKikikoseiId, tkKikanId, kikikoseiId, kikikoseiNk, setNm,
            tehaiseiban, hyojijun, loginUser, loginUser);
    }

    public void insertKikimeisai(BigDecimal tkKikimeisaiId, BigDecimal tkKikikoseiId,
            BigDecimal kikikoseiId, BigDecimal kikimeisaiId,
            BigDecimal seizomakerId, String seizomakerNk, String kikihinmeiNk,
            String kikikatashiki, BigDecimal suryoNm, BigDecimal atsukaikikiId,
            BigDecimal hyojijun, String loginUser) {
        jdbcTemplate.update(
            "INSERT INTO MCM.MCM_TK_KIKIMEISAI" +
            " (TK_KIKIMEISAI_ID, TK_KIKIKOSEI_ID, KIKIKOSEI_ID, KIKIMEISAI_ID," +
            "  SEIZOMAKER_ID, SEIZOMAKER_NK, KIKIHINMEI_NK, KIKIKATASHIKI," +
            "  SURYO_NM, ATSUKAIKIKI_ID, HYOJIJUN," +
            "  CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY)" +
            " VALUES (?,?,?,?,?,?,?,?,?,?,?,GETDATE(),?,GETDATE(),?)",
            tkKikimeisaiId, tkKikikoseiId, kikikoseiId, kikimeisaiId,
            seizomakerId, seizomakerNk, kikihinmeiNk, kikikatashiki,
            suryoNm, atsukaikikiId, hyojijun, loginUser, loginUser);
    }

    public void insertKotaimeisai(BigDecimal tkKotaimeisaiId, BigDecimal tkKikikoseiId,
            BigDecimal tkKikimeisaiId, BigDecimal kikikoseiId, BigDecimal kikimeisaiId,
            BigDecimal kotaikanriId, String kotaiNk, String serialNo,
            String itijinonyugDt, String setchibasyo, String tekkyoDt,
            String enchokeiyakukigenDt, String loginUser) {
        jdbcTemplate.update(
            "INSERT INTO MCM.MCM_TK_KOTAIMEISAI" +
            " (TK_KOTAIMEISAI_ID, TK_KIKIKOSEI_ID, TK_KIKIMEISAI_ID," +
            "  KIKIKOSEI_ID, KIKIMEISAI_ID, KOTAIKANRI_ID, KOTAI_NK, SERIAL_NO," +
            "  ITIJINONYU_DT, SETCHIBASYO, TEKKYOBI_DT, ENCHOKEIYAKUKIGEN_DT," +
            "  CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY)" +
            " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,GETDATE(),?,GETDATE(),?)",
            tkKotaimeisaiId, tkKikikoseiId, tkKikimeisaiId,
            kikikoseiId, kikimeisaiId, kotaikanriId, kotaiNk, serialNo,
            itijinonyugDt, setchibasyo, tekkyoDt, enchokeiyakukigenDt,
            loginUser, loginUser);
    }

    public void insertTanka(BigDecimal tkTankaId, BigDecimal tkKikikoseiId, BigDecimal tkKikimeisaiId,
            BigDecimal kikikoseiId, BigDecimal kikimeisaiId,
            BigDecimal tmKikimeisaiId, BigDecimal tmTankaId, Integer packFlg,
            String keiyakuNo, String keiyakunaiyo, BigDecimal torihosyujikanId,
            BigDecimal tenkenumu, String hosyuhoho, String servicekeitai,
            BigDecimal hyojunKin, BigDecimal sikiriKin, BigDecimal suryoNm,
            String kaisiDt, String syuryoDt, String loginUser) {
        BigDecimal hyojunkeiKin = (hyojunKin != null && suryoNm != null) ? hyojunKin.multiply(suryoNm) : BigDecimal.ZERO;
        BigDecimal sikirikeiKin = (sikiriKin != null && suryoNm != null) ? sikiriKin.multiply(suryoNm) : BigDecimal.ZERO;
        jdbcTemplate.update(
            "INSERT INTO MCM.MCM_TK_TANKA" +
            " (TK_TANKA_ID, TK_KIKIKOSEI_ID, TK_KIKIMEISAI_ID," +
            "  KIKIKOSEI_ID, KIKIMEISAI_ID, TM_KIKIMEISAI_ID, TM_TANKA_ID," +
            "  PACK_FLG, KEIYAKU_NO, KEIYAKUNAIYO, TORIHOSYUJIKAN_ID," +
            "  TENKENUMU, HOSYUHOHO, SERVICEKEITAI," +
            "  HYOJUN_KIN, SIKIRI_KIN, HYOJUNKEI_KIN, SIKIRIKEI_KIN," +
            "  SURYO_NM, KAISI_DT, SYURYO_DT," +
            "  CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY)" +
            " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,GETDATE(),?,GETDATE(),?)",
            tkTankaId, tkKikikoseiId, tkKikimeisaiId,
            kikikoseiId, kikimeisaiId, tmKikimeisaiId, tmTankaId,
            packFlg, keiyakuNo, keiyakunaiyo, torihosyujikanId,
            tenkenumu, hosyuhoho, servicekeitai,
            hyojunKin, sikiriKin, hyojunkeiKin, sikirikeiKin,
            suryoNm, kaisiDt, syuryoDt, loginUser, loginUser);
    }

    // --- CASCADE DELETE ---

    /** MCM1006U結果適用前の機器情報一括削除（kikikosei配下のtenken/meisai/kotai/tanka） */
    public void deleteKikisForKikanId(BigDecimal tkKikanId) {
        String sub = "(SELECT TK_KIKIKOSEI_ID FROM MCM.MCM_TK_KIKIKOSEI WHERE TK_KIKAN_ID=?)";
        jdbcTemplate.update("DELETE FROM MCM.MCM_TK_KOTAIMEISAI WHERE TK_KIKIKOSEI_ID IN " + sub, tkKikanId);
        jdbcTemplate.update("DELETE FROM MCM.MCM_TK_TANKA WHERE TK_KIKIKOSEI_ID IN " + sub, tkKikanId);
        jdbcTemplate.update("DELETE FROM MCM.MCM_TK_TENKENMEISAI WHERE TK_TENKEN_ID IN" +
            " (SELECT TK_TENKEN_ID FROM MCM.MCM_TK_TENKEN WHERE TK_KIKIKOSEI_ID IN " + sub + ")", tkKikanId);
        jdbcTemplate.update("DELETE FROM MCM.MCM_TK_TENKEN WHERE TK_KIKIKOSEI_ID IN " + sub, tkKikanId);
        jdbcTemplate.update("DELETE FROM MCM.MCM_TK_KIKIMEISAI WHERE TK_KIKIKOSEI_ID IN " + sub, tkKikanId);
        jdbcTemplate.update("DELETE FROM MCM.MCM_TK_KIKIKOSEI WHERE TK_KIKAN_ID=?", tkKikanId);
    }

    // ===================================================================
    // RowMapper
    // ===================================================================

    private static class KeiyakuRowMapper implements RowMapper<McmTkKeiyakuEntity> {
        @Override
        public McmTkKeiyakuEntity mapRow(ResultSet rs, int rowNum) throws SQLException {
            McmTkKeiyakuEntity e = new McmTkKeiyakuEntity();
            e.setTkKeiyakuId(rs.getBigDecimal("TK_KEIYAKU_ID"));
            e.setIraijigyosyoNk(rs.getString("IRAIJIGYOSYO_NK"));
            e.setIraitantosya(rs.getString("IRAITANTOSYA"));
            e.setKeiyakuNo(rs.getString("KEIYAKU_NO"));
            e.setKeiyakuDt(toLocalDT(rs, "KEIYAKU_DT"));
            e.setKeiyakukaisiDt(toLocalDT(rs, "KEIYAKUKAISI_DT"));
            e.setKeiyakusyuryoDt(toLocalDT(rs, "KEIYAKUSYURYO_DT"));
            e.setJidokosinFlg(rs.getBigDecimal("JIDOKOSIN_FLG"));
            e.setJikaikosinDt(toLocalDT(rs, "JIKAIKOSIN_DT"));
            e.setKeiyakumanryoDt(toLocalDT(rs, "KEIYAKUMANRYO_DT"));
            e.setEntyokeiyakumanryoDt(toLocalDT(rs, "ENTYOKEIYAKUMANRYO_DT"));
            e.setKaiyakuDt(toLocalDT(rs, "KAIYAKU_DT"));
            e.setJotai(rs.getString("JOTAI"));
            e.setShoninjotai(rs.getString("SHONINJOTAI"));
            e.setPackFlg(rs.getBigDecimal("PACK_FLG"));
            e.setPackkeiyakunaiyo(rs.getString("PACKKEIYAKUNAIYO"));
            e.setPackkeiyakuNo(rs.getString("PACKKEIYAKU_NO"));
            e.setBiko(rs.getString("BIKO"));
            e.setKosinnaiyo(rs.getString("KOSINNAIYO"));
            e.setShokaiKeiyakuDt(toLocalDT(rs, "SHOKAI_KEIYAKU_DT"));
            e.setCreatedDt(toLocalDT(rs, "CREATED_DT"));
            e.setCreatedBy(rs.getString("CREATED_BY"));
            e.setLastupdateDt(toLocalDT(rs, "LASTUPDATE_DT"));
            e.setLastupdateBy(rs.getString("LASTUPDATE_BY"));
            return e;
        }
    }

    private static class KikanRowMapper implements RowMapper<McmTkKikanEntity> {
        @Override
        public McmTkKikanEntity mapRow(ResultSet rs, int rowNum) throws SQLException {
            McmTkKikanEntity e = new McmTkKikanEntity();
            e.setTkKikanId(rs.getBigDecimal("TK_KIKAN_ID"));
            e.setTkKeiyakuId(rs.getBigDecimal("TK_KEIYAKU_ID"));
            e.setKaisiDt(toLocalDT(rs, "KAISI_DT"));
            e.setSyuryoDt(toLocalDT(rs, "SYURYO_DT"));
            e.setHyojungokeiKin(rs.getBigDecimal("HYOJUNGOKEI_KIN"));
            e.setSikirisyokeiKin(rs.getBigDecimal("SIKIRISYOKEI_KIN"));
            e.setSyusseinebikiKin(rs.getBigDecimal("SYUSSEINEBIKI_KIN"));
            e.setSikirigokeiKin(rs.getBigDecimal("SIKIRIGOKEI_KIN"));
            e.setNonyusakiId(rs.getBigDecimal("NONYUSAKI_ID"));
            e.setNonyusakiCd(rs.getString("NONYUSAKI_CD"));
            e.setNonyusakiNk(rs.getString("NONYUSAKI_NK"));
            e.setNonyusakijusyo1Nk(rs.getString("NONYUSAKIJUSYO1_NK"));
            e.setNonyusakijusyo2Nk(rs.getString("NONYUSAKIJUSYO2_NK"));
            e.setPlantId(rs.getBigDecimal("PLANT_ID"));
            e.setSupportId(rs.getString("SUPPORT_ID"));
            e.setPlantNk(rs.getString("PLANT_NK"));
            e.setNonyubusyoNk(rs.getString("NONYUBUSYO_NK"));
            e.setNonyutantosyaNk(rs.getString("NONYUTANTOSYA_NK"));
            e.setNonyutelNo(rs.getString("NONYUTEL_NO"));
            e.setNonyufaxNo(rs.getString("NONYUFAX_NO"));
            e.setBiko(rs.getString("BIKO"));
            e.setTorihikisakiId(rs.getBigDecimal("TORIHIKISAKI_ID"));
            e.setTorihikisakiCd(rs.getString("TORIHIKISAKI_CD"));
            e.setTorihikisakiNk(rs.getString("TORIHIKISAKI_NK"));
            e.setToritelNo(rs.getString("TORITEL_NO"));
            e.setTorifaxNo(rs.getString("TORIFAX_NO"));
            e.setTorijigyosyoNk(rs.getString("TORIJIGYOSYO_NK"));
            e.setTorisyutantosyaNk(rs.getString("TORISYUTANTOSYA_NK"));
            e.setToriassistantNk(rs.getString("TORIASSISTANT_NK"));
            e.setKeiyakujikantai(rs.getString("KEIYAKUJIKANTAI"));
            e.setHosyuhoho(rs.getString("HOSYUHOHO"));
            e.setYukoFlg(rs.getBigDecimal("YUKO_FLG"));
            e.setKeiyakuNo(rs.getString("KEIYAKU_NO"));
            e.setCreatedDt(toLocalDT(rs, "CREATED_DT"));
            e.setCreatedBy(rs.getString("CREATED_BY"));
            e.setLastupdateDt(toLocalDT(rs, "LASTUPDATE_DT"));
            e.setLastupdateBy(rs.getString("LASTUPDATE_BY"));
            return e;
        }
    }

    private static class TenkenRowMapper implements RowMapper<McmTkTenkenEntity> {
        @Override
        public McmTkTenkenEntity mapRow(ResultSet rs, int rowNum) throws SQLException {
            McmTkTenkenEntity e = new McmTkTenkenEntity();
            e.setTkTenkenId(rs.getBigDecimal("TK_TENKEN_ID"));
            e.setTkKikokoseiId(rs.getBigDecimal("TK_KIKIKOSEI_ID"));
            e.setTenkenkaisu(rs.getBigDecimal("TENKENKAISU"));
            e.setTenkenkanoyobi(rs.getString("TENKENKANOYOBI"));
            e.setYakantaioumu(rs.getString("YAKANTAIOUMU"));
            e.setHoshugaishaId(rs.getBigDecimal("HOSHUGAISHA_ID"));
            e.setYakanhoshugaishaId(rs.getBigDecimal("YAKANHOSHUGAISHA_ID"));
            e.setOyakikibunruiCd(rs.getString("OYAKIKIBUNRUI_CD"));
            e.setBiko(rs.getString("BIKO"));
            e.setCreatedDt(toLocalDT(rs, "CREATED_DT"));
            e.setCreatedBy(rs.getString("CREATED_BY"));
            e.setLastupdateDt(toLocalDT(rs, "LASTUPDATE_DT"));
            e.setLastupdateBy(rs.getString("LASTUPDATE_BY"));
            e.setTkKikanId(rs.getBigDecimal("TK_KIKAN_ID"));
            e.setKikikoseiNk(rs.getString("KIKIKOSEI_NK"));
            return e;
        }
    }

    private static class SiharaiRowMapper implements RowMapper<McmTkSiharaiEntity> {
        @Override
        public McmTkSiharaiEntity mapRow(ResultSet rs, int rowNum) throws SQLException {
            McmTkSiharaiEntity e = new McmTkSiharaiEntity();
            e.setTkSiharaiId(rs.getBigDecimal("TK_SIHARAI_ID"));
            e.setTkKikanId(rs.getBigDecimal("TK_KIKAN_ID"));
            e.setKaisu(rs.getBigDecimal("KAISU"));
            e.setBiko(rs.getString("BIKO"));
            e.setCreatedDt(toLocalDT(rs, "CREATED_DT"));
            e.setCreatedBy(rs.getString("CREATED_BY"));
            e.setLastupdateDt(toLocalDT(rs, "LASTUPDATE_DT"));
            e.setLastupdateBy(rs.getString("LASTUPDATE_BY"));
            return e;
        }
    }

    private static class TenpuRowMapper implements RowMapper<McmTkTenpuEntity> {
        @Override
        public McmTkTenpuEntity mapRow(ResultSet rs, int rowNum) throws SQLException {
            McmTkTenpuEntity e = new McmTkTenpuEntity();
            e.setTkTenpuId(rs.getBigDecimal("TK_TENPU_ID"));
            e.setTkKikanId(rs.getBigDecimal("TK_KIKAN_ID"));
            e.setTenpufileNk(rs.getString("TENPUFILE_NK"));
            e.setDirectory(rs.getString("DIRECTORY"));
            e.setShoninjotai(rs.getString("SHONINJOTAI"));
            e.setBiko(rs.getString("BIKO"));
            e.setCreatedDt(toLocalDT(rs, "CREATED_DT"));
            e.setCreatedBy(rs.getString("CREATED_BY"));
            e.setLastupdateDt(toLocalDT(rs, "LASTUPDATE_DT"));
            e.setLastupdateBy(rs.getString("LASTUPDATE_BY"));
            return e;
        }
    }

    private static LocalDateTime toLocalDT(ResultSet rs, String col) throws SQLException {
        return rs.getTimestamp(col) != null ? rs.getTimestamp(col).toLocalDateTime() : null;
    }
}
