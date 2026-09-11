package com.daifuku.mcm.repository;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Repository;

import com.daifuku.mcm.form.Mcm2006uForm;
import com.daifuku.mcm.form.Mcm2006uForm.BrandRowForm;
import com.daifuku.mcm.form.Mcm2006uForm.KikanTabForm;
import com.daifuku.mcm.form.Mcm2006uForm.KotaiRowForm;
import com.daifuku.mcm.form.Mcm2006uForm.KoseiRowForm;
import com.daifuku.mcm.form.Mcm2006uForm.MeisaiRowForm;
import com.daifuku.mcm.form.Mcm2006uForm.SeibanRowForm;
import com.daifuku.mcm.form.Mcm2006uForm.TenkenRowForm;

/**
 * 【変換元】Mcm2006uDataSet.Designer.vb / Mcm2006uTabControl.vb
 * MCM2006U ユーザ契約内容リポジトリ
 *
 * SQL Server構文使用。全テーブル名にMCM.スキーマ付与。
 * SP_UKストアドプロシージャ・ロック解除はWeb版省略。
 */
@Repository
public class Mcm2006uRepository {

    @Autowired
    private JdbcTemplate jdbc;

    private static final DateTimeFormatter DT_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");

    private String fmtDt(java.sql.Date d) {
        if (d == null) return "";
        return d.toLocalDate().format(DT_FMT);
    }

    private String fmtLdt(LocalDateTime ldt) {
        if (ldt == null) return "";
        return ldt.toLocalDate().format(DT_FMT);
    }

    // ===================================================================
    // MAX ID 採番 (新規INSERT時)
    // ===================================================================

    public BigDecimal nextKeiyakuId() {
        return jdbc.queryForObject(
            "SELECT ISNULL(MAX(UK_KEIYAKU_ID),0)+1 FROM MCM.MCM_UK_KEIYAKU WITH (UPDLOCK,HOLDLOCK)",
            BigDecimal.class);
    }

    public BigDecimal nextKikanId() {
        return jdbc.queryForObject(
            "SELECT ISNULL(MAX(UK_KIKAN_ID),0)+1 FROM MCM.MCM_UK_KIKAN WITH (UPDLOCK,HOLDLOCK)",
            BigDecimal.class);
    }

    public BigDecimal nextSeibanId() {
        return jdbc.queryForObject(
            "SELECT ISNULL(MAX(UK_SEIBAN_ID),0)+1 FROM MCM.MCM_UK_SEIBAN WITH (UPDLOCK,HOLDLOCK)",
            BigDecimal.class);
    }

    public BigDecimal nextBrandId() {
        return jdbc.queryForObject(
            "SELECT ISNULL(MAX(UK_BRAND_ID),0)+1 FROM MCM.MCM_UK_BRAND WITH (UPDLOCK,HOLDLOCK)",
            BigDecimal.class);
    }

    public BigDecimal nextKikoseiId() {
        return jdbc.queryForObject(
            "SELECT ISNULL(MAX(UK_KIKIKOSEI_ID),0)+1 FROM MCM.MCM_UK_KIKIKOSEI WITH (UPDLOCK,HOLDLOCK)",
            BigDecimal.class);
    }

    public BigDecimal nextKikimeisaiId() {
        return jdbc.queryForObject(
            "SELECT ISNULL(MAX(UK_KIKIMEISAI_ID),0)+1 FROM MCM.MCM_UK_KIKIMEISAI WITH (UPDLOCK,HOLDLOCK)",
            BigDecimal.class);
    }

    public BigDecimal nextKotaimeisaiId() {
        return jdbc.queryForObject(
            "SELECT ISNULL(MAX(UK_KOTAIMEISAI_ID),0)+1 FROM MCM.MCM_UK_KOTAIMEISAI WITH (UPDLOCK,HOLDLOCK)",
            BigDecimal.class);
    }

    public BigDecimal nextTenkenId() {
        return jdbc.queryForObject(
            "SELECT ISNULL(MAX(UK_TENKEN_ID),0)+1 FROM MCM.MCM_UK_TENKEN WITH (UPDLOCK,HOLDLOCK)",
            BigDecimal.class);
    }

    // ===================================================================
    // 契約ヘッダー取得 (MCM_UK_KEIYAKU)
    // 【変換元】MCM_UK_KEIYAKUTableAdapter.Fill
    // ===================================================================

    public Mcm2006uForm findKeiyaku(BigDecimal ukKeiyakuId) {
        List<Mcm2006uForm> list = jdbc.query(
            "SELECT A.LASTUPDATE_DT, A.UK_KEIYAKU_ID, A.KEIYAKU_DT, A.SHOKAI_KEIYAKU_DT, A.JIKAIKOSIN_DT, A.KEIYAKUMANRYO_DT, A.ENTYOKEIYAKUMANRYO_DT, A.KAIYAKU_DT, A.JOTAI, A.AUTO_FLG, K.NONYUSAKI_ID, K.NONYUSAKI_CD, K.NONYUSAKI_NK, K.NONYUSAKIJUSYO1_NK, K.NONYUSAKIJUSYO2_NK, K.PLANT_ID, K.SUPPORT_ID, K.PLANT_NK, K.NONYUBUSYO_NK, K.NONYUTANTOSYA_NK, K.NONYUTEL_NO, K.NONYUFAX_NO " +
            "FROM MCM.MCM_UK_KEIYAKU A LEFT JOIN MCM.MCM_UK_KIKAN K ON K.UK_KIKAN_ID = " +
            "(SELECT TOP (1) UK_KIKAN_ID FROM MCM.MCM_UK_KIKAN WHERE UK_KEIYAKU_ID=A.UK_KEIYAKU_ID ORDER BY KAISI_DT DESC,UK_KIKAN_ID DESC) " +
            "WHERE A.UK_KEIYAKU_ID = ?",
            new Object[]{ukKeiyakuId},
            (rs, rn) -> {
                Mcm2006uForm f = new Mcm2006uForm();
                f.setUkKeiyakuId(rs.getBigDecimal("UK_KEIYAKU_ID"));
                f.setVersion(rs.getTimestamp("LASTUPDATE_DT")==null?null:rs.getTimestamp("LASTUPDATE_DT").toString());
                f.setNonyusakiId(rs.getBigDecimal("NONYUSAKI_ID"));
                f.setNonyusakiCd(rs.getString("NONYUSAKI_CD"));
                f.setNonyusakiNk(rs.getString("NONYUSAKI_NK"));
                f.setNonyusakijusyo1Nk(rs.getString("NONYUSAKIJUSYO1_NK"));
                f.setNonyusakijusyo2Nk(rs.getString("NONYUSAKIJUSYO2_NK"));
                f.setPlantId(rs.getBigDecimal("PLANT_ID"));
                f.setSupportId(rs.getString("SUPPORT_ID"));
                f.setPlantNk(rs.getString("PLANT_NK"));
                f.setNonyubusyoNk(rs.getString("NONYUBUSYO_NK"));
                f.setNonyutantosyaNk(rs.getString("NONYUTANTOSYA_NK"));
                f.setNonyutelNo(rs.getString("NONYUTEL_NO"));
                f.setNonyufaxNo(rs.getString("NONYUFAX_NO"));
                f.setKeiyakuDt(fmtDt(rs.getDate("KEIYAKU_DT")));
                f.setShokaiKeiyakuDt(fmtDt(rs.getDate("SHOKAI_KEIYAKU_DT")));
                f.setJikaikosinDt(fmtDt(rs.getDate("JIKAIKOSIN_DT")));
                f.setKeiyakumanryoDt(fmtDt(rs.getDate("KEIYAKUMANRYO_DT")));
                f.setEntyokeiyakumanryoDt(fmtDt(rs.getDate("ENTYOKEIYAKUMANRYO_DT")));
                f.setKaiyakuDt(fmtDt(rs.getDate("KAIYAKU_DT")));
                f.setJotai(rs.getString("JOTAI"));
                f.setAutoFlg(rs.getBigDecimal("AUTO_FLG")==null?null:rs.getBigDecimal("AUTO_FLG").stripTrailingZeros().toPlainString());
                return f;
            });
        return list.isEmpty() ? new Mcm2006uForm() : list.get(0);
    }

    // ===================================================================
    // 製番一覧取得 (MCM_UK_SEIBAN)
    // 【変換元】MCM_UK_SEIBANTableAdapter.Fill
    // ===================================================================

    public List<SeibanRowForm> findSeibanList(BigDecimal ukKeiyakuId) {
        return jdbc.query(
            "SELECT UK_SEIBAN_ID, UK_KEIYAKU_ID, KAISI_DT, SYURYO_DT, KAKUNIN_KBN " +
            "FROM MCM.MCM_UK_SEIBAN " +
            "WHERE UK_KEIYAKU_ID = ? " +
            "ORDER BY KAISI_DT",
            new Object[]{ukKeiyakuId},
            (rs, rn) -> {
                SeibanRowForm r = new SeibanRowForm();
                r.setUkSeibanId(rs.getBigDecimal("UK_SEIBAN_ID"));
                r.setUkKeiyakuId(rs.getBigDecimal("UK_KEIYAKU_ID"));
                r.setKaisiDt(fmtDt(rs.getDate("KAISI_DT")));
                r.setSyuryoDt(fmtDt(rs.getDate("SYURYO_DT")));
                r.setKakuninKbn(rs.getString("KAKUNIN_KBN"));
                return r;
            });
    }

    // ===================================================================
    // 期間タブ一覧取得 (MCM_UK_KIKAN)
    // 【変換元】MCM_UK_KIKANTableAdapter.Fill
    // ===================================================================

    public List<KikanTabForm> findKikanList(BigDecimal ukKeiyakuId) {
        return jdbc.query(
            "SELECT UK_KIKAN_ID, UK_KEIYAKU_ID, KAISI_DT, SYURYO_DT, " +
            "       NONYUSAKI_ID, NONYUSAKI_CD, NONYUSAKI_NK, " +
            "       NONYUSAKIJUSYO1_NK, NONYUSAKIJUSYO2_NK, " +
            "       PLANT_ID, SUPPORT_ID, PLANT_NK, " +
            "       NONYUBUSYO_NK, NONYUTANTOSYA_NK, NONYUTEL_NO, NONYUFAX_NO, " +
            "       KEIYAKUJIKANTAI, HOSYUHOHO, HOSYU_GKIN, " +
            "       IRAITENPO_ID, IRAIMEISHO1_NK, IRAIMEISHO2_NK, IRAIMEISHO3_NK, IRAIMEISHO4_NK, " +
            "       IRAITENPORYAKU_NK, IRAITANTO_NK, BIKO, YUKO_FLG " +
            "FROM MCM.MCM_UK_KIKAN " +
            "WHERE UK_KEIYAKU_ID = ? " +
            "ORDER BY KAISI_DT",
            new Object[]{ukKeiyakuId},
            (rs, rn) -> mapKikan(rs));
    }

    private KikanTabForm mapKikan(java.sql.ResultSet rs) throws java.sql.SQLException {
        KikanTabForm k = new KikanTabForm();
        k.setUkKikanId(rs.getBigDecimal("UK_KIKAN_ID"));
        k.setUkKeiyakuId(rs.getBigDecimal("UK_KEIYAKU_ID"));
        k.setKaisiDt(fmtDt(rs.getDate("KAISI_DT")));
        k.setSyuryoDt(fmtDt(rs.getDate("SYURYO_DT")));
        k.setNonyusakiId(rs.getBigDecimal("NONYUSAKI_ID"));
        k.setNonyusakiCd(rs.getString("NONYUSAKI_CD"));
        k.setNonyusakiNk(rs.getString("NONYUSAKI_NK"));
        k.setNonyusakijusyo1Nk(rs.getString("NONYUSAKIJUSYO1_NK"));
        k.setNonyusakijusyo2Nk(rs.getString("NONYUSAKIJUSYO2_NK"));
        k.setPlantId(rs.getBigDecimal("PLANT_ID"));
        k.setSupportId(rs.getString("SUPPORT_ID"));
        k.setPlantNk(rs.getString("PLANT_NK"));
        k.setNonyubusyoNk(rs.getString("NONYUBUSYO_NK"));
        k.setNonyutantosyaNk(rs.getString("NONYUTANTOSYA_NK"));
        k.setNonyutelNo(rs.getString("NONYUTEL_NO"));
        k.setNonyufaxNo(rs.getString("NONYUFAX_NO"));
        k.setKeiyakujikantai(rs.getString("KEIYAKUJIKANTAI"));
        k.setHosyuhoho(rs.getString("HOSYUHOHO"));
        k.setHosyuGkin(rs.getBigDecimal("HOSYU_GKIN"));
        k.setIraitenpoId(rs.getBigDecimal("IRAITENPO_ID"));
        k.setIraimeisho1Nk(rs.getString("IRAIMEISHO1_NK"));
        k.setIraimeisho2Nk(rs.getString("IRAIMEISHO2_NK"));
        k.setIraimeisho3Nk(rs.getString("IRAIMEISHO3_NK"));
        k.setIraimeisho4Nk(rs.getString("IRAIMEISHO4_NK"));
        k.setIraitenporyakuNk(rs.getString("IRAITENPORYAKU_NK"));
        k.setIraitantoNk(rs.getString("IRAITANTO_NK"));
        k.setBiko(rs.getString("BIKO"));
        k.setYukoFlg(rs.getString("YUKO_FLG"));
        return k;
    }

    // ===================================================================
    // タブ内容取得 — ブランド (MCM_UK_BRAND + MCM_MA_BRAND_KOSEI)
    // 【変換元】MCM_UK_BRANDTableAdapter.Fill
    // ===================================================================

    public List<BrandRowForm> findBrandRows(BigDecimal ukKikanId) {
        return jdbc.query(
            "SELECT B.UK_BRAND_ID, B.UK_KIKAN_ID, B.BRANDKOSEI_ID, " +
            "       B.KEIYAKUJIKANTAI, B.HOSYUHOHO, " +
            "       B.BRAND_NK, B.BRANDSYOSAI_NK, " +
            "       MB.UM_MITSUMORI_ID, UMB.UM_KIHON_BRAND_ID " +
            "FROM MCM.MCM_UK_BRAND B " +
            "INNER JOIN MCM.MCM_MA_BRAND_KOSEI MBC ON MBC.BRANDKOSEI_ID = B.BRANDKOSEI_ID " +
            "LEFT JOIN MCM.MCM_UK_MITSUMORI_BRAND MB ON MB.UK_BRAND_ID = B.UK_BRAND_ID " +
            "LEFT JOIN MCM.MCM_UM_MITSUMORI UM ON UM.UM_MITSUMORI_ID=MB.UM_MITSUMORI_ID " +
            "LEFT JOIN MCM.MCM_UM_KIHON_BRAND UMB ON UMB.UM_KIHON_MITSUMORI_ID=UM.UM_KIHON_MITSUMORI_ID AND UMB.BRANDKOSEI_ID=B.BRANDKOSEI_ID " +
            "WHERE B.UK_KIKAN_ID = ? " +
            "ORDER BY MBC.HYOJIJUN",
            new Object[]{ukKikanId},
            (rs, rn) -> {
                BrandRowForm r = new BrandRowForm();
                r.setUkBrandId(rs.getBigDecimal("UK_BRAND_ID"));
                r.setUkKikanId(rs.getBigDecimal("UK_KIKAN_ID"));
                r.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));
                r.setKeiyakujikantai(rs.getString("KEIYAKUJIKANTAI"));
                r.setHosyuhoho(rs.getString("HOSYUHOHO"));
                r.setBrandNk(rs.getString("BRAND_NK"));
                r.setBrandsyosaiNk(rs.getString("BRANDSYOSAI_NK"));
                r.setUmMitsumoriId(rs.getBigDecimal("UM_MITSUMORI_ID"));
                r.setUmKihonBrandId(rs.getBigDecimal("UM_KIHON_BRAND_ID"));
                return r;
            });
    }

    // ===================================================================
    // タブ内容取得 — 機器構成 (MCM_UK_KIKIKOSEI)
    // 【変換元】MCM_UK_KIKIKOSEITableAdapter.Fill
    // ===================================================================

    public List<KoseiRowForm> findKoseiRows(BigDecimal ukKikanId) {
        return jdbc.query(
            "SELECT K.UK_KIKIKOSEI_ID, K.UK_BRAND_ID, K.KIKIKOSEI_ID, B.BRANDKOSEI_ID, " +
            "       K.KIKIKOSEI_NK, K.SET_NM, K.TANI, K.TEHAISEIBAN, " +
            "       K.CONTROLLER_FLG, K.HOSYUHOHO, K.HYOJIJUN " +
            "FROM MCM.MCM_UK_KIKIKOSEI K " +
            "INNER JOIN MCM.MCM_UK_BRAND B ON B.UK_BRAND_ID = K.UK_BRAND_ID " +
            "WHERE B.UK_KIKAN_ID = ? " +
            "ORDER BY K.HYOJIJUN",
            new Object[]{ukKikanId},
            (rs, rn) -> {
                KoseiRowForm r = new KoseiRowForm();
                r.setUkKikoseiId(rs.getBigDecimal("UK_KIKIKOSEI_ID"));
                r.setUkBrandId(rs.getBigDecimal("UK_BRAND_ID"));
                r.setKikikoseiId(rs.getBigDecimal("KIKIKOSEI_ID"));
                r.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));
                r.setKikikoseiNk(rs.getString("KIKIKOSEI_NK"));
                r.setSetNm(rs.getString("SET_NM"));
                r.setTani(rs.getString("TANI"));
                r.setTehaiseiban(rs.getString("TEHAISEIBAN"));
                r.setControllerFlg(rs.getString("CONTROLLER_FLG"));
                r.setHosyuhoho(rs.getString("HOSYUHOHO"));
                r.setHyojijun(rs.getInt("HYOJIJUN"));
                return r;
            });
    }

    // ===================================================================
    // タブ内容取得 — 機器明細 (MCM_UK_KIKIMEISAI)
    // 【変換元】MCM_UK_KIKIMEISAITableAdapter.Fill
    // ===================================================================

    public List<MeisaiRowForm> findMeisaiRows(BigDecimal ukKikanId) {
        return jdbc.query(
            "SELECT M.UK_KIKIMEISAI_ID, M.UK_KIKIKOSEI_ID, M.KIKIMEISAI_ID, K.KIKIKOSEI_ID, " +
            "       B.BRANDKOSEI_ID, M.SEIZOMAKER_ID, M.SEIZOMAKER_NK, " +
            "       M.KIKIHINMEI_NK, M.KIKIKATASHIKI, M.SURYO_NM, " +
            "       M.KEIYAKUNAIYO, M.KEIYAKU_NO, M.SERVICEKEITAI, " +
            "       M.TORIHOSYUJIKAN_ID, M.DAIFUKUHOSYUJIKAN_ID, " +
            "       M.TENKENKAISU, M.TENKENYOBI, M.HOSYUHOHO, " +
            "       K.HYOJIJUN AS MAE_HYOJIJUN, M.HYOJIJUN " +
            "FROM MCM.MCM_UK_KIKIMEISAI M " +
            "INNER JOIN MCM.MCM_UK_KIKIKOSEI K ON K.UK_KIKIKOSEI_ID = M.UK_KIKIKOSEI_ID " +
            "INNER JOIN MCM.MCM_UK_BRAND B ON B.UK_BRAND_ID = K.UK_BRAND_ID " +
            "WHERE B.UK_KIKAN_ID = ? " +
            "ORDER BY K.HYOJIJUN, M.HYOJIJUN",
            new Object[]{ukKikanId},
            (rs, rn) -> {
                MeisaiRowForm r = new MeisaiRowForm();
                r.setUkKikimeisaiId(rs.getBigDecimal("UK_KIKIMEISAI_ID"));
                r.setUkKikoseiId(rs.getBigDecimal("UK_KIKIKOSEI_ID"));
                r.setKikimeisaiId(rs.getBigDecimal("KIKIMEISAI_ID"));
                r.setKikikoseiId(rs.getBigDecimal("KIKIKOSEI_ID"));
                r.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));
                r.setSeizomakerId(rs.getBigDecimal("SEIZOMAKER_ID"));
                r.setSeizomankerNk(rs.getString("SEIZOMAKER_NK"));
                r.setKikihinmeiNk(rs.getString("KIKIHINMEI_NK"));
                r.setKikikatashiki(rs.getString("KIKIKATASHIKI"));
                r.setSuryoNm(rs.getString("SURYO_NM"));
                r.setKeiyakunaiyo(rs.getString("KEIYAKUNAIYO"));
                r.setKeiyakuNo(rs.getString("KEIYAKU_NO"));
                r.setServicekeitai(rs.getString("SERVICEKEITAI"));
                r.setTorihosyujikanId(rs.getBigDecimal("TORIHOSYUJIKAN_ID"));
                r.setDaifukuhosyujikanId(rs.getBigDecimal("DAIFUKUHOSYUJIKAN_ID"));
                r.setTenkenkaisu(rs.getString("TENKENKAISU"));
                r.setTenkenyobi(rs.getString("TENKENYOBI"));
                r.setHosyuhoho(rs.getString("HOSYUHOHO"));
                r.setMaeHyojijun(rs.getInt("MAE_HYOJIJUN"));
                r.setHyojijun(rs.getInt("HYOJIJUN"));
                return r;
            });
    }

    // ===================================================================
    // タブ内容取得 — 個体明細 (MCM_UK_KOTAIMEISAI)
    // 【変換元】MCM_UK_KOTAIMEISAITableAdapter.Fill
    // ===================================================================

    public List<KotaiRowForm> findKotaiRows(BigDecimal ukKikanId) {
        return jdbc.query(
            "SELECT KO.UK_KOTAIMEISAI_ID, KO.UK_KIKIMEISAI_ID, M.UK_KIKIKOSEI_ID, " +
            "       KO.KOTAIKANRI_ID, M.KIKIMEISAI_ID, K.KIKIKOSEI_ID, B.BRANDKOSEI_ID, " +
            "       KO.KOTAI_NK, KO.SERIAL_NO, KO.ITIJINONYU_DT, KO.SETCHIBASYO, " +
            "       KO.KEIYAKUKIGEN_DT, KO.ENCHOKEIYAKUKIGEN_DT " +
            "FROM MCM.MCM_UK_KOTAIMEISAI KO " +
            "INNER JOIN MCM.MCM_UK_KIKIMEISAI M ON M.UK_KIKIMEISAI_ID = KO.UK_KIKIMEISAI_ID " +
            "INNER JOIN MCM.MCM_UK_KIKIKOSEI K ON K.UK_KIKIKOSEI_ID = M.UK_KIKIKOSEI_ID " +
            "INNER JOIN MCM.MCM_UK_BRAND B ON B.UK_BRAND_ID = K.UK_BRAND_ID " +
            "WHERE B.UK_KIKAN_ID = ? " +
            "ORDER BY B.BRANDKOSEI_ID, K.KIKIKOSEI_ID",
            new Object[]{ukKikanId},
            (rs, rn) -> {
                KotaiRowForm r = new KotaiRowForm();
                r.setUkKotaimeisaiId(rs.getBigDecimal("UK_KOTAIMEISAI_ID"));
                r.setUkKikimeisaiId(rs.getBigDecimal("UK_KIKIMEISAI_ID"));
                r.setUkKikoseiId(rs.getBigDecimal("UK_KIKIKOSEI_ID"));
                r.setKotaikanriId(rs.getBigDecimal("KOTAIKANRI_ID"));
                r.setKikimeisaiId(rs.getBigDecimal("KIKIMEISAI_ID"));
                r.setKikikoseiId(rs.getBigDecimal("KIKIKOSEI_ID"));
                r.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));
                r.setKotaiNk(rs.getString("KOTAI_NK"));
                r.setSerialNo(rs.getString("SERIAL_NO"));
                r.setItizinonnyuDt(fmtDt(rs.getDate("ITIJINONYU_DT")));
                r.setSetchibasyo(rs.getString("SETCHIBASYO"));
                r.setKeiyakukigenDt(fmtDt(rs.getDate("KEIYAKUKIGEN_DT")));
                r.setEnchokeiyakukigenDt(fmtDt(rs.getDate("ENCHOKEIYAKUKIGEN_DT")));
                return r;
            });
    }

    // ===================================================================
    // タブ内容取得 — 点検 (MCM_UK_TENKEN)
    // ===================================================================

    public List<TenkenRowForm> findTenkenRows(BigDecimal id) {
        return jdbc.query("SELECT UK_TENKEN_ID,UK_KIKAN_ID,NAIYO,M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12,BIKO FROM MCM.MCM_UK_TENKEN WHERE UK_KIKAN_ID=? ORDER BY UK_TENKEN_ID",new Object[]{id}, (rs,n)->{
            var r=new TenkenRowForm(); r.setUkTenkenId(rs.getBigDecimal("UK_TENKEN_ID")); r.setUkKikanId(rs.getBigDecimal("UK_KIKAN_ID"));
            r.setNaiyo(rs.getString("NAIYO")); r.setBiko(rs.getString("BIKO"));
            for(int i=1;i<=12;i++) r.getMonths().add(rs.getString(String.format("M%02d",i))); return r;
        });
    }

    // ===================================================================
    // 存在チェック
    // ===================================================================

    public boolean existsKeiyaku(BigDecimal ukKeiyakuId) {
        if (ukKeiyakuId == null) return false;
        Integer cnt = jdbc.queryForObject(
            "SELECT COUNT(1) FROM MCM.MCM_UK_KEIYAKU WHERE UK_KEIYAKU_ID=?",
            Integer.class, ukKeiyakuId);
        return cnt != null && cnt > 0;
    }

    public boolean existsKikan(BigDecimal ukKikanId) {
        if (ukKikanId == null) return false;
        Integer cnt = jdbc.queryForObject(
            "SELECT COUNT(1) FROM MCM.MCM_UK_KIKAN WHERE UK_KIKAN_ID=?",
            Integer.class, ukKikanId);
        return cnt != null && cnt > 0;
    }

    // ===================================================================
    // 保存: MCM_UK_KEIYAKU (INSERT / UPDATE)
    // 【変換元】UpdateButtonTabNaiyo → UpdateAll(MCM_UK_KEIYAKU)
    // ===================================================================

    public KikanTabForm findSourceDefaults(BigDecimal estimate) {
        var rows=jdbc.query("SELECT A.NONYUBUSYO_NK,A.NONYUTANTOSYA_NK,A.NONYUTEL_NO,A.NONYUFAX_NO,A.IRAITENPO_ID,A.IRAIMEISHO1_NK,A.IRAIMEISHO2_NK,A.IRAIMEISHO3_NK,A.IRAIMEISHO4_NK,A.IRAITENPORYAKU_NK,A.IRAITANTO_NK FROM MCM.MCM_UM_KIHON_MITSUMORI A JOIN MCM.MCM_UM_MITSUMORI M ON M.UM_KIHON_MITSUMORI_ID=A.UM_KIHON_MITSUMORI_ID WHERE M.UM_MITSUMORI_ID=?",new org.springframework.jdbc.core.BeanPropertyRowMapper<>(KikanTabForm.class),estimate);
        return rows.isEmpty()?new KikanTabForm():rows.get(0);
    }
    public boolean validSource(BrandRowForm brand,BigDecimal plant) {
        return jdbc.queryForObject("SELECT COUNT(*) FROM MCM.MCM_UM_BRAND L JOIN MCM.MCM_UM_KIHON_BRAND B ON B.UM_KIHON_BRAND_ID=L.UM_KIHON_BRAND_ID JOIN MCM.MCM_UM_KIHON_MITSUMORI A ON A.UM_KIHON_MITSUMORI_ID=B.UM_KIHON_MITSUMORI_ID WHERE L.UM_MITSUMORI_ID=? AND B.UM_KIHON_BRAND_ID=? AND B.BRANDKOSEI_ID=? AND A.PLANT_ID=? AND A.SYOUNIN_JOTAI='3' AND A.JOTAI<>'3'",Integer.class,brand.getUmMitsumoriId(),brand.getUmKihonBrandId(),brand.getBrandkoseiId(),plant)==1;
    }
    public void insertKeiyaku(Mcm2006uForm f, String user) {
        jdbc.update("INSERT INTO MCM.MCM_UK_KEIYAKU (UK_KEIYAKU_ID,KEIYAKU_DT,SHOKAI_KEIYAKU_DT,KAISI_DT,SYURYO_DT,JOTAI,AUTO_FLG,CREATED_DT,CREATED_BY,LASTUPDATE_DT,LASTUPDATE_BY) VALUES (?,?,?,?,?,?,?,GETDATE(),?,GETDATE(),?)",
            f.getUkKeiyakuId(),dtOrNull(f.getKeiyakuDt()),dtOrNull(f.getShokaiKeiyakuDt()),dtOrNull(f.getKikanTabs().get(0).getKaisiDt()),dtOrNull(f.getKikanTabs().get(f.getKikanTabs().size()-1).getSyuryoDt()),f.getJotai(),f.getAutoFlg(),user,user);
    }

    public void lock(BigDecimal id) {
        jdbc.queryForObject("SELECT UK_KEIYAKU_ID FROM MCM.MCM_UK_KEIYAKU WITH (UPDLOCK,HOLDLOCK) WHERE UK_KEIYAKU_ID=?",BigDecimal.class,id);
    }
    public void updateKeiyaku(Mcm2006uForm f, String user) {
        int rows = jdbc.update("UPDATE MCM.MCM_UK_KEIYAKU SET LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? WHERE UK_KEIYAKU_ID=? AND JOTAI=?",
            user, f.getUkKeiyakuId(), f.getJotai());
        if (rows != 1) throw new IllegalStateException("契約情報が変更されています。再検索してください。");
    }

    // ===================================================================
    // 保存: MCM_UK_SEIBAN (INSERT / UPDATE)
    // ===================================================================

    public void insertSeiban(SeibanRowForm r, String loginUser) {
        jdbc.update(
            "INSERT INTO MCM.MCM_UK_SEIBAN " +
            "(UK_SEIBAN_ID, UK_KEIYAKU_ID, KAISI_DT, SYURYO_DT, KAKUNIN_KBN, " +
            " CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY) " +
            "VALUES (?,?,?,?,?,GETDATE(),?,GETDATE(),?)",
            r.getUkSeibanId(), r.getUkKeiyakuId(),
            dtOrNull(r.getKaisiDt()), dtOrNull(r.getSyuryoDt()),
            r.getKakuninKbn(), loginUser, loginUser);
    }

    public void updateSeiban(SeibanRowForm r, String loginUser) {
        jdbc.update(
            "UPDATE MCM.MCM_UK_SEIBAN SET " +
            "KAISI_DT=?, SYURYO_DT=?, KAKUNIN_KBN=?, " +
            "LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? " +
            "WHERE UK_SEIBAN_ID=?",
            dtOrNull(r.getKaisiDt()), dtOrNull(r.getSyuryoDt()),
            r.getKakuninKbn(), loginUser, r.getUkSeibanId());
    }

    // ===================================================================
    // 保存: MCM_UK_KIKAN (INSERT / UPDATE)
    // 【変換元】UpdateAll(MCM_UK_KIKAN)
    // ===================================================================

    public void insertKikan(KikanTabForm k, String loginUser) {
        jdbc.update(
            "INSERT INTO MCM.MCM_UK_KIKAN " +
            "(UK_KIKAN_ID, UK_KEIYAKU_ID, KAISI_DT, SYURYO_DT, " +
            " NONYUSAKI_ID, NONYUSAKI_CD, NONYUSAKI_NK, " +
            " NONYUSAKIJUSYO1_NK, NONYUSAKIJUSYO2_NK, " +
            " PLANT_ID, SUPPORT_ID, PLANT_NK, " +
            " NONYUBUSYO_NK, NONYUTANTOSYA_NK, NONYUTEL_NO, NONYUFAX_NO, " +
            " KEIYAKUJIKANTAI, HOSYUHOHO, HOSYU_GKIN, " +
            " IRAITENPO_ID, IRAIMEISHO1_NK, IRAIMEISHO2_NK, IRAIMEISHO3_NK, IRAIMEISHO4_NK, " +
            " IRAITENPORYAKU_NK, IRAITANTO_NK, BIKO, YUKO_FLG, " +
            " CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY) " +
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,GETDATE(),?,GETDATE(),?)",
            k.getUkKikanId(), k.getUkKeiyakuId(),
            dtOrNull(k.getKaisiDt()), dtOrNull(k.getSyuryoDt()),
            k.getNonyusakiId(), k.getNonyusakiCd(), k.getNonyusakiNk(),
            k.getNonyusakijusyo1Nk(), k.getNonyusakijusyo2Nk(),
            k.getPlantId(), k.getSupportId(), k.getPlantNk(),
            k.getNonyubusyoNk(), k.getNonyutantosyaNk(), k.getNonyutelNo(), k.getNonyufaxNo(),
            k.getKeiyakujikantai(), k.getHosyuhoho(), k.getHosyuGkin(),
            k.getIraitenpoId(), k.getIraimeisho1Nk(), k.getIraimeisho2Nk(),
            k.getIraimeisho3Nk(), k.getIraimeisho4Nk(),
            k.getIraitenporyakuNk(), k.getIraitantoNk(), k.getBiko(), k.getYukoFlg(),
            loginUser, loginUser);
    }

    public void updateKikan(KikanTabForm k, String loginUser) {
        jdbc.update(
            "UPDATE MCM.MCM_UK_KIKAN SET " +
            "KAISI_DT=?, SYURYO_DT=?, " +
            "NONYUSAKI_ID=?, NONYUSAKI_CD=?, NONYUSAKI_NK=?, " +
            "NONYUSAKIJUSYO1_NK=?, NONYUSAKIJUSYO2_NK=?, " +
            "PLANT_ID=?, SUPPORT_ID=?, PLANT_NK=?, " +
            "NONYUBUSYO_NK=?, NONYUTANTOSYA_NK=?, NONYUTEL_NO=?, NONYUFAX_NO=?, " +
            "KEIYAKUJIKANTAI=?, HOSYUHOHO=?, HOSYU_GKIN=?, " +
            "IRAITENPO_ID=?, IRAIMEISHO1_NK=?, IRAIMEISHO2_NK=?, IRAIMEISHO3_NK=?, IRAIMEISHO4_NK=?, " +
            "IRAITENPORYAKU_NK=?, IRAITANTO_NK=?, BIKO=?, YUKO_FLG=?, " +
            "LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? " +
            "WHERE UK_KIKAN_ID=?",
            dtOrNull(k.getKaisiDt()), dtOrNull(k.getSyuryoDt()),
            k.getNonyusakiId(), k.getNonyusakiCd(), k.getNonyusakiNk(),
            k.getNonyusakijusyo1Nk(), k.getNonyusakijusyo2Nk(),
            k.getPlantId(), k.getSupportId(), k.getPlantNk(),
            k.getNonyubusyoNk(), k.getNonyutantosyaNk(), k.getNonyutelNo(), k.getNonyufaxNo(),
            k.getKeiyakujikantai(), k.getHosyuhoho(), k.getHosyuGkin(),
            k.getIraitenpoId(), k.getIraimeisho1Nk(), k.getIraimeisho2Nk(),
            k.getIraimeisho3Nk(), k.getIraimeisho4Nk(),
            k.getIraitenporyakuNk(), k.getIraitantoNk(), k.getBiko(), k.getYukoFlg(),
            loginUser, k.getUkKikanId());
    }

    // ===================================================================
    // 保存: MCM_UK_BRAND (INSERT / UPDATE)
    // ===================================================================

    public void insertBrand(BrandRowForm r, String loginUser) {
        jdbc.update(
            "INSERT INTO MCM.MCM_UK_BRAND " +
            "(UK_BRAND_ID, UK_KIKAN_ID, BRANDKOSEI_ID, BRAND_NK, BRANDSYOSAI_NK, KEIYAKUJIKANTAI, HOSYUHOHO, " +
            " CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY) " +
            "VALUES (?,?,?,?,?,?,?,GETDATE(),?,GETDATE(),?)",
            r.getUkBrandId(), r.getUkKikanId(), r.getBrandkoseiId(), r.getBrandNk(), r.getBrandsyosaiNk(),
            r.getKeiyakujikantai(), r.getHosyuhoho(), loginUser, loginUser);
    }

    public void insertMitsumoriLink(BrandRowForm r) {
        if (r.getUmMitsumoriId() == null) return;
        jdbc.update(
            "INSERT INTO MCM.MCM_UK_MITSUMORI_BRAND (UK_BRAND_ID, UM_MITSUMORI_ID) " +
            "VALUES (?,?)",
            r.getUkBrandId(), r.getUmMitsumoriId());
    }

    public void updateBrand(BrandRowForm r, String loginUser) {
        jdbc.update(
            "UPDATE MCM.MCM_UK_BRAND SET " +
            "KEIYAKUJIKANTAI=?, HOSYUHOHO=?, LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? " +
            "WHERE UK_BRAND_ID=?",
            r.getKeiyakujikantai(), r.getHosyuhoho(), loginUser, r.getUkBrandId());
    }

    // ===================================================================
    // 保存: MCM_UK_KIKIKOSEI (INSERT)
    // ===================================================================

    public void insertKosei(KoseiRowForm r, String loginUser) {
        jdbc.update(
            "INSERT INTO MCM.MCM_UK_KIKIKOSEI " +
            "(UK_KIKIKOSEI_ID, UK_BRAND_ID, KIKIKOSEI_ID, " +
            " KIKIKOSEI_NK, SET_NM, TANI, TEHAISEIBAN, CONTROLLER_FLG, HOSYUHOHO, HYOJIJUN, " +
            " CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY) " +
            "VALUES (?,?,?,?,?,?,?,?,?,?,GETDATE(),?,GETDATE(),?)",
            r.getUkKikoseiId(), r.getUkBrandId(), r.getKikikoseiId(),
            r.getKikikoseiNk(), r.getSetNm(), r.getTani(), r.getTehaiseiban(),
            r.getControllerFlg(), r.getHosyuhoho(), r.getHyojijun(),
            loginUser, loginUser);
    }

    // ===================================================================
    // 保存: MCM_UK_KIKIMEISAI (INSERT)
    // ===================================================================

    public void insertMeisai(MeisaiRowForm r, String loginUser) {
        jdbc.update(
            "INSERT INTO MCM.MCM_UK_KIKIMEISAI " +
            "(UK_KIKIMEISAI_ID, UK_KIKIKOSEI_ID, KIKIMEISAI_ID, " +
            " SEIZOMAKER_ID, SEIZOMAKER_NK, KIKIHINMEI_NK, KIKIKATASHIKI, SURYO_NM, " +
            " KEIYAKUNAIYO, KEIYAKU_NO, SERVICEKEITAI, " +
            " TORIHOSYUJIKAN_ID, DAIFUKUHOSYUJIKAN_ID, TENKENKAISU, TENKENYOBI, HOSYUHOHO, " +
            " HYOJIJUN, CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY) " +
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,GETDATE(),?,GETDATE(),?)",
            r.getUkKikimeisaiId(), r.getUkKikoseiId(), r.getKikimeisaiId(), r.getSeizomakerId(), r.getSeizomankerNk(),
            r.getKikihinmeiNk(), r.getKikikatashiki(), r.getSuryoNm(),
            r.getKeiyakunaiyo(), r.getKeiyakuNo(), r.getServicekeitai(),
            r.getTorihosyujikanId(), r.getDaifukuhosyujikanId(),
            r.getTenkenkaisu(), r.getTenkenyobi(), r.getHosyuhoho(),
            r.getHyojijun(), loginUser, loginUser);
    }

    // ===================================================================
    // 保存: MCM_UK_KOTAIMEISAI (INSERT)
    // ===================================================================

    public void insertKotai(KotaiRowForm r, String user) {
        jdbc.update("INSERT INTO MCM.MCM_UK_KOTAIMEISAI (UK_KOTAIMEISAI_ID,UK_KIKIMEISAI_ID,KOTAIKANRI_ID,KOTAI_NK,SERIAL_NO,ITIJINONYU_DT,SETCHIBASYO,KEIYAKUKIGEN_DT,ENCHOKEIYAKUKIGEN_DT,CREATED_DT,CREATED_BY,LASTUPDATE_DT,LASTUPDATE_BY) VALUES (?,?,?,?,?,?,?,?,?,GETDATE(),?,GETDATE(),?)",
            r.getUkKotaimeisaiId(),r.getUkKikimeisaiId(),r.getKotaikanriId(),r.getKotaiNk(),r.getSerialNo(),dtOrNull(r.getItizinonnyuDt()),r.getSetchibasyo(),dtOrNull(r.getKeiyakukigenDt()),dtOrNull(r.getEnchokeiyakukigenDt()),user,user);
    }

    // ===================================================================
    // 契約状態更新
    // ===================================================================

    public void updateJotai(BigDecimal ukKeiyakuId, String jotai, String loginUser) {
        jdbc.update(
            "UPDATE MCM.MCM_UK_KEIYAKU SET JOTAI=?, LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? " +
            "WHERE UK_KEIYAKU_ID=?",
            jotai, loginUser, ukKeiyakuId);
    }

    // ===================================================================
    // 次回更新日設定
    // jikaikosinDt = 最終タブの SYURYO_DT + 1日
    // ===================================================================

    public void updateJikaikosinDt(BigDecimal ukKeiyakuId, String syuryoDt, String loginUser) {
        LocalDate next = LocalDate.parse(syuryoDt, DT_FMT).plusDays(1);
        jdbc.update(
            "UPDATE MCM.MCM_UK_KEIYAKU SET JIKAIKOSIN_DT=?, LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? " +
            "WHERE UK_KEIYAKU_ID=?",
            java.sql.Date.valueOf(next), loginUser, ukKeiyakuId);
    }

    // ===================================================================
    // MCM_UM_KIHON_MITSUMORIから初期値取得
    // 新規KIKAN追加時: KEIYAKUJIKANTAI, HOSYUHOHO をコピー
    // ===================================================================

    public String[] findKihonMitsumoriDefaults(BigDecimal umKihonMitsumoriId) {
        List<String[]> list = jdbc.query(
            "SELECT KEIYAKUJIKANTAI, HOSYUHOHO FROM MCM.MCM_UM_KIHON_MITSUMORI " +
            "WHERE UM_KIHON_MITSUMORI_ID=?",
            new Object[]{umKihonMitsumoriId},
            (rs, rn) -> new String[]{rs.getString("KEIYAKUJIKANTAI"), rs.getString("HOSYUHOHO")});
        return list.isEmpty() ? new String[]{"", ""} : list.get(0);
    }

    // ===================================================================
    // MCM_MA_BRAND_KOSEIから表示名取得
    // ===================================================================

    public String[] findBrandKoseiNames(BigDecimal brandkoseiId) {
        List<String[]> list = jdbc.query(
            "SELECT BRAND_NK, BRANDSYOSAI_NK FROM MCM.MCM_MA_BRAND_KOSEI WHERE BRANDKOSEI_ID=?",
            new Object[]{brandkoseiId},
            (rs, rn) -> new String[]{rs.getString("BRAND_NK"), rs.getString("BRANDSYOSAI_NK")});
        return list.isEmpty() ? new String[]{"", ""} : list.get(0);
    }

    // ───────────────────────────────────────────────────────────────────
    // ユーティリティ: 日付文字列 → java.sql.Date (null/空は null を返す)
    // ───────────────────────────────────────────────────────────────────

    /** 同一期間の機器再選定。対象期間の子だけを削除してから再構築する。 */
    public void deleteEquipment(BigDecimal id) {
        String brands="SELECT UK_BRAND_ID FROM MCM.MCM_UK_BRAND WHERE UK_KIKAN_ID=?";
        String kosei="SELECT UK_KIKIKOSEI_ID FROM MCM.MCM_UK_KIKIKOSEI WHERE UK_BRAND_ID IN ("+brands+")";
        String meisai="SELECT UK_KIKIMEISAI_ID FROM MCM.MCM_UK_KIKIMEISAI WHERE UK_KIKIKOSEI_ID IN ("+kosei+")";
        jdbc.update("DELETE FROM MCM.MCM_UK_KOTAIMEISAI WHERE UK_KIKIMEISAI_ID IN ("+meisai+")",id);
        jdbc.update("DELETE FROM MCM.MCM_UK_KIKIMEISAI WHERE UK_KIKIKOSEI_ID IN ("+kosei+")",id);
        jdbc.update("DELETE FROM MCM.MCM_UK_KIKIKOSEI WHERE UK_BRAND_ID IN ("+brands+")",id);
        jdbc.update("DELETE FROM MCM.MCM_UK_MITSUMORI_BRAND WHERE UK_BRAND_ID IN ("+brands+")",id);
        jdbc.update("DELETE FROM MCM.MCM_UK_BRAND WHERE UK_KIKAN_ID=?",id);
    }
    private java.sql.Date dtOrNull(String yyyymmdd) {
        if (yyyymmdd == null || yyyymmdd.isBlank()) return null;
        try {
            return java.sql.Date.valueOf(LocalDate.parse(yyyymmdd, DT_FMT));
        } catch (Exception e) {
            return null;
        }
    }
}
