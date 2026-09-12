package com.daifuku.mcm.repository;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Repository;

import com.daifuku.mcm.form.Mcm2003uForm;

/**
 * 【変換元】Mcm2003uDataSet.Designer.vb / Mcm2003uScreen.vb
 * MCM2003U 店舗見積内容基本設定リポジトリ
 *
 * SQL Server 構文使用。全テーブル名に MCM. スキーマ付与。
 * Excel帳票出力・メール送信・SP_UMストアドプロシージャはWeb版省略。
 */
@Repository
public class Mcm2003uRepository {

    @Autowired
    private JdbcTemplate jdbc;

    private static final DateTimeFormatter DT_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");
    private static final DateTimeFormatter DT_FMT_DATETIME = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");

    // ===================================================================
    // 基本見積ヘッダー取得
    // 【変換元】Mcm2003uDataSet — MCM_UM_KIHON_MITSUMORI SELECT
    // ===================================================================

    public Mcm2003uForm findKihonMitsumori(BigDecimal umKihonMitsumoriId) {
        return jdbc.queryForObject(
            "SELECT UM_KIHON_MITSUMORI_ID, UM_MITSUMORI_NO, NONYUSAKI_ID, NONYUSAKI_CD, NONYUSAKI_NK, " +
            "       NONYUSAKIJUSYO1_NK, NONYUSAKIJUSYO2_NK, PLANT_ID, SUPPORT_ID, PLANT_NK, " +
            "       NONYUBUSYO_NK, NONYUTANTOSYA_NK, NONYUTEL_NO, NONYUFAX_NO, " +
            "       MITSUMORI_DT, MITSUMORISAKUSEISYA_NK, MITSUMORIKIGEN, " +
            "       IRAITENPO_ID, IRAIMEISHO1_NK, IRAIMEISHO2_NK, IRAIMEISHO3_NK, IRAIMEISHO4_NK, " +
            "       IRAITENPORYAKU_NK, IRAITANTO_NK, " +
            "       SOFUTENPO_ID, SOFUMEISHO1_NK, SOFUMEISHO2_NK, SOFUMEISHO3_NK, SOFUMEISHO4_NK, " +
            "       SOFUTENPORYAKU_NK, SOFUTANTO_NK, " +
            "       KEIYAKUJIKANTAI, HOSYUHOHO, JOTAI, MITSUMORILEVEL, SYOUNIN_JOTAI, BIKO, MITSUMORI_JOUKEN, " +
            "       CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY " +
            "FROM MCM.MCM_UM_KIHON_MITSUMORI " +
            "WHERE UM_KIHON_MITSUMORI_ID = ?",
            new Object[]{umKihonMitsumoriId},
            (rs, rowNum) -> {
                Mcm2003uForm form = new Mcm2003uForm();
                form.setUmKihonMitsumoriId(rs.getBigDecimal("UM_KIHON_MITSUMORI_ID"));
                form.setUmMitsumoriNo(rs.getString("UM_MITSUMORI_NO"));
                form.setNonyusakiId(rs.getString("NONYUSAKI_ID"));
                form.setNonyusakiCd(rs.getString("NONYUSAKI_CD"));
                form.setNonyusakiNk(rs.getString("NONYUSAKI_NK"));
                form.setNonyusakijusyo1Nk(rs.getString("NONYUSAKIJUSYO1_NK"));
                form.setNonyusakijusyo2Nk(rs.getString("NONYUSAKIJUSYO2_NK"));
                form.setPlantId(rs.getString("PLANT_ID"));
                form.setSupportId(rs.getString("SUPPORT_ID"));
                form.setPlantNk(rs.getString("PLANT_NK"));
                form.setNonyubusyoNk(rs.getString("NONYUBUSYO_NK"));
                form.setNonyutantosyaNk(rs.getString("NONYUTANTOSYA_NK"));
                form.setNonyutelNo(rs.getString("NONYUTEL_NO"));
                form.setNonyufaxNo(rs.getString("NONYUFAX_NO"));
                LocalDateTime mitsumoriDt = rs.getObject("MITSUMORI_DT", LocalDateTime.class);
                form.setMitsumoriDt(mitsumoriDt != null ? DT_FMT.format(mitsumoriDt) : "");
                form.setMitsumorisakuseisyaNk(rs.getString("MITSUMORISAKUSEISYA_NK"));
                form.setMitsumorigiken(rs.getString("MITSUMORIKIGEN"));
                form.setIraitenpoId(rs.getBigDecimal("IRAITENPO_ID"));
                form.setIraimeisho1Nk(rs.getString("IRAIMEISHO1_NK"));
                form.setIraimeisho2Nk(rs.getString("IRAIMEISHO2_NK"));
                form.setIraimeisho3Nk(rs.getString("IRAIMEISHO3_NK"));
                form.setIraimeisho4Nk(rs.getString("IRAIMEISHO4_NK"));
                form.setIraitenporyakuNk(rs.getString("IRAITENPORYAKU_NK"));
                form.setIraitantoNk(rs.getString("IRAITANTO_NK"));
                form.setSofutenpoId(rs.getBigDecimal("SOFUTENPO_ID"));
                form.setSofumeisho1Nk(rs.getString("SOFUMEISHO1_NK"));
                form.setSofumeisho2Nk(rs.getString("SOFUMEISHO2_NK"));
                form.setSofumeisho3Nk(rs.getString("SOFUMEISHO3_NK"));
                form.setSofumeisho4Nk(rs.getString("SOFUMEISHO4_NK"));
                form.setSofutenporyakuNk(rs.getString("SOFUTENPORYAKU_NK"));
                form.setSofutantoNk(rs.getString("SOFUTANTO_NK"));
                form.setKeiyakujikantai(rs.getString("KEIYAKUJIKANTAI"));
                form.setHosyuhoho(rs.getString("HOSYUHOHO"));
                form.setJotai(rs.getString("JOTAI"));
                form.setMitsumoriLevel(rs.getBigDecimal("MITSUMORILEVEL")==null?"":rs.getBigDecimal("MITSUMORILEVEL").stripTrailingZeros().toPlainString());
                form.setSyouninJotai(rs.getString("SYOUNIN_JOTAI"));
                form.setBiko(rs.getString("BIKO"));
                form.setMitsumoriJouken(rs.getString("MITSUMORI_JOUKEN"));
                LocalDateTime createdDt = rs.getObject("CREATED_DT", LocalDateTime.class);
                form.setCreatedDt(createdDt != null ? DT_FMT_DATETIME.format(createdDt) : "");
                form.setCreatedBy(rs.getString("CREATED_BY"));
                LocalDateTime lastupdateDt = rs.getObject("LASTUPDATE_DT", LocalDateTime.class);
                form.setLastupdateDt(lastupdateDt != null ? lastupdateDt.toString() : "");
                form.setLastupdateBy(rs.getString("LASTUPDATE_BY"));
                return form;
            });
    }

    // ===================================================================
    // 契約金額変動時期取得
    // 【変換元】Mcm2003uDataSet — MCM_UM_MITSUMORI SELECT
    // ===================================================================

    public List<Mcm2003uForm.MitsumoriRowForm> findMitsumoriRows(BigDecimal umKihonMitsumoriId) {
        return jdbc.query(
            "SELECT UM_MITSUMORI_ID, UM_KIHON_MITSUMORI_ID, KAISI_DT, SYURYO_DT, " +
            "       (SELECT SUM(B.HARDHOSYU_KIN) FROM MCM.MCM_UM_BRAND B WHERE B.UM_MITSUMORI_ID=MCM_UM_MITSUMORI.UM_MITSUMORI_ID) AS HARDHOSYU_KIN, MITSUMORI_GKIN, 0 AS CHECK_FLG " +
            "FROM MCM.MCM_UM_MITSUMORI " +
            "WHERE UM_KIHON_MITSUMORI_ID = ? " +
            "ORDER BY KAISI_DT,UM_MITSUMORI_ID",
            new Object[]{umKihonMitsumoriId},
            (rs, rowNum) -> {
                Mcm2003uForm.MitsumoriRowForm row = new Mcm2003uForm.MitsumoriRowForm();
                row.setUmMitsumoriId(rs.getBigDecimal("UM_MITSUMORI_ID"));
                row.setUmKihonMitsumoriId(rs.getBigDecimal("UM_KIHON_MITSUMORI_ID"));
                LocalDateTime kaisiDt = rs.getObject("KAISI_DT", LocalDateTime.class);
                row.setKaisiDt(kaisiDt != null ? DT_FMT.format(kaisiDt) : "");
                LocalDateTime syuryoDt = rs.getObject("SYURYO_DT", LocalDateTime.class);
                row.setSyuryoDt(syuryoDt != null ? DT_FMT.format(syuryoDt) : "");
                row.setHardhosyuKin(rs.getBigDecimal("HARDHOSYU_KIN"));
                row.setMitsumoriGkin(rs.getBigDecimal("MITSUMORI_GKIN"));
                row.setCheckFlg("1".equals(rs.getString("CHECK_FLG")));
                return row;
            });
    }

    // ===================================================================
    // 添付ファイル取得
    // 【変換元】Mcm2003uDataSet — MCM_UM_TENPU SELECT
    // ===================================================================

    public List<Mcm2003uForm.TenpuRowForm> findTenpuRows(BigDecimal umKihonMitsumoriId) {
        return jdbc.query(
            "SELECT UM_TENPU_ID, UM_KIHON_MITSUMORI_ID, TENPUFILE_NK, DIRECTORY, BIKO " +
            "FROM MCM.MCM_UM_TENPU " +
            "WHERE UM_KIHON_MITSUMORI_ID = ? " +
            "ORDER BY UM_TENPU_ID",
            new Object[]{umKihonMitsumoriId},
            (rs, rowNum) -> {
                Mcm2003uForm.TenpuRowForm row = new Mcm2003uForm.TenpuRowForm();
                row.setUmTenpuId(rs.getBigDecimal("UM_TENPU_ID"));
                row.setUmKihonMitsumoriId(rs.getBigDecimal("UM_KIHON_MITSUMORI_ID"));
                row.setTenpufileNk(rs.getString("TENPUFILE_NK"));
                row.setDirectory(rs.getString("DIRECTORY"));
                row.setBiko(rs.getString("BIKO"));
                return row;
            });
    }

    // ===================================================================
    // ブランドタブ取得
    // 【変換元】Mcm2003uDataSet — MCM_UM_KIHON_BRAND SELECT
    // ===================================================================

    public List<Mcm2003uForm.BrandRowForm> findBrandRows(BigDecimal umKihonMitsumoriId) {
        return jdbc.query(
            "SELECT UMB.UM_KIHON_BRAND_ID, UMB.UM_KIHON_MITSUMORI_ID, UMB.BRAND_NK, UMB.BRANDSYOSAI_NK, " +
            "       UMB.KEIYAKUJIKANTAI, UMB.HARDHOSYU_KIN, ISNULL(MAD.HYOJIJUN, 0) AS HYOJIJUN, " +
            "       UMB.SYSTEMSEKKEI_KIN, UMB.KIHONSEKKEI_KIN, UMB.PROGRAMSAKUSEI_KIN, " +
            "       ISNULL(UMB.SYSTEMSEKKEI_KIN,0) + ISNULL(UMB.KIHONSEKKEI_KIN,0) + ISNULL(UMB.PROGRAMSAKUSEI_KIN,0) AS SYSTEMSUPPORTSYOKEI_KIN, " +
            "       UMB.SYSTEMSUPPORT_KIN, UMB.DTSSUPPORT_KIN, UMB.DAIFUKUGIJUTSU_KIN, UMB.SOFTHOSHU_KIN, " +
            "       UMB.HOSEISYSTEMSUPPORT_KIN, UMB.HOSEIDTSSUPPORT_KIN, UMB.HOSEIDAIFUKUGIJUTSU_KIN, UMB.HOSEISOFTHOSHU_KIN, " +
            "       UMB.SOFTHOSYUHOHO, UMB.SOFT_FLG, UMB.DREMOS_FLG, UMB.REMOTE_FLG, " +
            "       UMB.CREATED_DT, UMB.CREATED_BY, UMB.LASTUPDATE_DT, UMB.LASTUPDATE_BY " +
            "FROM MCM.MCM_UM_KIHON_BRAND UMB LEFT JOIN MCM.MCM_MA_BRAND_KOSEI MAD ON UMB.BRANDKOSEI_ID = MAD.BRANDKOSEI_ID " +
            "WHERE UMB.UM_KIHON_MITSUMORI_ID = ? " +
            "ORDER BY ISNULL(MAD.HYOJIJUN, 0)",
            new Object[]{umKihonMitsumoriId},
            (rs, rowNum) -> {
                Mcm2003uForm.BrandRowForm row = new Mcm2003uForm.BrandRowForm();
                row.setUmKihonBrandId(rs.getBigDecimal("UM_KIHON_BRAND_ID"));
                row.setUmKihonMitsumoriId(rs.getBigDecimal("UM_KIHON_MITSUMORI_ID"));
                row.setBrandNk(rs.getString("BRAND_NK"));
                row.setBrandsyosaiNk(rs.getString("BRANDSYOSAI_NK"));
                row.setKeiyakujikantai(rs.getString("KEIYAKUJIKANTAI"));
                row.setHardhosyuKin(rs.getBigDecimal("HARDHOSYU_KIN"));
                row.setHyojijun(rs.getInt("HYOJIJUN"));
                row.setSystemSekkeiKin(rs.getBigDecimal("SYSTEMSEKKEI_KIN"));
                row.setKihonSekkeiKin(rs.getBigDecimal("KIHONSEKKEI_KIN"));
                row.setProgramSakuseiKin(rs.getBigDecimal("PROGRAMSAKUSEI_KIN"));
                row.setSystemSupportKin(rs.getBigDecimal("SYSTEMSUPPORT_KIN"));
                row.setDtsSupportKin(rs.getBigDecimal("DTSSUPPORT_KIN"));
                row.setDaifukuGijutsuKin(rs.getBigDecimal("DAIFUKUGIJUTSU_KIN"));
                row.setSoftHosyuKin(rs.getBigDecimal("SOFTHOSHU_KIN"));
                row.setHoseiSystemSupportKin(rs.getBigDecimal("HOSEISYSTEMSUPPORT_KIN"));
                row.setHoseiDtsSupportKin(rs.getBigDecimal("HOSEIDTSSUPPORT_KIN"));
                row.setHoseiDaifukuGijutsuKin(rs.getBigDecimal("HOSEIDAIFUKUGIJUTSU_KIN"));
                row.setHoseiSoftHosyuKin(rs.getBigDecimal("HOSEISOFTHOSHU_KIN"));
                row.setSoftHosyuhoho(rs.getString("SOFTHOSYUHOHO"));
                row.setSoftFlg(rs.getBigDecimal("SOFT_FLG") != null && rs.getBigDecimal("SOFT_FLG").signum() != 0);
                row.setDremosFlg(rs.getBigDecimal("DREMOS_FLG") != null && rs.getBigDecimal("DREMOS_FLG").signum() != 0);
                row.setRemoteFlg(rs.getBigDecimal("REMOTE_FLG") != null && rs.getBigDecimal("REMOTE_FLG").signum() != 0);
                row.setSystemSupportSyokeiKin(rs.getBigDecimal("SYSTEMSUPPORTSYOKEI_KIN"));
                LocalDateTime bCreatedDt = rs.getObject("CREATED_DT", LocalDateTime.class);
                row.setCreatedDt(bCreatedDt != null ? DT_FMT_DATETIME.format(bCreatedDt) : "");
                row.setCreatedBy(rs.getString("CREATED_BY"));
                LocalDateTime bLastupdateDt = rs.getObject("LASTUPDATE_DT", LocalDateTime.class);
                row.setLastupdateDt(bLastupdateDt != null ? DT_FMT_DATETIME.format(bLastupdateDt) : "");
                row.setLastupdateBy(rs.getString("LASTUPDATE_BY"));
                return row;
            });
    }

    // ===================================================================
    // 機器構成パターン取得
    // 【変換元】Mcm2003uDataSet — MCM_UM_KIKIKOSEI SELECT
    // ===================================================================

    public List<Mcm2003uForm.KoseiRowForm> findKoseiRows(BigDecimal umKihonMitsumoriId) {
        return jdbc.query(
            "SELECT UMC.UM_KIKIKOSEI_ID, UMC.UM_KIHON_BRAND_ID, UMC.KIKIKOSEI_NK, UMC.SET_NM, UMC.TANI, " +
            "       UMC.TEHAISEIBAN, UMC.HOSYUHOHO, UMC.BIKO, ISNULL(MAE.HYOJIJUN, 0) AS HYOJIJUN " +
            "FROM MCM.MCM_UM_KIKIKOSEI UMC " +
            "  JOIN MCM.MCM_UM_KIHON_BRAND UMB ON UMC.UM_KIHON_BRAND_ID = UMB.UM_KIHON_BRAND_ID " +
            "  LEFT JOIN MCM.MCM_MA_KIKIKOSEI MAE ON UMC.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID " +
            "WHERE UMB.UM_KIHON_MITSUMORI_ID = ? " +
            "ORDER BY ISNULL(MAE.HYOJIJUN, 0)",
            new Object[]{umKihonMitsumoriId},
            (rs, rowNum) -> {
                Mcm2003uForm.KoseiRowForm row = new Mcm2003uForm.KoseiRowForm();
                row.setUmKikikoseiId(rs.getBigDecimal("UM_KIKIKOSEI_ID"));
                row.setUmKihonBrandId(rs.getBigDecimal("UM_KIHON_BRAND_ID"));
                row.setKikikoseiNk(rs.getString("KIKIKOSEI_NK"));
                row.setSetNm(rs.getBigDecimal("SET_NM"));
                row.setTani(rs.getString("TANI"));
                row.setTehaiseiban(rs.getString("TEHAISEIBAN"));
                row.setHosyuhoho(rs.getString("HOSYUHOHO"));
                row.setBiko(rs.getString("BIKO"));
                row.setHyojijun(rs.getInt("HYOJIJUN"));
                return row;
            });
    }

    // ===================================================================
    // 機器明細取得
    // 【変換元】Mcm2003uDataSet — MCM_UM_KIKIMEISAI SELECT
    // ===================================================================

    public List<Mcm2003uForm.MeisaiRowForm> findMeisaiRows(BigDecimal umKihonMitsumoriId) {
        return jdbc.query(
            "SELECT UMD.UM_KIKIMEISAI_ID, UMD.UM_KIKIKOSEI_ID, UMD.SEIZOMAKER_NK, UMD.KIKIHINMEI_NK, " +
            "       UMD.KIKIKATASHIKI, UMD.SURYO_NM, UMD.SOSU_NM, UMD.BIKO " +
            "FROM MCM.MCM_UM_KIKIMEISAI UMD " +
            "  JOIN MCM.MCM_UM_KIKIKOSEI UMC ON UMD.UM_KIKIKOSEI_ID = UMC.UM_KIKIKOSEI_ID " +
            "  JOIN MCM.MCM_UM_KIHON_BRAND UMB ON UMC.UM_KIHON_BRAND_ID = UMB.UM_KIHON_BRAND_ID " +
            "  LEFT JOIN MCM.MCM_MA_KIKIMEISAI MAF ON UMD.KIKIMEISAI_ID = MAF.KIKIMEISAI_ID " +
            "  LEFT JOIN MCM.MCM_MA_KIKIKOSEI MAE ON UMC.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID " +
            "WHERE UMB.UM_KIHON_MITSUMORI_ID = ? " +
            "ORDER BY ISNULL(MAE.HYOJIJUN, 0), ISNULL(MAF.HYOJIJUN, 0)",
            new Object[]{umKihonMitsumoriId},
            (rs, rowNum) -> {
                Mcm2003uForm.MeisaiRowForm row = new Mcm2003uForm.MeisaiRowForm();
                row.setUmKikimeisaiId(rs.getBigDecimal("UM_KIKIMEISAI_ID"));
                row.setUmKikikoseiId(rs.getBigDecimal("UM_KIKIKOSEI_ID"));
                row.setSeizomakerNk(rs.getString("SEIZOMAKER_NK"));
                row.setKikihinmeiNk(rs.getString("KIKIHINMEI_NK"));
                row.setKikikatashiki(rs.getString("KIKIKATASHIKI"));
                row.setSuryoNm(rs.getBigDecimal("SURYO_NM"));
                row.setSosuNm(rs.getBigDecimal("SOSU_NM"));
                row.setBiko(rs.getString("BIKO"));
                return row;
            });
    }

    // ===================================================================
    // 基本見積ヘッダー更新
    // 【変換元】SenteiNashiUpdate() / SenteiAriUpdate() → MCM_UM_KIHON_MITSUMORI UPDATE
    // ===================================================================

    /** 管理者が変更した宛先はマスタから補完し、未変更の宛先表記は保持する。 */
    public void prepareHeaderChanges(Mcm2003uForm form,Mcm2003uForm current){
        if(!java.util.Objects.equals(form.getUmMitsumoriNo(),current.getUmMitsumoriNo())){
            String no=form.getUmMitsumoriNo();if(no==null||no.isBlank()||no.length()>20)throw new IllegalStateException("見積NOは1～20文字で入力してください。");
            no=no.trim();if(jdbc.queryForObject("SELECT COUNT(*) FROM MCM.MCM_UM_KIHON_MITSUMORI WITH (UPDLOCK,HOLDLOCK) WHERE UM_MITSUMORI_NO=? AND UM_KIHON_MITSUMORI_ID<>?",Integer.class,no,form.getUmKihonMitsumoriId())>0)throw new IllegalStateException("同じ見積NOが登録されています。");
            form.setUmMitsumoriNo(no);
        }
        if(!sameAddress(form.getIraitenpoId(),current.getIraitenpoId()))fillAddress(form,"Irai",form.getIraitenpoId());
        if(!sameAddress(form.getSofutenpoId(),current.getSofutenpoId()))fillAddress(form,"Sofu",form.getSofutenpoId());
    }
    private boolean sameAddress(BigDecimal a,BigDecimal b){return a==null?b==null:b!=null&&a.compareTo(b)==0;}
    private void fillAddress(Mcm2003uForm form,String prefix,BigDecimal id){
        java.util.Map<String,Object> row=java.util.Map.of();
        if(id!=null&&id.signum()!=0){var rows=jdbc.queryForList("SELECT * FROM MCM.MCM_MA_TENPO WHERE TENPO_ID=?",id);if(rows.size()!=1)throw new IllegalStateException("宛先の店舗・事業所を選択し直してください。");row=rows.get(0);}
        var properties=new org.springframework.beans.BeanWrapperImpl(form);
        for(String column:java.util.List.of("MEISHO1_NK","MEISHO2_NK","MEISHO3_NK","MEISHO4_NK","TENPORYAKU_NK")){
            String field=prefix.toLowerCase()+(column.startsWith("MEISHO")?"meisho"+column.charAt(6)+"Nk":"tenporyakuNk");
            properties.setPropertyValue(field,row.getOrDefault(column,""));
        }
    }

    public void updateKihonMitsumori(Mcm2003uForm form, String loginUser) {
        jdbc.update(
            "UPDATE MCM.MCM_UM_KIHON_MITSUMORI " +
            "SET UM_MITSUMORI_NO=?, MITSUMORI_DT=?, " +
            "    MITSUMORISAKUSEISYA_NK=?, MITSUMORIKIGEN=?, " +
            "    IRAITENPO_ID=?, IRAIMEISHO1_NK=?, IRAIMEISHO2_NK=?, IRAIMEISHO3_NK=?, IRAIMEISHO4_NK=?, " +
            "    IRAITENPORYAKU_NK=?, IRAITANTO_NK=?, " +
            "    SOFUTENPO_ID=?, SOFUMEISHO1_NK=?, SOFUMEISHO2_NK=?, SOFUMEISHO3_NK=?, SOFUMEISHO4_NK=?, " +
            "    SOFUTENPORYAKU_NK=?, SOFUTANTO_NK=?, " +
            "    KEIYAKUJIKANTAI=?, HOSYUHOHO=?, BIKO=?, MITSUMORI_JOUKEN=?, MITSUMORILEVEL=?, " +
            "    LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? " +
            "WHERE UM_KIHON_MITSUMORI_ID=?",
            form.getUmMitsumoriNo(),
            form.getMitsumoriDt()==null || form.getMitsumoriDt().isBlank()?null:java.sql.Date.valueOf(com.daifuku.mcm.common.CustomerScreenSupport.date(form.getMitsumoriDt())),
            form.getMitsumorisakuseisyaNk(), form.getMitsumorigiken(),
            form.getIraitenpoId(),
            form.getIraimeisho1Nk(), form.getIraimeisho2Nk(),
            form.getIraimeisho3Nk(), form.getIraimeisho4Nk(),
            form.getIraitenporyakuNk(), form.getIraitantoNk(),
            form.getSofutenpoId(),
            form.getSofumeisho1Nk(), form.getSofumeisho2Nk(),
            form.getSofumeisho3Nk(), form.getSofumeisho4Nk(),
            form.getSofutenporyakuNk(), form.getSofutantoNk(),
            form.getKeiyakujikantai(), form.getHosyuhoho(),
            form.getBiko(), form.getMitsumoriJouken(), form.getMitsumoriLevel(),
            loginUser,
            form.getUmKihonMitsumoriId());
    }

    // ===================================================================
    // 状態更新（見積発行: JOTAI=1）
    // 【変換元】HakkoButton_Click → JOTAI = JOTAI_MITSUMORI
    // ===================================================================

    public void updateJotai(BigDecimal umKihonMitsumoriId, String jotai, String loginUser) {
        jdbc.update(
            "UPDATE MCM.MCM_UM_KIHON_MITSUMORI " +
            "SET JOTAI=?, LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? " +
            "WHERE UM_KIHON_MITSUMORI_ID=?",
            jotai, loginUser, umKihonMitsumoriId);
    }

    // ===================================================================
    // 承認状態更新（申請: SYOUNIN_JOTAI=1）
    // 【変換元】SinseiButton_Click → SYOUNIN_JOTAI = SHONINJOTAI_SHINSACHU_CD
    // ===================================================================

    public void updateSyouninJotai(BigDecimal umKihonMitsumoriId, String syouninJotai, String loginUser) {
        jdbc.update(
            "UPDATE MCM.MCM_UM_KIHON_MITSUMORI " +
            "SET SYOUNIN_JOTAI=?, LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? " +
            "WHERE UM_KIHON_MITSUMORI_ID=?",
            syouninJotai, loginUser, umKihonMitsumoriId);
    }

    // ===================================================================
    // 添付ファイル件数取得
    // 【変換元】SinseiButton_Click — tenpuCount チェック
    // ===================================================================

    public int countTenpu(BigDecimal umKihonMitsumoriId) {
        Integer count = jdbc.queryForObject(
            "SELECT COUNT(*) FROM MCM.MCM_UM_TENPU WHERE UM_KIHON_MITSUMORI_ID=?",
            Integer.class, umKihonMitsumoriId);
        return count != null ? count : 0;
    }
    /** 管理者解除は担当者マスターと画面別更新権限を毎回照合する。 */
    public boolean canReleaseLock(String loginId) {
        return jdbc.queryForObject("""
            SELECT COUNT(*) FROM MCM.MCM_MO_TANTO A
            WHERE A.LOGIN_ID=? AND A.MAINTENANCE_FLG=1 AND EXISTS (
                SELECT 1 FROM MCM.MCM_MO_TANTOKENGEN B
                JOIN MCM.MCM_MO_KENGENKOSEI C ON C.KENGENBUNRUI_ID=B.KENGENBUNRUI_ID AND C.RIYOKENGEN_KBN=B.RIYOKENGEN_KBN
                JOIN MCM.MCM_MO_KENGENBUNRUI D ON D.KENGENBUNRUI_ID=C.KENGENBUNRUI_ID
                JOIN MCM.MCM_MO_KINO E ON E.KINO_ID=C.KINO_ID
                WHERE B.TANTO_ID=A.TANTO_ID AND E.KINO_ID='MCM2003U' AND C.RIYOKENGEN_KBN='2')
            """,Integer.class,loginId)==1;
    }

    public void lock(BigDecimal id) {
        jdbc.queryForObject("SELECT UM_KIHON_MITSUMORI_ID FROM MCM.MCM_UM_KIHON_MITSUMORI WITH (UPDLOCK,HOLDLOCK) WHERE UM_KIHON_MITSUMORI_ID=?",BigDecimal.class,id);
    }
    public void insertAttachment(BigDecimal estimate,String name,String directory,String user) {
        BigDecimal id=jdbc.queryForObject("SELECT ISNULL(MAX(UM_TENPU_ID),0)+1 FROM MCM.MCM_UM_TENPU WITH (UPDLOCK,HOLDLOCK)",BigDecimal.class);
        jdbc.update("INSERT INTO MCM.MCM_UM_TENPU (UM_TENPU_ID,UM_KIHON_MITSUMORI_ID,TENPUFILE_NK,DIRECTORY,CREATED_DT,CREATED_BY,LASTUPDATE_DT,LASTUPDATE_BY) VALUES (?,?,?,?,GETDATE(),?,GETDATE(),?)",id,estimate,name,directory,user,user);
    }
    public void deleteAttachment(BigDecimal estimate,BigDecimal id) {
        if(jdbc.update("DELETE FROM MCM.MCM_UM_TENPU WHERE UM_KIHON_MITSUMORI_ID=? AND UM_TENPU_ID=?",estimate,id)!=1)
            throw new IllegalStateException("添付資料が変更されています。再読み込みしてください。");
    }
    /** ヘッダーのロック下で、編集したブランドの保守方法のみを更新する。 */
    public void updateBrandMethods(Mcm2003uForm form,String user) {
        for(var brand:form.getBrandRows())if(brand.isSoftHosyuhohoEdited()) {
            if(brand.getSoftHosyuhoho()!=null && brand.getSoftHosyuhoho().length()>4000)throw new IllegalStateException("保守方法は4000文字以内で入力してください。");
            var current=findBrandRows(form.getUmKihonMitsumoriId()).stream().filter(b->b.getUmKihonBrandId().compareTo(brand.getUmKihonBrandId())==0).findFirst().orElseThrow(()->new IllegalStateException("ブランド情報が変更されています。再読み込みしてください。"));
            if(!java.util.Objects.equals(current.getLastupdateDt(),brand.getLastupdateDt()) || !java.util.Objects.equals(current.getSoftHosyuhoho(),brand.getOriginalSoftHosyuhoho()))throw new IllegalStateException("ブランド情報が更新されています。再読み込みしてください。");
            if(jdbc.update("UPDATE MCM.MCM_UM_KIHON_BRAND SET SOFTHOSYUHOHO=?,LASTUPDATE_DT=GETDATE(),LASTUPDATE_BY=? WHERE UM_KIHON_MITSUMORI_ID=? AND UM_KIHON_BRAND_ID=? AND ((SOFTHOSYUHOHO IS NULL AND ? IS NULL) OR SOFTHOSYUHOHO=?)",brand.getSoftHosyuhoho(),user,form.getUmKihonMitsumoriId(),brand.getUmKihonBrandId(),brand.getOriginalSoftHosyuhoho(),brand.getOriginalSoftHosyuhoho())!=1)throw new IllegalStateException("ブランド情報を保存できませんでした。");
        }
        updateEditedRows(form,user);
    }
    private void updateEditedRows(Mcm2003uForm form,String user) {
        var brands=findBrandRows(form.getUmKihonMitsumoriId());var configs=findKoseiRows(form.getUmKihonMitsumoriId());
        for(var b:form.getBrandRows())if(!b.getEditedOriginals().isEmpty()) {
            var current=brands.stream().filter(r->r.getUmKihonBrandId().compareTo(b.getUmKihonBrandId())==0).findFirst().orElseThrow(()->new IllegalStateException("ブランドが変更されています。"));
            updateEditedRow("MCM_UM_KIHON_BRAND","UM_KIHON_BRAND_ID",b.getUmKihonBrandId(),b,current,com.daifuku.mcm.common.Mcm2003uEdits.BRAND,user);
        }
        for(var k:form.getKoseiRows())if(!k.getEditedOriginals().isEmpty()) {
            var current=configs.stream().filter(r->r.getUmKikikoseiId().compareTo(k.getUmKikikoseiId())==0).findFirst().orElseThrow(()->new IllegalStateException("機器構成が変更されています。"));
            updateEditedRow("MCM_UM_KIKIKOSEI","UM_KIKIKOSEI_ID",k.getUmKikikoseiId(),k,current,com.daifuku.mcm.common.Mcm2003uEdits.KOSEI,user);
        }
    }
    private void updateEditedRow(String table,String pk,BigDecimal id,Object edited,Object current,java.util.Map<String,String> columns,String user) {
        var setters=new java.util.ArrayList<String>();var filters=new java.util.ArrayList<String>();var values=new java.util.ArrayList<Object>();var originals=new java.util.ArrayList<Object>();
        for(var e:com.daifuku.mcm.common.Mcm2003uEdits.originals(edited).entrySet()) {
            String column=columns.get(e.getKey());if(column==null)throw new IllegalStateException("編集項目が不正です。");
            if(!com.daifuku.mcm.common.Mcm2003uEdits.equal(com.daifuku.mcm.common.Mcm2003uEdits.get(current,e.getKey()),e.getValue()))throw new IllegalStateException("編集中の項目が更新されています。再読み込みして確認してください。");
            setters.add(column+"=?");values.add(com.daifuku.mcm.common.Mcm2003uEdits.databaseValue(com.daifuku.mcm.common.Mcm2003uEdits.get(edited,e.getKey())));
            if(e.getValue()==null)filters.add(column+" IS NULL");
            else if(Boolean.FALSE.equals(e.getValue()))filters.add("("+column+" IS NULL OR "+column+"=0)");
            else{filters.add(column+"=?");originals.add(com.daifuku.mcm.common.Mcm2003uEdits.databaseValue(e.getValue()));}
        }
        values.add(user);values.add(id);values.addAll(originals);
        if(jdbc.update("UPDATE MCM."+table+" SET "+String.join(",",setters)+",LASTUPDATE_DT=GETDATE(),LASTUPDATE_BY=? WHERE "+pk+"=? AND "+String.join(" AND ",filters),values.toArray())!=1)throw new IllegalStateException("編集中の項目が更新されています。再読み込みして確認してください。");
    }
}
