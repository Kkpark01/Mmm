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
 * 実列のみ更新し、集計はサービスでトランザクション内に実施。
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
            "SELECT M.SEIZOMAKER_NK,M.KIKIKATASHIKI,ISNULL(V.SOSU_NM,M.SOSU_NM) AS SOSU_NM,T.CREATED_DT,T.CREATED_BY,T.LASTUPDATE_DT,T.LASTUPDATE_BY,T.UM_TANKA_ID,K.UM_KIKIKOSEI_ID,ISNULL(V.SURYO_NM,M.SURYO_NM) AS SURYO_NM,T.HYOJUN_KIN, " +
            "ISNULL(ISNULL(V.SURYO_NM,M.SURYO_NM),0)*ISNULL(T.HYOJUN_KIN,0) AS HYOJUNSHOKEI_KIN, " +
            "ISNULL(ISNULL(V.SURYO_NM,M.SURYO_NM),0)*ISNULL(T.SIKIRI_KIN,0) AS SIKIRISHOKEI_KIN,ISNULL(A.CONTROLLER_FLG,K.CONTROLLER_FLG) CONTROLLER_FLG,K.KIKIKOSEI_NK,T.SIKIRI_KIN,T.TORIHOSYUJIKAN_ID,T.DAIFUKUHOSYUJIKAN_ID,T.TENKENKAISU,T.TENKENYOBI,T.HOSYUHOHO,T.PACK_FLG,T.KEIYAKUNAIYO,T.KEIYAKU_NO,T.BIKO,K.SET_NM,K.HYOJIJUN,M.KIKIHINMEI_NK AS MEISAI_NM " +
            "FROM MCM.MCM_UM_TANKA T JOIN MCM.MCM_UM_KIKIMEISAI M ON M.UM_KIKIMEISAI_ID=T.UM_KIKIMEISAI_ID JOIN MCM.MCM_UM_KIKIKOSEI K ON K.UM_KIKIKOSEI_ID=M.UM_KIKIKOSEI_ID " +
            "LEFT JOIN MCM.MCM_MA_ATSUKAIKIKI A ON A.ATSUKAIKIKI_ID=M.ATSUKAIKIKI_ID " +
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
                row.setSeizomakerNk(rs.getString("SEIZOMAKER_NK"));row.setKikikatashiki(rs.getString("KIKIKATASHIKI"));
                row.setSosuNm(rs.getBigDecimal("SOSU_NM")==null?"":rs.getBigDecimal("SOSU_NM").stripTrailingZeros().toPlainString());
                row.setCreatedDt(displayDate(rs.getObject("CREATED_DT",LocalDateTime.class)));row.setCreatedBy(rs.getString("CREATED_BY"));row.setLastupdateDt(displayDate(rs.getObject("LASTUPDATE_DT",LocalDateTime.class)));row.setLastupdateBy(rs.getString("LASTUPDATE_BY"));
                row.setUmTankaId(rs.getBigDecimal("UM_TANKA_ID"));
                row.setUmKikoseiId(rs.getBigDecimal("UM_KIKIKOSEI_ID"));
                row.setSetNm(rs.getString("SET_NM"));
                row.setMeisaiNm(rs.getString("MEISAI_NM"));
                row.setSuryoNm(rs.getBigDecimal("SURYO_NM")==null?"0":rs.getBigDecimal("SURYO_NM").stripTrailingZeros().toPlainString());
                row.setKikikoseiNk(rs.getString("KIKIKOSEI_NK"));
                row.setSikiriKin(rs.getBigDecimal("SIKIRI_KIN"));
                row.setTorihosyujikanId(rs.getBigDecimal("TORIHOSYUJIKAN_ID"));
                row.setDaifukuhosyujikanId(rs.getBigDecimal("DAIFUKUHOSYUJIKAN_ID"));
                row.setTenkenkaisu(rs.getBigDecimal("TENKENKAISU"));
                row.setTenkenyobi(rs.getString("TENKENYOBI"));
                row.setHosyuhoho(rs.getString("HOSYUHOHO"));
                row.setPackFlg(rs.getBoolean("PACK_FLG"));
                row.setKeiyakunaiyo(rs.getString("KEIYAKUNAIYO"));
                row.setKeiyakuNo(rs.getString("KEIYAKU_NO"));
                row.setBiko(rs.getString("BIKO"));
                row.setHyojunKin(rs.getBigDecimal("HYOJUN_KIN"));
                row.setHyojunshokeiKin(rs.getBigDecimal("HYOJUNSHOKEI_KIN"));
                row.setSikirishokeiKin(rs.getBigDecimal("SIKIRISHOKEI_KIN"));
                row.setControllerFlg(rs.getInt("CONTROLLER_FLG")==1?"1":"0");
                return row;
            });
    }

    // ===================================================================
    // 見積条件保存 MCM_UM_BRAND と MCM_UM_KIHON_BRAND
    // 【変換元】UpdateButtonTabNaiyo → MyBase.UpdateAll(MCM_UM_BRAND)
    //
    // HOSHU_KIN = HARDHOSYU_KIN(MCM_UM_BRAND) + HOSEISOFTHOSHU_KIN(MCM_UM_KIHON_BRAND) - CHOSEI_KIN
    // 存在するブランド期間だけを更新する。集計値は共通集計サービスが計算する。
    // ===================================================================

    public void mergeBrand(Mcm2005uForm f,String user) {
        int count=jdbc.update("UPDATE MCM.MCM_UM_BRAND SET MITSUMORICHUKI=?,CHOSEI_KIN=?,BIKO=?,LASTUPDATE_DT=GETDATE(),LASTUPDATE_BY=? WHERE UM_KIHON_BRAND_ID=? AND UM_MITSUMORI_ID=?",f.getMitsumorichuki(),f.getChoseiKin(),f.getBiko(),user,f.getUmKihonBrandId(),f.getUmMitsumoriId());
        if(count!=1)throw new IllegalStateException("ブランド期間の情報を確認できません。再読み込みしてください。");
        jdbc.update("UPDATE MCM.MCM_UM_KIHON_BRAND SET DREMOS_FLG=?,REMOTE_FLG=?,LASTUPDATE_DT=GETDATE(),LASTUPDATE_BY=? WHERE UM_KIHON_BRAND_ID=?",f.getDremosFlg()?1:0,f.getRemoteFlg()?1:0,user,f.getUmKihonBrandId());
    }

    // ===================================================================
    // 機器単価更新 (MCM_UM_TANKA)
    // 【変換元】UpdateButtonTabNaiyo → MyBase.UpdateAll(MCM_UM_TANKA)
    //
    // 小計は数量×単価の計算値として表示し、テーブルの列として更新しない。
    // ===================================================================

    public void updateTanka(Mcm2005uForm.TankaRowForm row,String user) {
        // VB同様、参照単価と数量は変更せず、保守・点検・契約条件のみ保存する。
        if(jdbc.update("UPDATE MCM.MCM_UM_TANKA SET DAIFUKUHOSYUJIKAN_ID=?,TENKENKAISU=?,TENKENYOBI=?,HOSYUHOHO=?,PACK_FLG=?,KEIYAKUNAIYO=?,KEIYAKU_NO=?,BIKO=?,LASTUPDATE_DT=GETDATE(),LASTUPDATE_BY=? WHERE UM_TANKA_ID=?",row.getDaifukuhosyujikanId(),row.getTenkenkaisu(),row.getTenkenyobi(),row.getHosyuhoho(),row.getPackFlg()?1:0,row.getKeiyakunaiyo(),row.getKeiyakuNo(),row.getBiko(),user,row.getUmTankaId())!=1)throw new IllegalStateException("単価明細が変更されています。再読み込みしてください。");
    }
    public void fillVersion(Mcm2005uForm f){
        var rows=jdbc.queryForList("SELECT H.LASTUPDATE_DT,B.DREMOS_FLG,B.REMOTE_FLG FROM MCM.MCM_UM_KIHON_BRAND B JOIN MCM.MCM_UM_KIHON_MITSUMORI H ON H.UM_KIHON_MITSUMORI_ID=B.UM_KIHON_MITSUMORI_ID JOIN MCM.MCM_UM_MITSUMORI P ON P.UM_KIHON_MITSUMORI_ID=H.UM_KIHON_MITSUMORI_ID WHERE B.UM_KIHON_BRAND_ID=? AND P.UM_MITSUMORI_ID=?",f.getUmKihonBrandId(),f.getUmMitsumoriId());
        if(rows.size()!=1)throw new IllegalStateException("同じ見積のブランドと期間を選択してください。");
        var r=rows.get(0);f.setVersion(r.get("LASTUPDATE_DT")==null?"":((java.sql.Timestamp)r.get("LASTUPDATE_DT")).toLocalDateTime().toString());f.setDremosFlg("1".equals(String.valueOf(r.get("DREMOS_FLG")))||r.get("DREMOS_FLG") instanceof Number n&&n.intValue()==1);f.setRemoteFlg(r.get("REMOTE_FLG") instanceof Number n&&n.intValue()==1);
    }
    public List<java.util.Map<String,Object>> hours(){var rows=jdbc.queryForList("SELECT * FROM MCM.MCM_MA_TORIHOSYUJIKAN WHERE EXISTS (SELECT 1 FROM MCM.MCM_MA_TORIHIKISAKI R WHERE R.TORIHIKISAKI_ID=MCM_MA_TORIHOSYUJIKAN.TORIHIKISAKI_ID AND R.DAIFUKU_FLG=1) ORDER BY HYOJIJUN,TORIHOSYUJIKAN_ID");for(var row:rows)row.put("label",hourLabel(row));return rows;}
    public String hourName(BigDecimal id){if(id==null||id.signum()==0)return "";var rows=jdbc.queryForList("SELECT * FROM MCM.MCM_MA_TORIHOSYUJIKAN WHERE TORIHOSYUJIKAN_ID=?",id);return rows.isEmpty()?"設定を確認してください":hourLabel(rows.get(0));}
    private static String hourLabel(java.util.Map<String,Object> r){String label=numberText(r.get("HOSYUJIKAN_DT"))+"H "+day(r.get("KAISIYOBI"))+"～"+day(r.get("SYURYOYOBI"));if(r.get("KAISIJIKAN_DT")!=null)label+=" "+numberText(r.get("KAISIJIKAN_DT"))+"～"+numberText(r.get("SYURYOJIKAN_DT"));if(r.get("REIGAIYOBI")!=null&&!numberText(r.get("REIGAIYOBI")).equals("0"))label+=" ("+day(r.get("REIGAIYOBI"))+" "+numberText(r.get("REIGAIKAISIJIKAN_DT"))+"～"+numberText(r.get("REIGAISYURYOJIKAN_DT"))+")";return label;}
    private static String numberText(Object v){return v==null?"":v instanceof BigDecimal n?n.stripTrailingZeros().toPlainString():v.toString();}
    private static String day(Object v){return java.util.Map.of("1","月","2","火","3","水","4","木","5","金","6","土","7","日").getOrDefault(numberText(v),"");}
    public List<java.util.Map<String,Object>> tenpoChoices(){
        var rows=jdbc.queryForList("SELECT TENPO_ID,MEISHO1_NK,MEISHO2_NK,MEISHO3_NK,MEISHO4_NK FROM MCM.MCM_MA_TENPO ORDER BY MEISHO1_NK,MEISHO4_NK,TENPO_ID");
        for(var row:rows)row.put("label",java.util.List.of("MEISHO1_NK","MEISHO2_NK","MEISHO3_NK","MEISHO4_NK").stream().map(k->java.util.Objects.toString(row.get(k),"")).filter(v->!v.isBlank()).collect(java.util.stream.Collectors.joining(" ")));return rows;
    }
    public void updateHeader(Mcm2005uForm f,String user){
        var header=f.getHeader();
        var columns=new java.util.ArrayList<String>();var args=new java.util.ArrayList<Object>();
        for(String prefix:java.util.List.of("IRAI","SOFU")){
            BigDecimal id=prefix.equals("IRAI")?header.getIraitenpoId():header.getSofutenpoId();
            var rows=id==null?java.util.List.<java.util.Map<String,Object>>of():jdbc.queryForList("SELECT * FROM MCM.MCM_MA_TENPO WHERE TENPO_ID=?",id);
            if(id!=null&&rows.isEmpty())throw new IllegalStateException("依頼元・送付先の店舗を選択し直してください。");
            columns.add(prefix+"TENPO_ID=?");args.add(id);
            for(String name:java.util.List.of("MEISHO1_NK","MEISHO2_NK","MEISHO3_NK","MEISHO4_NK","TENPORYAKU_NK")){columns.add(prefix+name+"=?");args.add(rows.isEmpty()?"":rows.get(0).get(name));}
        }
        columns.add("IRAITANTO_NK=?");args.add(header.getIraitantoNk());columns.add("SOFUTANTO_NK=?");args.add(header.getSofutantoNk());args.add(user);args.add(f.getUmKihonMitsumoriId());
        if(jdbc.update("UPDATE MCM.MCM_UM_KIHON_MITSUMORI SET "+String.join(",",columns)+",LASTUPDATE_DT=GETDATE(),LASTUPDATE_BY=? WHERE UM_KIHON_MITSUMORI_ID=?",args.toArray())!=1)throw new IllegalStateException("見積情報が見つかりません。");
        if(jdbc.update("UPDATE MCM.MCM_UM_MITSUMORI SET BIKO=?,LASTUPDATE_DT=GETDATE(),LASTUPDATE_BY=? WHERE UM_MITSUMORI_ID=? AND UM_KIHON_MITSUMORI_ID=?",f.getPeriodBiko(),user,f.getUmMitsumoriId(),f.getUmKihonMitsumoriId())!=1)throw new IllegalStateException("契約期間が見つかりません。");
    }
    private static String displayDate(LocalDateTime value){return value==null?"":value.format(DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss"));}
    public void fillLayout(Mcm2005uForm f){
        var p=jdbc.queryForMap("SELECT BIKO,CREATED_DT,CREATED_BY,LASTUPDATE_DT,LASTUPDATE_BY FROM MCM.MCM_UM_MITSUMORI WHERE UM_MITSUMORI_ID=? AND UM_KIHON_MITSUMORI_ID=?",f.getUmMitsumoriId(),f.getUmKihonMitsumoriId());f.setPeriodBiko((String)p.get("BIKO"));f.setPeriodMeta(displayMeta(p));
        var rows=jdbc.queryForList("SELECT CREATED_DT,CREATED_BY,LASTUPDATE_DT,LASTUPDATE_BY FROM MCM.MCM_UM_BRAND WHERE UM_MITSUMORI_ID=? AND UM_KIHON_BRAND_ID=?",f.getUmMitsumoriId(),f.getUmKihonBrandId());f.setBrandPeriodMeta(displayMeta(rows.isEmpty()?new java.util.LinkedHashMap<>():rows.get(0)));
    }
    private java.util.Map<String,Object> displayMeta(java.util.Map<String,Object> row){for(String key:java.util.List.of("CREATED_DT","LASTUPDATE_DT")){Object v=row.get(key);row.put(key,v instanceof java.sql.Timestamp t?displayDate(t.toLocalDateTime()):v==null?"":v.toString());}return row;}
}
