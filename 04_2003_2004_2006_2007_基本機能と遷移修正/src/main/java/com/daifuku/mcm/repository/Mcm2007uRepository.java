package com.daifuku.mcm.repository;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Repository;
import com.daifuku.mcm.form.Mcm2007uForm.*;
import static com.daifuku.mcm.common.CustomerScreenSupport.format;

/** VB版Mcm2007uDataSet / Mcm2007u2DataSetのマスタと見積の対応で選択候補を取得する。 */
@Repository
public class Mcm2007uRepository {
    @Autowired private JdbcTemplate jdbc;
    private String date(java.sql.Date value) { return value == null ? "" : format(value.toLocalDate()); }
    private String marks(List<BigDecimal> ids) { return String.join(",", java.util.Collections.nCopies(ids.size(), "?")); }
    private Object[] params(List<BigDecimal> ids, BigDecimal plant) { var result=new ArrayList<Object>(ids);result.add(plant);return result.toArray(); }

    public BigDecimal findInitialMitsumoriId(BigDecimal id) {
        var list=jdbc.queryForList("SELECT TOP (1) UM_MITSUMORI_ID FROM MCM.MCM_UM_MITSUMORI WHERE UM_KIHON_MITSUMORI_ID=? AND KAISI_DT IS NOT NULL AND (SYURYO_DT IS NULL OR SYURYO_DT>=CAST(GETDATE() AS date)) ORDER BY CASE WHEN KAISI_DT<=CAST(GETDATE() AS date) THEN 0 ELSE 1 END,KAISI_DT",BigDecimal.class,id);
        return list.isEmpty()?null:list.get(0);
    }
    public List<MitsumoriRowForm> findMitsumoriList(BigDecimal plant) {
        return jdbc.query("SELECT G.UM_MITSUMORI_ID,COALESCE(A.UM_MITSUMORI_NO,'----') AS UM_MITSUMORI_NO,G.KAISI_DT,G.SYURYO_DT,A.JOTAI,A.MITSUMORI_DT,G.MITSUMORI_GKIN FROM MCM.MCM_UM_MITSUMORI G JOIN MCM.MCM_UM_KIHON_MITSUMORI A ON A.UM_KIHON_MITSUMORI_ID=G.UM_KIHON_MITSUMORI_ID WHERE A.PLANT_ID=? AND A.SYOUNIN_JOTAI='3' AND A.JOTAI<>'3' ORDER BY A.UM_MITSUMORI_NO,G.KAISI_DT,G.UM_MITSUMORI_ID",new Object[]{plant},(rs,n)->{
            var r=new MitsumoriRowForm();r.setDisplay(display(rs));r.setUmMitsumoriId(rs.getBigDecimal("UM_MITSUMORI_ID"));r.setUmMitsumoriNo(rs.getString("UM_MITSUMORI_NO"));r.setKaisiDt(date(rs.getDate("KAISI_DT")));r.setSyuryoDt(date(rs.getDate("SYURYO_DT")));r.setJotai(rs.getString("JOTAI"));return r;
        });
    }
    public List<KihonBrandRowForm> findKihonBrandList(List<BigDecimal> ids) {
        if(ids.isEmpty())return new ArrayList<>();
        return jdbc.query("SELECT L.HOSHU_KIN,L.HARDHOSYU_KIN,L.CHOSEI_KIN,B.HOSEISOFTHOSHU_KIN,B.DREMOS_FLG,B.REMOTE_FLG,B.SOFT_FLG,B.SOFTHOSYUHOHO,B.BIKO,B.UM_KIHON_BRAND_ID,L.UM_MITSUMORI_ID,B.BRANDKOSEI_ID,B.BRAND_NK,B.BRANDSYOSAI_NK,B.KEIYAKUJIKANTAI,A.HOSYUHOHO FROM MCM.MCM_UM_BRAND L JOIN MCM.MCM_UM_KIHON_BRAND B ON B.UM_KIHON_BRAND_ID=L.UM_KIHON_BRAND_ID JOIN MCM.MCM_UM_KIHON_MITSUMORI A ON A.UM_KIHON_MITSUMORI_ID=B.UM_KIHON_MITSUMORI_ID WHERE L.UM_MITSUMORI_ID IN ("+marks(ids)+") ORDER BY L.UM_MITSUMORI_ID,(SELECT D.HYOJIJUN FROM MCM.MCM_MA_BRAND_KOSEI D WHERE D.BRANDKOSEI_ID=B.BRANDKOSEI_ID),B.UM_KIHON_BRAND_ID",ids.toArray(),(rs,n)->{
            var r=new KihonBrandRowForm();r.setDisplay(display(rs));r.setHoshuKin(rs.getBigDecimal("HOSHU_KIN"));r.setUmKihonBrandId(rs.getBigDecimal("UM_KIHON_BRAND_ID"));r.setUmMitsumoriId(rs.getBigDecimal("UM_MITSUMORI_ID"));r.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));r.setBrandNk(rs.getString("BRAND_NK"));r.setBrandsyosaiNk(rs.getString("BRANDSYOSAI_NK"));r.setKeiyakujikantai(rs.getString("KEIYAKUJIKANTAI"));r.setHosyuhoho(rs.getString("HOSYUHOHO"));return r;
        });
    }
    /** VB Step1の機器明細。選定元見積の保存済み単価・条件を表示するだけで、Step2の選択候補には使わない。 */
    public List<java.util.Map<String,String>> findPreviewRows(List<BigDecimal> ids) {
        if(ids.isEmpty())return new ArrayList<>();
        return jdbc.query("""
            SELECT T.UM_MITSUMORI_ID,B.BRAND_NK,B.BRANDSYOSAI_NK,C.TEHAISEIBAN,
                   U.SEIZOMAKER_NK,U.KIKIHINMEI_NK,U.KIKIKATASHIKI,U.SURYO_NM,
                   T.HYOJUN_KIN,T.PACK_FLG,T.KEIYAKUNAIYO,T.KEIYAKU_NO,
                   T.DAIFUKUHOSYUJIKAN_ID,T.TENKENKAISU,T.TENKENYOBI,T.HOSYUHOHO,T.SERVICEKEITAI,T.BIKO
            FROM MCM.MCM_UM_KIKIMEISAI U
            JOIN MCM.MCM_UM_KIKIKOSEI C ON C.UM_KIKIKOSEI_ID=U.UM_KIKIKOSEI_ID
            JOIN MCM.MCM_UM_KIHON_BRAND B ON B.UM_KIHON_BRAND_ID=C.UM_KIHON_BRAND_ID
            JOIN MCM.MCM_UM_TANKA T ON T.UM_KIKIMEISAI_ID=U.UM_KIKIMEISAI_ID
            JOIN MCM.MCM_UM_MITSUMORI G ON G.UM_MITSUMORI_ID=T.UM_MITSUMORI_ID AND G.UM_KIHON_MITSUMORI_ID=B.UM_KIHON_MITSUMORI_ID
            JOIN MCM.MCM_MA_BRAND_KOSEI D ON D.BRANDKOSEI_ID=B.BRANDKOSEI_ID
            JOIN MCM.MCM_MA_KIKIKOSEI K ON K.KIKIKOSEI_ID=C.KIKIKOSEI_ID
            JOIN MCM.MCM_MA_KIKIMEISAI M ON M.KIKIMEISAI_ID=U.KIKIMEISAI_ID
            WHERE T.UM_MITSUMORI_ID IN (
            """+marks(ids)+") ORDER BY D.HYOJIJUN,K.HYOJIJUN,M.HYOJIJUN,U.UM_KIKIMEISAI_ID",ids.toArray(),(rs,n)->display(rs));
    }
    private java.util.Map<String,String> display(java.sql.ResultSet rs)throws java.sql.SQLException {
        var result=new java.util.LinkedHashMap<String,String>();
        for(int i=1;i<=rs.getMetaData().getColumnCount();i++){
            String key=rs.getMetaData().getColumnLabel(i);Object value=rs.getObject(i);String text="";
            if(value instanceof Number number)text=key.endsWith("_KIN")||key.endsWith("_GKIN")?new java.text.DecimalFormat("#,##0").format(number):new BigDecimal(number.toString()).stripTrailingZeros().toPlainString();
            else if(key.endsWith("_DT")&&value!=null)text=date(rs.getDate(i));
            else if(value!=null)text=value.toString();
            result.put(key,text);
        }
        return result;
    }
    /** VB Step1の重複候補通知。VB同様、該当しても選定画面への遷移自体は止めない。 */
    public boolean hasDuplicateIndividuals(List<BigDecimal> ids) {
        if(ids.isEmpty())return false;
        return jdbc.queryForObject("""
            SELECT COUNT(*) FROM (
              SELECT U.KIKIMEISAI_ID,T.KOTAIKANRI_ID
              FROM MCM.MCM_UM_BRAND L
              JOIN MCM.MCM_UM_KIHON_BRAND B ON B.UM_KIHON_BRAND_ID=L.UM_KIHON_BRAND_ID
              JOIN MCM.MCM_UM_KIKIKOSEI C ON C.UM_KIHON_BRAND_ID=B.UM_KIHON_BRAND_ID
              JOIN MCM.MCM_UM_KIKIMEISAI U ON U.UM_KIKIKOSEI_ID=C.UM_KIKIKOSEI_ID
              JOIN MCM.MCM_UM_KOTAIMEISAI T ON T.UM_KIKIMEISAI_ID=U.UM_KIKIMEISAI_ID
              WHERE L.UM_MITSUMORI_ID IN (
            """+marks(ids)+") GROUP BY U.KIKIMEISAI_ID,T.KOTAIKANRI_ID HAVING COUNT(*)>1) D",Integer.class,ids.toArray())>0;
    }
    public List<KoseiRowForm> findKoseiList(List<BigDecimal> ids,BigDecimal plant) {
        if(ids.isEmpty())return new ArrayList<>();
        String sql="SELECT E.KIKIKOSEI_ID,D.BRANDKOSEI_ID,E.KIKIKOSEI_NK,COALESCE(S.SET_NM,E.SET_NM) AS SET_NM,E.TANI,E.BIKO,E.TEHAISEIBAN,E.CONTROLLER_FLG,E.HYOJIJUN,S.UM_KIHON_BRAND_ID,S.HOSYUHOHO,CASE WHEN S.KIKIKOSEI_ID IS NULL THEN 0 ELSE 1 END AS CHECK_FLG FROM MCM.MCM_MA_KIKIKOSEI E JOIN MCM.MCM_MA_BRAND_KOSEI D ON D.PLANT_ID=E.PLANT_ID LEFT JOIN (SELECT B.BRANDKOSEI_ID,C.KIKIKOSEI_ID,MIN(B.UM_KIHON_BRAND_ID) AS UM_KIHON_BRAND_ID,MAX(C.HOSYUHOHO) AS HOSYUHOHO,MAX(C.SET_NM) AS SET_NM FROM MCM.MCM_UM_KIHON_BRAND B JOIN MCM.MCM_UM_BRAND L ON L.UM_KIHON_BRAND_ID=B.UM_KIHON_BRAND_ID JOIN MCM.MCM_UM_KIKIKOSEI C ON C.UM_KIHON_BRAND_ID=B.UM_KIHON_BRAND_ID WHERE L.UM_MITSUMORI_ID IN ("+marks(ids)+") GROUP BY B.BRANDKOSEI_ID,C.KIKIKOSEI_ID) S ON S.BRANDKOSEI_ID=D.BRANDKOSEI_ID AND S.KIKIKOSEI_ID=E.KIKIKOSEI_ID WHERE E.PLANT_ID=? ORDER BY D.HYOJIJUN,E.HYOJIJUN";
        return jdbc.query(sql,params(ids,plant),(rs,n)->{
            var r=new KoseiRowForm();r.setDisplay(display(rs));r.setQuoted(rs.getInt("CHECK_FLG")==1);r.setKikikoseiId(rs.getBigDecimal("KIKIKOSEI_ID"));r.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));r.setUmKihonBrandId(rs.getBigDecimal("UM_KIHON_BRAND_ID"));r.setKikikoseiNk(rs.getString("KIKIKOSEI_NK"));r.setSetNm(rs.getString("SET_NM"));r.setTani(rs.getString("TANI"));r.setTehaiseiban(rs.getString("TEHAISEIBAN"));r.setControllerFlg(rs.getString("CONTROLLER_FLG"));r.setHosyuhoho(rs.getString("HOSYUHOHO"));r.setHyojijun(rs.getInt("HYOJIJUN"));r.setCheckFlg(rs.getInt("CHECK_FLG")==1);return r;
        });
    }
    public List<MeisaiRowForm> findMeisaiList(List<BigDecimal> ids,BigDecimal plant) {
        if(ids.isEmpty())return new ArrayList<>();
        String sql="SELECT M.KIKIMEISAI_ID,M.KIKIKOSEI_ID,D.BRANDKOSEI_ID,J.SEIZOMAKER_ID,J.SEIZOMAKER_NK,M.KIKIHINMEI_NK,M.KIKIKATASHIKI,M.SURYO_NM,M.BIKO,Q.KOTAIKANRI_FLG,E.HYOJIJUN AS MAE_HYOJIJUN,M.HYOJIJUN,CASE WHEN S.KIKIMEISAI_ID IS NULL THEN 0 ELSE 1 END AS CHECK_FLG FROM MCM.MCM_MA_KIKIMEISAI M JOIN MCM.MCM_MA_KIKIKOSEI E ON E.KIKIKOSEI_ID=M.KIKIKOSEI_ID JOIN MCM.MCM_MA_BRAND_KOSEI D ON D.PLANT_ID=E.PLANT_ID LEFT JOIN MCM.MCM_MA_ATSUKAIKIKI H ON H.ATSUKAIKIKI_ID=M.ATSUKAIKIKI_ID LEFT JOIN MCM.MCM_MA_SEIZOMAKER J ON J.SEIZOMAKER_ID=H.SEIZOMAKER_ID LEFT JOIN MCM.MCM_MA_KIKIBUNRUI Q ON Q.KIKIBUNRUI_ID=H.KIKIBUNRUI_ID LEFT JOIN (SELECT DISTINCT B.BRANDKOSEI_ID,U.KIKIMEISAI_ID FROM MCM.MCM_UM_KIHON_BRAND B JOIN MCM.MCM_UM_BRAND L ON L.UM_KIHON_BRAND_ID=B.UM_KIHON_BRAND_ID JOIN MCM.MCM_UM_KIKIKOSEI C ON C.UM_KIHON_BRAND_ID=B.UM_KIHON_BRAND_ID JOIN MCM.MCM_UM_KIKIMEISAI U ON U.UM_KIKIKOSEI_ID=C.UM_KIKIKOSEI_ID WHERE L.UM_MITSUMORI_ID IN ("+marks(ids)+")) S ON S.BRANDKOSEI_ID=D.BRANDKOSEI_ID AND S.KIKIMEISAI_ID=M.KIKIMEISAI_ID WHERE E.PLANT_ID=? ORDER BY D.HYOJIJUN,E.HYOJIJUN,M.HYOJIJUN";
        return jdbc.query(sql,params(ids,plant),(rs,n)->{
            var r=new MeisaiRowForm();r.setDisplay(display(rs));r.setQuoted(rs.getInt("CHECK_FLG")==1);r.setKikimeisaiId(rs.getBigDecimal("KIKIMEISAI_ID"));r.setKikikoseiId(rs.getBigDecimal("KIKIKOSEI_ID"));r.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));r.setSeizomakerId(rs.getBigDecimal("SEIZOMAKER_ID"));r.setSeizomankerNk(rs.getString("SEIZOMAKER_NK"));r.setKikihinmeiNk(rs.getString("KIKIHINMEI_NK"));r.setKikikatashiki(rs.getString("KIKIKATASHIKI"));r.setSuryoNm(rs.getString("SURYO_NM"));r.setMaeHyojijun(rs.getInt("MAE_HYOJIJUN"));r.setHyojijun(rs.getInt("HYOJIJUN"));r.setCheckFlg(rs.getInt("CHECK_FLG")==1);return r;
        });
    }
    public List<KotaiRowForm> findKotaiList(List<BigDecimal> ids,BigDecimal plant) {
        if(ids.isEmpty())return new ArrayList<>();
        String sql="SELECT V.MAG_KOTAIKANRI_ID AS KOTAIKANRI_ID,V.MAF_KIKIMEISAI_ID AS KIKIMEISAI_ID,V.MAE_KIKIKOSEI_ID AS KIKIKOSEI_ID,V.MAD_BRANDKOSEI_ID AS BRANDKOSEI_ID,V.MAG_KOTAI_NK AS KOTAI_NK,V.MAG_SERIAL_NO AS SERIAL_NO,V.MAG_ITIJINONYU_DT AS ITIJINONYU_DT,V.MAG_SETCHIBASYO AS SETCHIBASYO,V.MAF_KIKIHINMEI_NK AS KIKIHINMEI_NK,V.MAF_KIKIKATASHIKI AS KIKIKATASHIKI,V.MAF_SURYO_NM AS SURYO_NM,V.MAK_TORIHIKISAKI_NK AS TORIHIKISAKI_NK,V.MAG_KEIYAKUKIGEN_DT AS MAG_ENCHOKEIYAKUKIGEN_DT,CASE WHEN S.KOTAIKANRI_ID IS NULL THEN 0 ELSE 1 END AS CHECK_FLG FROM MCM.MCM_MA_KOTAIMEISAI_V V LEFT JOIN (SELECT DISTINCT B.BRANDKOSEI_ID,U.KIKIMEISAI_ID,T.KOTAIKANRI_ID FROM MCM.MCM_UM_KIHON_BRAND B JOIN MCM.MCM_UM_BRAND L ON L.UM_KIHON_BRAND_ID=B.UM_KIHON_BRAND_ID JOIN MCM.MCM_UM_KIKIKOSEI C ON C.UM_KIHON_BRAND_ID=B.UM_KIHON_BRAND_ID JOIN MCM.MCM_UM_KIKIMEISAI U ON U.UM_KIKIKOSEI_ID=C.UM_KIKIKOSEI_ID JOIN MCM.MCM_UM_KOTAIMEISAI T ON T.UM_KIKIMEISAI_ID=U.UM_KIKIMEISAI_ID WHERE L.UM_MITSUMORI_ID IN ("+marks(ids)+")) S ON S.BRANDKOSEI_ID=V.MAD_BRANDKOSEI_ID AND S.KIKIMEISAI_ID=V.MAF_KIKIMEISAI_ID AND S.KOTAIKANRI_ID=V.MAG_KOTAIKANRI_ID WHERE V.MAE_PLANT_ID=? ORDER BY V.MAD_HYOJIJUN,V.MAE_HYOJIJUN,V.MAF_HYOJIJUN,V.MAG_HYOJIJUN";
        return jdbc.query(sql,params(ids,plant),(rs,n)->{
            var r=new KotaiRowForm();r.setDisplay(display(rs));r.setQuoted(rs.getInt("CHECK_FLG")==1);r.setKotaikanriId(rs.getBigDecimal("KOTAIKANRI_ID"));r.setKikimeisaiId(rs.getBigDecimal("KIKIMEISAI_ID"));r.setKikikoseiId(rs.getBigDecimal("KIKIKOSEI_ID"));r.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));r.setKotaiNk(rs.getString("KOTAI_NK"));r.setSerialNo(rs.getString("SERIAL_NO"));r.setItizinonnyuDt(date(rs.getDate("ITIJINONYU_DT")));r.setSetchibasyo(rs.getString("SETCHIBASYO"));r.setCheckFlg(rs.getInt("CHECK_FLG")==1);return r;
        });
    }
    public List<java.util.Map<String,String>> findPlantBrands(BigDecimal plant) {
        return jdbc.query("SELECT BRANDKOSEI_ID,BRANDSYOSAI_NK FROM MCM.MCM_MA_BRAND_KOSEI WHERE PLANT_ID=? ORDER BY HYOJIJUN,BRANDKOSEI_ID",new Object[]{plant},(rs,n)->display(rs));
    }
    public List<java.util.Map<String,String>> findSelectionSources(List<BigDecimal> ids) {
        if(ids.isEmpty())return new ArrayList<>();
        return jdbc.query("""
            SELECT L.UM_MITSUMORI_ID,A.UM_MITSUMORI_NO,B.BRANDKOSEI_ID,C.KIKIKOSEI_ID,U.KIKIMEISAI_ID,T.KOTAIKANRI_ID
            FROM MCM.MCM_UM_BRAND L
            JOIN MCM.MCM_UM_KIHON_BRAND B ON B.UM_KIHON_BRAND_ID=L.UM_KIHON_BRAND_ID
            JOIN MCM.MCM_UM_KIHON_MITSUMORI A ON A.UM_KIHON_MITSUMORI_ID=B.UM_KIHON_MITSUMORI_ID
            LEFT JOIN MCM.MCM_UM_KIKIKOSEI C ON C.UM_KIHON_BRAND_ID=B.UM_KIHON_BRAND_ID
            LEFT JOIN MCM.MCM_UM_KIKIMEISAI U ON U.UM_KIKIKOSEI_ID=C.UM_KIKIKOSEI_ID
            LEFT JOIN MCM.MCM_UM_KOTAIMEISAI T ON T.UM_KIKIMEISAI_ID=U.UM_KIKIMEISAI_ID
            WHERE L.UM_MITSUMORI_ID IN (
            """+marks(ids)+") ORDER BY A.UM_MITSUMORI_NO,L.UM_MITSUMORI_ID",ids.toArray(),(rs,n)->display(rs));
    }
    public List<MeisaiRowForm> findEstimateTerms(BigDecimal brand,BigDecimal estimate) {
        return jdbc.query("SELECT U.KIKIMEISAI_ID,T.KEIYAKUNAIYO,T.KEIYAKU_NO,T.SERVICEKEITAI,T.TORIHOSYUJIKAN_ID,T.DAIFUKUHOSYUJIKAN_ID,T.TENKENKAISU,T.TENKENYOBI,T.HOSYUHOHO FROM MCM.MCM_UM_KIKIMEISAI U JOIN MCM.MCM_UM_KIKIKOSEI C ON C.UM_KIKIKOSEI_ID=U.UM_KIKIKOSEI_ID JOIN MCM.MCM_UM_TANKA T ON T.UM_KIKIMEISAI_ID=U.UM_KIKIMEISAI_ID WHERE C.UM_KIHON_BRAND_ID=? AND T.UM_MITSUMORI_ID=?",new org.springframework.jdbc.core.BeanPropertyRowMapper<>(MeisaiRowForm.class),brand,estimate);
    }
}
