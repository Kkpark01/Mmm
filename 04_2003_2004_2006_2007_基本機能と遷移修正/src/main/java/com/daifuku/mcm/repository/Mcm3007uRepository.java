package com.daifuku.mcm.repository;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.util.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Repository;
import com.daifuku.mcm.form.Mcm3007uForm.*;
/** MCM3007UのTableAdapter XMLに基づく検索・選択行の更新。 */
@Repository public class Mcm3007uRepository {
 @Autowired private JdbcTemplate jdbc;
 private static String marks(int n){return String.join(",",Collections.nCopies(n,"?"));}
 private static List<String> states(List<String> values){var r=new ArrayList<String>();for(String s:values){r.add(s);r.add(switch(s){case "1"->"見積";case "2"->"契約";case "3"->"破棄";case "4"->"解約";default->s;});}return r;}
 private static void filters(StringBuilder sql,List<Object> args,String alias,String cd,String nk,String sid,String pnk,boolean prefix){
  if(cd!=null&&!cd.isBlank()){sql.append(" AND "+alias+".NONYUSAKI_CD LIKE ?");args.add((prefix?"":"%")+cd.trim()+"%");}
  if(nk!=null&&!nk.isBlank()){sql.append(" AND UPPER(CONCAT("+alias+".NONYUSAKI_NK,' ',"+alias+".KYUNONYUSAKI_NK,' ',"+alias+".NONYUSAKIKOJO_NK,' ',"+alias+".NONYUSAKIKANA_KN,' ',"+alias+".NONYUSAKIEIMEI_EN)) LIKE ?");args.add("%"+nk.trim().toUpperCase(Locale.ROOT)+"%");}
  if(sid!=null&&!sid.isBlank()){sql.append(" AND "+alias+".SUPPORT_ID LIKE ?");args.add((prefix?"":"%")+sid.trim()+"%");}
  if(pnk!=null&&!pnk.isBlank()){sql.append(" AND UPPER("+alias+".PLANT_NK) LIKE ?");args.add("%"+pnk.trim().toUpperCase(Locale.ROOT)+"%");}
 }
 public static Map<String,String> display(Map<String,Object> row){var d=new LinkedHashMap<String,String>();row.forEach((k,v)->{String s=v==null?"":v.toString();if(v instanceof BigDecimal b)s=b.stripTrailingZeros().toPlainString();else if(v instanceof java.util.Date&&k.endsWith("_DT"))s=s.substring(0,10).replace('-','/');d.put(k.toUpperCase(Locale.ROOT),s);});return d;}
 public static BigDecimal id(Map<String,String>d,String key){String v=d.get(key);return v==null||v.isBlank()?null:new BigDecimal(v);}
 private static String state(String v){return switch(v==null?"":v){case "1"->"見積";case "2"->"契約";case "3"->"破棄";case "4"->"解約";default->v;};}
 public List<Grid1RowForm> searchNonyusaki(String cd,String nk,String sid,String pnk){return searchNonyusaki(cd,nk,sid,pnk,List.of("1","2","3","4"));}
 public List<Grid1RowForm> searchNonyusaki(String cd,String nk,String sid,String pnk,List<String> status){
  if(status.isEmpty())return List.of();var st=states(status);var args=new ArrayList<Object>(st);
  var sql=new StringBuilder("SELECT DISTINCT N.*,P.PLANT_ID,P.SUPPORT_ID,P.PLANT_NK,P.NONYU_DT,P.TEKKYO_DT,P.HOSYUSYUSOKU_DT FROM MCM.MCM_MA_NONYUSAKI N JOIN MCM.MCM_MA_PLANT P ON P.NONYUSAKI_ID=N.NONYUSAKI_ID JOIN MCM.MCM_2004_V V ON V.NONYUSAKI_ID=N.NONYUSAKI_ID AND V.PLANT_ID=P.PLANT_ID WHERE V.JOTAI IN ("+marks(st.size())+")");
  filters(sql,args,"V",cd,nk,sid,pnk,true);sql.append(" ORDER BY N.NONYUSAKI_ID,P.PLANT_ID");
  return jdbc.queryForList(sql.toString(),args.toArray()).stream().map(m->{var d=display(m);var r=new Grid1RowForm();r.setDisplay(d);r.setNonyusakiId(id(d,"NONYUSAKI_ID"));r.setNonyusakiCd(d.get("NONYUSAKI_CD"));r.setNonyusakiNk(d.get("NONYUSAKI_NK"));r.setSupportId(d.get("SUPPORT_ID"));r.setPlantNk(d.get("PLANT_NK"));return r;}).toList();
 }
 public List<Grid2RowForm> searchUva(String cd,String nk,String sid,String pnk,List<String> status){
  if(status.isEmpty())return List.of();var st=states(status);var args=new ArrayList<Object>(st);
  var sql=new StringBuilder("SELECT V.*,K.LASTUPDATE_DT AS CONTRACT_VERSION FROM MCM.MCM_2004_V V LEFT JOIN MCM.MCM_UK_KEIYAKU K ON K.UK_KEIYAKU_ID=V.UK_KEIYAKU_ID WHERE V.JOTAI IN ("+marks(st.size())+")");filters(sql,args,"V",cd,nk,sid,pnk,false);sql.append(" ORDER BY V.UM_MITSUMORI_NO,V.KEIYAKU_NO");
  return jdbc.queryForList(sql.toString(),args.toArray()).stream().map(m->{var d=display(m);d.put("JOTAI",state(d.get("JOTAI")));var r=new Grid2RowForm();r.setDisplay(d);r.setUkKeiyakuId(id(d,"UK_KEIYAKU_ID"));r.setUmKihonMitsumoriId(id(d,"UM_KIHON_MITSUMORI_ID"));r.setJotai(d.get("JOTAI"));r.setShoruiNo(d.get("UM_MITSUMORI_NO"));r.setNonyusakiNk(d.get("NONYUSAKI_NK"));r.setSupportId(d.get("SUPPORT_ID"));r.setPlantNk(d.get("PLANT_NK"));r.setKeiyakuNo(d.get("KEIYAKU_NO"));r.setKeiyakuDt(d.get("KEIYAKU_DT"));r.setKaiyakuDt(d.get("KAIYAKU_DT"));r.setShoninJotai(d.get("SYOUNIN_JOTAI"));return r;}).toList();
 }
 public List<Grid3RowForm> searchTka(String cd,String nk,String sid,String pnk,List<String> status){
  if(status.isEmpty())return List.of();var st=states(status);var args=new ArrayList<Object>(st);
  var sql=new StringBuilder("SELECT DISTINCT V.*,B.UM_KIHON_MITSUMORI_ID,K.LASTUPDATE_DT AS CONTRACT_VERSION FROM MCM.MCM_1003_V V JOIN MCM.MCM_TM_KIKIKOSEI C ON C.TM_IRAI_ID=V.TM_IRAI_ID JOIN MCM.MCM_TM_KIKIMEISAI D ON D.TM_KIKIKOSEI_ID=C.TM_KIKIKOSEI_ID JOIN MCM.MCM_TM_KOTAIMEISAI E ON E.TM_KIKIMEISAI_ID=D.TM_KIKIMEISAI_ID JOIN MCM.MCM_UM_KOTAIMEISAI UE ON UE.KOTAIKANRI_ID=E.KOTAIKANRI_ID JOIN MCM.MCM_UM_KIKIMEISAI UD ON UD.UM_KIKIMEISAI_ID=UE.UM_KIKIMEISAI_ID JOIN MCM.MCM_UM_KIKIKOSEI UC ON UC.UM_KIKIKOSEI_ID=UD.UM_KIKIKOSEI_ID JOIN MCM.MCM_UM_KIHON_BRAND B ON B.UM_KIHON_BRAND_ID=UC.UM_KIHON_BRAND_ID LEFT JOIN MCM.MCM_TK_KEIYAKU K ON K.TK_KEIYAKU_ID=V.TK_KEIYAKU_ID WHERE V.JOTAI IN ("+marks(st.size())+")");filters(sql,args,"V",cd,nk,sid,pnk,true);sql.append(" ORDER BY V.TM_IRAI_NO DESC,V.KEIYAKUJIKANTAI");
  return jdbc.queryForList(sql.toString(),args.toArray()).stream().map(m->{var d=display(m);d.put("JOTAI",state(d.get("JOTAI")));var r=new Grid3RowForm();r.setDisplay(d);r.setTkKeiyakuId(id(d,"TK_KEIYAKU_ID"));r.setTmKeiyakujikanId(id(d,"TM_KEIYAKUJIKAN_ID"));r.setJotai(d.get("JOTAI"));r.setIraino(d.get("TM_IRAI_NO"));r.setTorihikisakiCd(d.get("TORIHIKISAKI_CD"));r.setTorihikisakiNk(d.get("TORIHIKISAKI_NK"));r.setNonyusakiNk(d.get("NONYUSAKI_NK"));r.setSupportId(d.get("SUPPORT_ID"));r.setPlantNk(d.get("PLANT_NK"));r.setKeiyakuNo(d.get("KEIYAKU_NO"));r.setKeiyakuDt(d.get("KEIYAKU_DT"));r.setKaiyakuDt(d.get("KAIYAKU_DT"));r.setShoninJotai(d.get("SHONINJOTAI"));return r;}).toList();
 }
 public void discard(boolean um,BigDecimal contract,BigDecimal source,String version,String user){
  String table=um?"MCM_UK_KEIYAKU":"MCM_TK_KEIYAKU",key=um?"UK_KEIYAKU_ID":"TK_KEIYAKU_ID";
  Timestamp ts=version==null||version.isBlank()?null:Timestamp.valueOf(version);
  int count=jdbc.update("UPDATE MCM."+table+" SET JOTAI='3',LASTUPDATE_DT=CURRENT_TIMESTAMP,LASTUPDATE_BY=? WHERE "+key+"=? AND JOTAI IN ('1','2') AND (LASTUPDATE_DT=? OR (LASTUPDATE_DT IS NULL AND ? IS NULL))",user,contract,ts,ts);
  if(count!=1)throw new IllegalStateException("契約が変更されています。再検索してください。");
  String sourceTable=um?"MCM_UM_KIHON_MITSUMORI":"MCM_TM_KEIYAKUJIKAN",sourceKey=um?"UM_KIHON_MITSUMORI_ID":"TM_KEIYAKUJIKAN_ID";
  if(jdbc.update("UPDATE MCM."+sourceTable+" SET JOTAI='1',LASTUPDATE_DT=CURRENT_TIMESTAMP,LASTUPDATE_BY=? WHERE "+sourceKey+"=?",user,source)!=1)throw new IllegalStateException("元の見積が見つかりません。再検索してください。");
 }
 public LocalDate getMinSiharaiTsuki(BigDecimal id){var d=jdbc.queryForObject("SELECT MIN(D.TSUKI) FROM MCM.MCM_TK_SIHARAIMEISAI D JOIN MCM.MCM_TK_SIHARAI S ON S.TK_SIHARAI_ID=D.TK_SIHARAI_ID JOIN MCM.MCM_TK_KIKAN K ON K.TK_KIKAN_ID=S.TK_KIKAN_ID WHERE K.TK_KEIYAKU_ID=?",java.sql.Date.class,id);return d==null?null:d.toLocalDate();}
}
