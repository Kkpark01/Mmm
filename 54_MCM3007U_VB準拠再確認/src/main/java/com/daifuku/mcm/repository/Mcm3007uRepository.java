package com.daifuku.mcm.repository;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.YearMonth;
import com.daifuku.mcm.common.DateUtils;
import com.daifuku.mcm.exception.ExclusiveControlException;
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
 private static void textFilter(StringBuilder sql,List<Object> args,String column,String value,boolean prefix) {
  if(value==null||value.isBlank())return;
  if(prefix){sql.append(" AND "+column+" LIKE ?");args.add(value.trim()+"%");}
  else{sql.append(" AND CHARINDEX(?,"+column+") > 0");args.add(value.trim());}
 }
 private static void names(StringBuilder sql,List<Object> args,String alias,String value,boolean factory,boolean like) {
  if(value==null||value.isBlank())return;
  String expression="UPPER(CONCAT("+alias+".NONYUSAKI_NK,' ',"+alias+".KYUNONYUSAKI_NK,' ',"+(factory?alias+".NONYUSAKIKOJO_NK,' ',":"")+alias+".NONYUSAKIKANA_KN,' ',"+alias+".NONYUSAKIEIMEI_EN))";
  sql.append(like?" AND "+expression+" LIKE UPPER(?)":" AND CHARINDEX(UPPER(?),"+expression+") > 0");
  args.add(like?"%"+value.trim()+"%":value.trim());
 }
 private static void plantName(StringBuilder sql,List<Object> args,String alias,String value,boolean like) {
  if(value==null||value.isBlank())return;
  sql.append(like?" AND UPPER("+alias+".PLANT_NK) LIKE UPPER(?)":" AND CHARINDEX(UPPER(?),UPPER("+alias+".PLANT_NK)) > 0");args.add(like?"%"+value.trim()+"%":value.trim());
 }
 private static void filters(StringBuilder sql,List<Object> args,String alias,String cd,String nk,String sid,String pnk,boolean prefix) {
  textFilter(sql,args,alias+".NONYUSAKI_CD",cd,prefix);names(sql,args,alias,nk,false,prefix);
  textFilter(sql,args,alias+".SUPPORT_ID",sid,prefix);plantName(sql,args,alias,pnk,prefix);
 }
 public static Map<String,String> display(Map<String,Object> row) {
  var d=new LinkedHashMap<String,String>();
  row.forEach((rawKey,v)->{
   String k=rawKey.toUpperCase(Locale.ROOT),text=v==null?"":v.toString();
   if(v instanceof BigDecimal b)text=b.stripTrailingZeros().toPlainString();
   else if(k.endsWith("_DT")) {
    if(v instanceof Timestamp t)text=DateUtils.formatDateTime(t.toLocalDateTime()).substring(0,10);
    else if(v instanceof java.sql.Date dt)text=dt.toLocalDate().toString().replace('-','/');
    else if(v instanceof LocalDateTime dt)text=DateUtils.formatDateTime(dt).substring(0,10);
    else if(v instanceof LocalDate dt)text=dt.toString().replace('-','/');
   }
   // CONTRACT_VERSION/SOURCE_VERSIONは更新判定用の原値。年月日へ丸めない。
   d.put(k,text);
  });
  return d;
 }
 private static String approval(String value) {
  return switch(value==null?"":value){case "0"->"作成中";case "1"->"審査中";case "2"->"承認中";case "3"->"承認済";case "4"->"差戻中";default->value;};
 }
 public static BigDecimal id(Map<String,String>d,String key){String v=d.get(key);return v==null||v.isBlank()?null:new BigDecimal(v);}
 private static String state(String v){return switch(v==null?"":v){case "1"->"見積";case "2"->"契約";case "3"->"破棄";case "4"->"解約";default->v;};}
 public List<Grid1RowForm> searchNonyusaki(String cd,String nk,String sid,String pnk){return searchNonyusaki(cd,nk,sid,pnk,List.of("1","2","3","4"));}
 public List<Grid1RowForm> searchNonyusaki(String cd,String nk,String sid,String pnk,List<String> status){
  if(status.isEmpty())return List.of();var st=states(status);var args=new ArrayList<Object>(st);
  var sql=new StringBuilder("SELECT DISTINCT N.*,P.PLANT_ID,P.SUPPORT_ID,P.PLANT_NK,P.NONYU_DT,P.TEKKYO_DT,P.HOSYUSYUSOKU_DT FROM MCM.MCM_MA_NONYUSAKI N JOIN MCM.MCM_MA_PLANT P ON P.NONYUSAKI_ID=N.NONYUSAKI_ID JOIN MCM.MCM_2004_V V ON V.NONYUSAKI_ID=N.NONYUSAKI_ID AND V.PLANT_ID=P.PLANT_ID WHERE V.JOTAI IN ("+marks(st.size())+")");
  textFilter(sql,args,"N.NONYUSAKI_CD",cd,true);names(sql,args,"N",nk,true,false);textFilter(sql,args,"P.SUPPORT_ID",sid,true);plantName(sql,args,"P",pnk,true);sql.append(" ORDER BY N.NONYUSAKI_ID,P.PLANT_ID");
  return jdbc.queryForList(sql.toString(),args.toArray()).stream().map(m->{var d=display(m);var r=new Grid1RowForm();r.setDisplay(d);r.setNonyusakiId(id(d,"NONYUSAKI_ID"));r.setNonyusakiCd(d.get("NONYUSAKI_CD"));r.setNonyusakiNk(d.get("NONYUSAKI_NK"));r.setSupportId(d.get("SUPPORT_ID"));r.setPlantNk(d.get("PLANT_NK"));return r;}).toList();
 }
 public List<Grid2RowForm> searchUva(String cd,String nk,String sid,String pnk,List<String> status){
  if(status.isEmpty())return List.of();var st=states(status);var args=new ArrayList<Object>(st);
  var sql=new StringBuilder("SELECT V.*,K.JOTAI AS CONTRACT_STATE,K.LASTUPDATE_DT AS CONTRACT_VERSION,Q.JOTAI AS SOURCE_STATE,Q.LASTUPDATE_DT AS SOURCE_VERSION FROM MCM.MCM_2004_V V LEFT JOIN MCM.MCM_UK_KEIYAKU K ON K.UK_KEIYAKU_ID=V.UK_KEIYAKU_ID LEFT JOIN MCM.MCM_UM_KIHON_MITSUMORI Q ON Q.UM_KIHON_MITSUMORI_ID=V.UM_KIHON_MITSUMORI_ID WHERE V.JOTAI IN ("+marks(st.size())+")");filters(sql,args,"V",cd,nk,sid,pnk,false);sql.append(" ORDER BY V.UM_MITSUMORI_NO,V.KEIYAKU_NO,V.UM_KIHON_MITSUMORI_ID,V.UK_KEIYAKU_ID");
  return jdbc.queryForList(sql.toString(),args.toArray()).stream().map(m->{var d=display(m);d.put("JOTAI",state(d.get("JOTAI")));d.put("SYOUNIN_JOTAI",approval(d.get("SYOUNIN_JOTAI")));var r=new Grid2RowForm();r.setDisplay(d);r.setUkKeiyakuId(id(d,"UK_KEIYAKU_ID"));r.setUmKihonMitsumoriId(id(d,"UM_KIHON_MITSUMORI_ID"));r.setJotai(d.get("JOTAI"));r.setShoruiNo(d.get("UM_MITSUMORI_NO"));r.setNonyusakiNk(d.get("NONYUSAKI_NK"));r.setSupportId(d.get("SUPPORT_ID"));r.setPlantNk(d.get("PLANT_NK"));r.setKeiyakuNo(d.get("KEIYAKU_NO"));r.setKeiyakuDt(d.get("KEIYAKU_DT"));r.setKaiyakuDt(d.get("KAIYAKU_DT"));r.setShoninJotai(d.get("SYOUNIN_JOTAI"));return r;}).toList();
 }
 public List<Grid3RowForm> searchTka(String cd,String nk,String sid,String pnk,List<String> status){
  if(status.isEmpty())return List.of();var st=states(status);var args=new ArrayList<Object>(st);
  var sql=new StringBuilder("SELECT DISTINCT V.*,B.UM_KIHON_MITSUMORI_ID,K.JOTAI AS CONTRACT_STATE,K.LASTUPDATE_DT AS CONTRACT_VERSION,TM.JOTAI AS SOURCE_STATE,TM.LASTUPDATE_DT AS SOURCE_VERSION,CASE WHEN K.JOTAI='3' AND TM.JOTAI='1' THEN N'見積' ELSE V.JOTAI END AS EFFECTIVE_JOTAI,CASE WHEN K.JOTAI='3' AND TM.JOTAI='1' THEN 1 ELSE 0 END AS RETURNED_ESTIMATE,TRY_CAST(V.KEIYAKUJIKANTAI AS DECIMAL(18,6)) AS SORT_HOURS FROM MCM.MCM_1003_V V JOIN MCM.MCM_TM_KIKIKOSEI C ON C.TM_IRAI_ID=V.TM_IRAI_ID JOIN MCM.MCM_TM_KIKIMEISAI D ON D.TM_KIKIKOSEI_ID=C.TM_KIKIKOSEI_ID JOIN MCM.MCM_TM_KOTAIMEISAI E ON E.TM_KIKIMEISAI_ID=D.TM_KIKIMEISAI_ID JOIN MCM.MCM_UM_KOTAIMEISAI UE ON UE.KOTAIKANRI_ID=E.KOTAIKANRI_ID JOIN MCM.MCM_UM_KIKIMEISAI UD ON UD.UM_KIKIMEISAI_ID=UE.UM_KIKIMEISAI_ID JOIN MCM.MCM_UM_KIKIKOSEI UC ON UC.UM_KIKIKOSEI_ID=UD.UM_KIKIKOSEI_ID JOIN MCM.MCM_UM_KIHON_BRAND B ON B.UM_KIHON_BRAND_ID=UC.UM_KIHON_BRAND_ID LEFT JOIN MCM.MCM_TK_KEIYAKU K ON K.TK_KEIYAKU_ID=V.TK_KEIYAKU_ID LEFT JOIN MCM.MCM_TM_KEIYAKUJIKAN TM ON TM.TM_KEIYAKUJIKAN_ID=V.TM_KEIYAKUJIKAN_ID WHERE (CASE WHEN K.JOTAI='3' AND TM.JOTAI='1' THEN N'見積' ELSE V.JOTAI END) IN ("+marks(st.size())+")");filters(sql,args,"V",cd,nk,sid,pnk,true);sql.append(" ORDER BY V.TM_IRAI_NO DESC,SORT_HOURS,V.TM_KEIYAKUJIKAN_ID,B.UM_KIHON_MITSUMORI_ID");
  return jdbc.queryForList(sql.toString(),args.toArray()).stream().map(m->{var d=display(m);d.put("JOTAI",state(d.get("JOTAI")));d.put("JOTAI",state(d.remove("EFFECTIVE_JOTAI")));if("1".equals(d.remove("RETURNED_ESTIMATE"))){d.put("KEIYAKU_KEIYAKU","契約");d.put("KEIYAKU_DEL","");}d.remove("SORT_HOURS");d.put("SHONINJOTAI",approval(d.get("SHONINJOTAI")));
   var r=new Grid3RowForm();r.setDisplay(d);r.setTkKeiyakuId(id(d,"TK_KEIYAKU_ID"));r.setTmKeiyakujikanId(id(d,"TM_KEIYAKUJIKAN_ID"));r.setJotai(d.get("JOTAI"));r.setIraino(d.get("TM_IRAI_NO"));r.setTorihikisakiCd(d.get("TORIHIKISAKI_CD"));r.setTorihikisakiNk(d.get("TORIHIKISAKI_NK"));r.setNonyusakiNk(d.get("NONYUSAKI_NK"));r.setSupportId(d.get("SUPPORT_ID"));r.setPlantNk(d.get("PLANT_NK"));r.setKeiyakuNo(d.get("KEIYAKU_NO"));r.setKeiyakuDt(d.get("KEIYAKU_DT"));r.setKaiyakuDt(d.get("KAIYAKU_DT"));r.setShoninJotai(d.get("SHONINJOTAI"));return r;}).toList();
 }
 /** 再検索とUPDATEの間に変わった契約/元見積も、同じトランザクション内で検出する。 */
 public void verifyDiscard(boolean um,BigDecimal contract,BigDecimal source,Map<String,String> snapshot) {
  String table=um?"MCM_UK_KEIYAKU":"MCM_TK_KEIYAKU",key=um?"UK_KEIYAKU_ID":"TK_KEIYAKU_ID";
  String sourceTable=um?"MCM_UM_KIHON_MITSUMORI":"MCM_TM_KEIYAKUJIKAN",sourceKey=um?"UM_KIHON_MITSUMORI_ID":"TM_KEIYAKUJIKAN_ID";
  checkLocked(table,key,contract,snapshot,"CONTRACT");checkLocked(sourceTable,sourceKey,source,snapshot,"SOURCE");
  var count=jdbc.queryForObject("SELECT COUNT(*) FROM MCM."+(um?"MCM_2004_V":"MCM_1003_V")+" WHERE "+key+"=? AND "+sourceKey+"=? AND KEIYAKU_DEL IS NOT NULL AND LTRIM(RTRIM(KEIYAKU_DEL)) <> ''",Integer.class,contract,source);
  if(count==null||count==0||!Set.of("1","2","4").contains(snapshot.getOrDefault("CONTRACT_STATE","")))throw conflict();
 }
 private void checkLocked(String table,String key,BigDecimal value,Map<String,String> snapshot,String prefix) {
  var rows=jdbc.queryForList("SELECT JOTAI AS CURRENT_STATE,LASTUPDATE_DT AS CURRENT_VERSION FROM MCM."+table+" WITH (UPDLOCK, HOLDLOCK) WHERE "+key+"=?",value);
  if(rows.size()!=1||!snapshot.containsKey(prefix+"_STATE")||!snapshot.containsKey(prefix+"_VERSION"))throw conflict();
  var current=display(rows.get(0));if(current.get("CURRENT_STATE").isBlank())throw conflict();
  if(!Objects.equals(current.get("CURRENT_STATE"),snapshot.get(prefix+"_STATE"))||!Objects.equals(current.get("CURRENT_VERSION"),snapshot.get(prefix+"_VERSION")))throw conflict();
 }
 private static ExclusiveControlException conflict(){return new ExclusiveControlException("他のユーザーによりデータが更新されています。再検索してください。");}
 private static Timestamp version(String value){return value==null||value.isBlank()?null:Timestamp.valueOf(value.replace('T',' '));}
 public void discard(boolean um,BigDecimal contract,BigDecimal source,Map<String,String> snapshot,String userName) {
  String table=um?"MCM_UK_KEIYAKU":"MCM_TK_KEIYAKU",key=um?"UK_KEIYAKU_ID":"TK_KEIYAKU_ID";
  Timestamp ts=version(snapshot.get("CONTRACT_VERSION"));
  int count=jdbc.update("UPDATE MCM."+table+" SET JOTAI='3',LASTUPDATE_DT=CURRENT_TIMESTAMP,LASTUPDATE_BY=? WHERE "+key+"=? AND JOTAI=? AND (LASTUPDATE_DT=? OR (LASTUPDATE_DT IS NULL AND ? IS NULL))",userName,contract,snapshot.get("CONTRACT_STATE"),ts,ts);
  if(count!=1)throw conflict();
  String sourceTable=um?"MCM_UM_KIHON_MITSUMORI":"MCM_TM_KEIYAKUJIKAN",sourceKey=um?"UM_KIHON_MITSUMORI_ID":"TM_KEIYAKUJIKAN_ID";
  Timestamp sourceTs=version(snapshot.get("SOURCE_VERSION"));
  count=jdbc.update("UPDATE MCM."+sourceTable+" SET JOTAI='1',LASTUPDATE_DT=CURRENT_TIMESTAMP,LASTUPDATE_BY=? WHERE "+sourceKey+"=? AND (JOTAI=? OR (JOTAI IS NULL AND ? IS NULL)) AND (LASTUPDATE_DT=? OR (LASTUPDATE_DT IS NULL AND ? IS NULL))",userName,source,snapshot.get("SOURCE_STATE"),snapshot.get("SOURCE_STATE"),sourceTs,sourceTs);
  if(count!=1)throw conflict();
 }
 public YearMonth currentMonth(){return YearMonth.from(jdbc.queryForObject("SELECT CURRENT_TIMESTAMP",Timestamp.class).toLocalDateTime());}
 public LocalDate getMinSiharaiTsuki(BigDecimal id){var d=jdbc.queryForObject("SELECT MIN(D.TSUKI) FROM MCM.MCM_TK_SIHARAIMEISAI D JOIN MCM.MCM_TK_SIHARAI S ON S.TK_SIHARAI_ID=D.TK_SIHARAI_ID JOIN MCM.MCM_TK_KIKAN K ON K.TK_KIKAN_ID=S.TK_KIKAN_ID WHERE K.TK_KEIYAKU_ID=?",java.sql.Date.class,id);return d==null?null:d.toLocalDate();}
 public boolean hasBrand(BigDecimal plantId){return plantId!=null&&jdbc.queryForObject("SELECT COUNT(*) FROM MCM.MCM_MA_BRAND_KOSEI WHERE PLANT_ID=?",Integer.class,plantId)>0;}
}
