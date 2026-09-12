package com.daifuku.mcm.service;

import java.math.BigDecimal;
import java.sql.Timestamp;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.security.MessageDigest;
import java.nio.charset.StandardCharsets;
import org.springframework.beans.BeanWrapperImpl;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.transaction.annotation.Isolation;
import com.daifuku.mcm.common.BaseService;
import com.daifuku.mcm.form.*;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;

/** 取引先見積内容。VBのTM期間・単価・点検を対象見積内で読み書きする。 */
@Service
public class Mcm1004uService extends BaseService {
    @Autowired private JdbcTemplate jdbc;
    @Autowired private Mcm2004uService permissions;
    private static final String TANKA_SQL = "SELECT TMC.TM_KIKIKOSEI_ID, TMC.TM_IRAI_ID, TMC.KIKIKOSEI_ID, TMC.KIKIKOSEI_NK, TMC.SET_NM, TMC.TEHAISEIBAN, TMD.SEIZOMAKER_ID, TMD.SEIZOMAKER_NK, TMD.KIKIMEISAI_ID, TMD.KIKIHINMEI_NK, TMD.KIKIKATASHIKI, TMD.SURYO_NM, TMD.ATSUKAIKIKI_ID, TMF.TM_TANKA_ID AS TM_TANKA_ID, ISNULL(TMF.TM_KIKAN_ID, 0) AS TM_KIKAN_ID, ISNULL(TMF.HYOJUN_KIN, 0) AS HYOJUN_KIN, ISNULL(TMF.HYOJUN_KIN, 0) * ISNULL(TMD.SURYO_NM, 0) AS HYOJUNKEI_KIN, ISNULL(TMF.SIKIRI_KIN, 0) AS SIKIRI_KIN, ISNULL(TMF.SIKIRI_KIN, 0) * ISNULL(TMD.SURYO_NM, 0) AS SIKIRIKEI_KIN, ISNULL(TMF.PACK_FLG, 0) AS PACK_FLG, TMF.KEIYAKUNAIYO, TMF.KEIYAKU_NO, TMF.TORIHOSYUJIKAN_ID, ISNULL(TMF.TENKENUMU, 0) AS TENKENUMU, TMF.HOSYUHOHO, TMF.SERVICEKEITAI, TMF.BIKO, TMF.LASTUPDATE_BY, TMF.LASTUPDATE_DT, TMF.CREATED_BY, TMF.CREATED_DT, TMD.TM_KIKIMEISAI_ID, N'照会' AS SHOKAI FROM MCM.MCM_TM_KIKIKOSEI TMC INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD ON TMC.TM_KIKIKOSEI_ID = TMD.TM_KIKIKOSEI_ID LEFT JOIN MCM.MCM_MA_KIKIMEISAI MAF ON TMD.KIKIMEISAI_ID = MAF.KIKIMEISAI_ID LEFT JOIN MCM.MCM_MA_KIKIKOSEI MAE ON TMC.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID LEFT JOIN MCM.MCM_TM_TANKA TMF ON TMD.TM_KIKIMEISAI_ID = TMF.TM_KIKIMEISAI_ID   AND TMF.TM_KIKAN_ID = ? WHERE TMC.TM_IRAI_ID = ? ORDER BY TMC.HYOJIJUN, TMD.HYOJIJUN, MAE.HYOJIJUN, MAF.HYOJIJUN";
    private static final String TENKEN_SQL = "SELECT DISTINCT ISNULL(TMH.TM_KIKAN_ID, 0) AS TM_KIKAN_ID, TMH.TENKENKAISU, TMH.TENKENKANOYOBI, ISNULL(TMH.YAKANTAIOUMU, 0) AS YAKANTAIOUMU, TMH.BIKO, TMH.CREATED_DT, TMH.CREATED_BY, TMH.LASTUPDATE_DT, TMH.LASTUPDATE_BY, TMC.TM_IRAI_ID, TMD.TM_KIKIKOSEI_ID, TMH.TM_TENKEN_ID AS TM_TENKEN_ID, VMA.MAM_OYAKIKIBUNRUI_CD AS OYAKIKIBUNRUI_CD, VMA.MAE_KIKIKOSEI_ID AS KIKIKOSEI_ID, VMA.MAE_KIKIKOSEI_NK AS KIKIKOSEI_NK, VMA.MAF_KIKIMEISAI_ID AS KIKIMEISAI_ID, VMA.MAF_KIKIHINMEI_NK AS KIKIHINMEI_NK, VMA.MAF_KIKIKATASHIKI AS KIKIKATASHIKI, VMA.MAE_HYOJIJUN, VMA.MAF_HYOJIJUN FROM MCM.MCM_TM_KIKIKOSEI TMC LEFT JOIN MCM.MCM_MA_KIKIKOSEI MAE ON MAE.KIKIKOSEI_ID = TMC.KIKIKOSEI_ID INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD ON TMC.TM_KIKIKOSEI_ID = TMD.TM_KIKIKOSEI_ID INNER JOIN MCM.MCM_TM_OYAKOTAI_V TMA ON TMD.TM_KIKIMEISAI_ID = TMA.TMD_TM_KIKIMEISAI_ID INNER JOIN MCM.MCM_MA_OYAKOTAI_V VMA   ON TMA.TMC_KIKIKOSEI_ID = VMA.MAE_KIKIKOSEI_ID   AND TMA.MAM_OYAKIKIBUNRUI_CD = VMA.MAM_OYAKIKIBUNRUI_CD LEFT JOIN MCM.MCM_TM_TENKEN TMH   ON TMA.MAM_OYAKIKIBUNRUI_CD = TMH.OYAKIKIBUNRUI_CD   AND TMA.TMC_TM_KIKIKOSEI_ID = TMH.TM_KIKIKOSEI_ID   AND TMH.TM_KIKAN_ID = ? LEFT JOIN MCM.MCM_MA_KIKIMEISAI MAF ON TMD.KIKIMEISAI_ID = MAF.KIKIMEISAI_ID WHERE TMC.TM_IRAI_ID = ? ORDER BY VMA.MAE_HYOJIJUN, VMA.MAF_HYOJIJUN";
    public static final Map<String,String> METHODS = options("F","ｵﾝｻｲﾄ","S","ｾﾝﾄﾞﾊﾞｯｸ","C","ｺﾝﾃｯｸ製品","H","持ち帰り","I","ｽﾎﾟｯﾄ","T","TEL対応");
    public static final Map<String,String> SERVICES = options("1","1：維持保守","2","2：出張修理","3","3：持込修理","4","4：引取修理","5","5：ﾌﾟﾘﾝﾀ出張修理","6","6：ｿﾌﾄｻﾎﾟｰﾄｻｰﾋﾞｽ");
    public static final Map<String,String> DAYS = options("1","月～金","2","月～土","3","月～日","4","土・日");
    private static final Map<String,String> WEEK = options("1","日","2","月","3","火","4","水","5","木","6","金","7","土");
    private static Map<String,String> options(String...v){var m=new LinkedHashMap<String,String>();for(int i=0;i<v.length;i+=2)m.put(v[i],v[i+1]);return Collections.unmodifiableMap(m);}
    public boolean canView(String user,String screen){return Set.of("1","2").contains(Objects.toString(permissions.getAuthority(user,screen),"0"));}
    public boolean canUpdate(String user){return "2".equals(permissions.getAuthority(user,"MCM1004U"));}
    public boolean canUnlock(String user){return canUpdate(user)&&permissions.isMaintenance(user);}
    private static String str(Object value){return value==null?"":value instanceof BigDecimal n?n.stripTrailingZeros().toPlainString():value.toString();}
    private static BigDecimal num(Object v){return v==null?null:v instanceof BigDecimal n?n:new BigDecimal(v.toString());}
    private static BigDecimal zero(BigDecimal v){return v==null?BigDecimal.ZERO:v;}
    private static LocalDate day(Object v){return v==null?null:v instanceof Timestamp t?t.toLocalDateTime().toLocalDate():v instanceof java.sql.Date d?d.toLocalDate():LocalDate.parse(v.toString().substring(0,10));}
    private static String audit(Object v){return v instanceof Timestamp t?t.toLocalDateTime().format(DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss")):str(v);}
    private static String property(String col){var b=new StringBuilder();boolean upper=false;for(char c:col.toLowerCase(Locale.ROOT).toCharArray()){if(c=='_'){upper=true;continue;}b.append(upper?Character.toUpperCase(c):c);upper=false;}return b.toString();}
    private static <T>T map(T target,Map<String,Object> row){var b=new BeanWrapperImpl(target);for(var e:row.entrySet()){String p=property(e.getKey());if(!b.isWritableProperty(p))continue;Class<?> type=b.getPropertyType(p);Object v=e.getValue();if(type==String.class)v=audit(v);else if(type==LocalDate.class)v=day(v);else if(type==Long.class)v=v==null?null:num(v).longValueExact();else if(type==BigDecimal.class)v=v==null?null:new BigDecimal(str(v));b.setPropertyValue(p,v);}return target;}
    private static String state(String s){return switch(s){case "0"->"依頼";case "1"->"見積";case "2"->"契約";case "3"->"破棄";case "4"->"解約";case "9"->"作成中";default->s;};}

    @Transactional(readOnly=true,isolation=Isolation.REPEATABLE_READ)
    public Mcm1004uForm loadScreen(BigDecimal id,int mode,boolean unlocked){
        var rows=jdbc.queryForList("SELECT M.*,K.TM_KEIYAKUJIKAN_ID,K.KEIYAKUJIKANTAI,K.JOTAI FROM MCM.MCM_TM_MITSUMORI M JOIN MCM.MCM_TM_KEIYAKUJIKAN K ON K.TM_IRAI_ID=M.TM_IRAI_ID WHERE K.TM_KEIYAKUJIKAN_ID=?",id);
        if(rows.size()!=1)throw new IllegalStateException("対象の見積依頼がありません。検索画面から開き直してください。");
        var f=map(new Mcm1004uForm(),rows.get(0));f.setMitsumoriDt(f.getMitsumoriDt().split(" ")[0]);f.setSeniMotoKbn(mode);f.setLockReleaseFlg(unlocked);f.setJotaiDisplay(state(f.getJotai()));
        f.setViewFlg(mode!=2&&(unlocked||Set.of("0","1","9").contains(f.getJotai())&&!checkUmMitsumori(id)&&!checkTkKiyaku(id)&&checkPlantId(id,f.getPlantId())));
        var periods=jdbc.queryForList("SELECT * FROM MCM.MCM_TM_KIKAN WHERE TM_KEIYAKUJIKAN_ID=? ORDER BY KAISI_DT,TM_KIKAN_ID",id);
        if(periods.isEmpty()){var tab=new Mcm1004uTabForm();tab.setTmKikanId(0L);f.getTabs().add(tab);}else for(var p:periods)f.getTabs().add(map(new Mcm1004uTabForm(),p));
        for(int i=0;i<f.getTabs().size();i++){
            var tab=f.getTabs().get(i);tab.setTabLabel(i==0?"初回":(i+1)+"回");
            for(var r:jdbc.queryForList(TANKA_SQL,tab.getTmKikanId(),f.getTmIraiId()))tab.getTankaRows().add(map(new Mcm1004uTankaRow(),r));
            for(var r:jdbc.queryForList(TENKEN_SQL,tab.getTmKikanId(),f.getTmIraiId()))tab.getTenkenRows().add(map(new Mcm1004uTenkenRow(),r));
            // The VB joins may produce the same inspection row for several devices. Keep one persisted classification row.
            var unique=new LinkedHashMap<String,Mcm1004uTenkenRow>();for(var r:tab.getTenkenRows())unique.putIfAbsent(tenkenKey(r),r);tab.setTenkenRows(new ArrayList<>(unique.values()));
            recalculate(tab);
        }
        var choices=new LinkedHashMap<String,String>();
        for(var r:jdbc.queryForList("SELECT * FROM MCM.MCM_MA_TORIHOSYUJIKAN WHERE TORIHIKISAKI_ID=? ORDER BY HYOJIJUN,TORIHOSYUJIKAN_ID",f.getTorihikisakiId())){
            String label=str(r.get("HOSYUJIKAN_DT"))+"H";if(!str(r.get("HOSYU_KBN")).isBlank())label+="("+str(r.get("HOSYU_KBN"))+")";
            if(r.get("KAISIYOBI")!=null)label+=" "+WEEK.getOrDefault(str(r.get("KAISIYOBI")),"")+"～"+WEEK.getOrDefault(str(r.get("SYURYOYOBI")),"");
            if(r.get("KAISIJIKAN_DT")!=null)label+=" "+str(r.get("KAISIJIKAN_DT"))+"～"+str(r.get("SYURYOJIKAN_DT"));
            if(r.get("REIGAIYOBI")!=null)label+=" ("+WEEK.getOrDefault(str(r.get("REIGAIYOBI")),"")+" "+str(r.get("REIGAIKAISIJIKAN_DT"))+"～"+str(r.get("REIGAISYURYOJIKAN_DT"))+")";
            choices.put(str(r.get("TORIHOSYUJIKAN_ID")),label);
        }
        f.setHoshuOptions(choices);f.setRevision(revision(id,f.getTmIraiId()));return f;
    }
    public boolean checkUmMitsumori(BigDecimal id){return jdbc.queryForObject("SELECT COUNT(*) FROM MCM.MCM_UM_TANKA U JOIN MCM.MCM_TM_TANKA T ON T.TM_TANKA_ID=U.TM_TANKA_ID JOIN MCM.MCM_TM_KIKAN P ON P.TM_KIKAN_ID=T.TM_KIKAN_ID WHERE P.TM_KEIYAKUJIKAN_ID=?",Integer.class,id)>0;}
    public boolean checkTkKiyaku(BigDecimal id){return jdbc.queryForObject("SELECT COUNT(*) FROM MCM.MCM_TM_MITSUMORI M JOIN MCM.MCM_TM_KEIYAKUJIKAN K ON K.TM_IRAI_ID=M.TM_IRAI_ID JOIN MCM.MCM_TM_KEIYAKU_NO_V V ON V.TMB_TM_KEIYAKUJIKAN_ID=K.TM_KEIYAKUJIKAN_ID JOIN MCM.MCM_MA_NONYUSAKI N ON N.NONYUSAKI_ID=M.NONYUSAKI_ID WHERE K.TM_KEIYAKUJIKAN_ID=?",Integer.class,id)>0;}
    public boolean checkPlantId(BigDecimal id,BigDecimal plant){var ids=jdbc.queryForList("SELECT DISTINCT A.PLANT_ID FROM MCM.MCM_TM_KEIYAKUJIKAN K JOIN MCM.MCM_TM_KIKIKOSEI C ON C.TM_IRAI_ID=K.TM_IRAI_ID JOIN MCM.MCM_MA_KIKIKOSEI A ON A.KIKIKOSEI_ID=C.KIKIKOSEI_ID WHERE K.TM_KEIYAKUJIKAN_ID=?",BigDecimal.class,id);return ids.size()==1&&same(ids.get(0),plant);}
    private String revision(BigDecimal id,BigDecimal irai){
        var values=new ArrayList<Object>();
        values.add(jdbc.queryForList("SELECT * FROM MCM.MCM_TM_MITSUMORI WHERE TM_IRAI_ID=?",irai));
        values.add(jdbc.queryForList("SELECT * FROM MCM.MCM_TM_KEIYAKUJIKAN WHERE TM_KEIYAKUJIKAN_ID=?",id));
        values.add(jdbc.queryForList("SELECT * FROM MCM.MCM_TM_KIKAN WHERE TM_KEIYAKUJIKAN_ID=? ORDER BY TM_KIKAN_ID",id));
        for(String[] pair:new String[][]{{"MCM_TM_TANKA","TM_TANKA_ID"},{"MCM_TM_TENKEN","TM_TENKEN_ID"}})values.add(jdbc.queryForList("SELECT X.* FROM MCM."+pair[0]+" X JOIN MCM.MCM_TM_KIKAN P ON P.TM_KIKAN_ID=X.TM_KIKAN_ID WHERE P.TM_KEIYAKUJIKAN_ID=? ORDER BY X."+pair[1],id));
        values.add(jdbc.queryForList("SELECT * FROM MCM.MCM_TM_KIKIKOSEI WHERE TM_IRAI_ID=? ORDER BY TM_KIKIKOSEI_ID",irai));
        values.add(jdbc.queryForList("SELECT D.* FROM MCM.MCM_TM_KIKIMEISAI D JOIN MCM.MCM_TM_KIKIKOSEI C ON C.TM_KIKIKOSEI_ID=D.TM_KIKIKOSEI_ID WHERE C.TM_IRAI_ID=? ORDER BY D.TM_KIKIMEISAI_ID",irai));
        try{return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(values.toString().getBytes(StandardCharsets.UTF_8)));}catch(java.security.NoSuchAlgorithmException e){throw new IllegalStateException(e);}
    }
    private static String tenkenKey(Mcm1004uTenkenRow r){return str(r.getTmKikikoseiId())+"/"+str(r.getOyakikibunruiCd());}
    public static void recalculate(Mcm1004uTabForm t){BigDecimal h=BigDecimal.ZERO,s=BigDecimal.ZERO;for(var r:t.getTankaRows()){r.setHyojunkeiKin(zero(r.getHyojunKin()).multiply(zero(r.getSuryoNm())));r.setSikirikeiKin(zero(r.getSikiriKin()).multiply(zero(r.getSuryoNm())));h=h.add(r.getHyojunkeiKin());s=s.add(r.getSikirikeiKin());}t.setHyojungokeiKin(h);t.setSikirisyokeiKin(s);t.setSikirigokeiKin(s.subtract(zero(t.getSyusseinebikiKin())));}
    public List<String> validatePeriods(List<Mcm1004uTabForm> tabs){
        var errors=new ArrayList<String>();if(tabs==null||tabs.isEmpty()){errors.add("期間がありません。");return errors;}
        for(int i=0;i<tabs.size();i++){var t=tabs.get(i);if(t.getKaisiDt()==null){errors.add((i+1)+"番目の期間の開始日は必須です。");continue;}if(t.getKaisiDt().getYear()<1||t.getKaisiDt().getYear()>9999||t.getSyuryoDt()!=null&&(t.getSyuryoDt().getYear()<1||t.getSyuryoDt().getYear()>9999))errors.add("期間は1～9999年で指定してください。");if(i<tabs.size()-1&&t.getSyuryoDt()==null)errors.add("最終期間以外の終了日は必須です。");if(t.getSyuryoDt()!=null&&t.getKaisiDt().isAfter(t.getSyuryoDt()))errors.add("開始日は終了日以前に設定してください。");if(i>0&&tabs.get(i-1).getSyuryoDt()!=null&&!tabs.get(i-1).getSyuryoDt().plusDays(1).equals(t.getKaisiDt()))errors.add("前の期間の終了日の翌日を開始日にしてください。");}return errors;
    }
    public List<String> validateForTabAdd(Mcm1004uForm f){var e=validatePeriods(f.getTabs());if(!e.isEmpty())return e;var t=f.getTabs().get(f.getTabs().size()-1);if(t.getSyuryoDt()==null)e.add("タブを追加するには最終期間の終了日が必要です。");else if(t.getSyuryoDt().getYear()>=9999)e.add("9999年を超える期間は追加できません。");return e;}
    public void addNewTab(Mcm1004uForm f){var errors=validateForTabAdd(f);if(!errors.isEmpty())throw new IllegalStateException(String.join(" ",errors));var last=f.getTabs().get(f.getTabs().size()-1);var t=copy(last);t.setTmKikanId(0L);t.setTabLabel((f.getTabs().size()+1)+"回");t.setKaisiDt(last.getSyuryoDt().plusDays(1));t.setSyuryoDt(null);t.setCreatedDt("");t.setCreatedBy("");t.setLastupdateDt("");t.setLastupdateBy("");for(var r:t.getTankaRows()){r.setTmTankaId(null);r.setTmKikanId(BigDecimal.ZERO);}for(var r:t.getTenkenRows()){r.setTmTenkenId(null);r.setTmKikanId(BigDecimal.ZERO);}f.getTabs().add(t);f.setSelectedIndex(f.getTabs().size()-1);}
    public void adjustDates(Mcm1004uForm f,int i,String field){
        if(i<0||i>=f.getTabs().size())throw new IllegalStateException(EXPIRED);
        var t=f.getTabs().get(i);
        if("start".equals(field)&&i>0){
            var prev=f.getTabs().get(i-1);
            if(t.getKaisiDt()==null||prev.getKaisiDt()==null||prev.getSyuryoDt()==null||!t.getKaisiDt().isAfter(prev.getKaisiDt())||t.getKaisiDt().isAfter(prev.getSyuryoDt().plusDays(1)))throw new IllegalStateException("開始日は前の期間内から、前の終了日の翌日までで指定してください。");
            prev.setSyuryoDt(t.getKaisiDt().minusDays(1));
        }else if("end".equals(field)&&i<f.getTabs().size()-1){
            var next=f.getTabs().get(i+1);
            if(t.getSyuryoDt()==null)throw new IllegalStateException("最終期間以外の終了日は必須です。");
            if(next.getSyuryoDt()!=null&&(t.getSyuryoDt().isBefore(next.getKaisiDt())||!t.getSyuryoDt().isBefore(next.getSyuryoDt())))throw new IllegalStateException("終了日は次の期間の開始日から、終了日の前日までで指定してください。");
            next.setKaisiDt(t.getSyuryoDt().plusDays(1));
        }else if(!Set.of("start","end").contains(field))throw new IllegalStateException(EXPIRED);
        var errors=validatePeriods(f.getTabs());if(!errors.isEmpty())throw new IllegalStateException(String.join(" ",errors));
    }

    /** The caller passes a server-side draft; posted IDs, totals and edit permissions are never bound. */
    @Transactional(isolation=Isolation.SERIALIZABLE)
    public void updateAll(Mcm1004uForm draft,String user){
        if(!canUpdate(user)||draft.isLockReleaseFlg()&&!canUnlock(user))throw new IllegalStateException("更新権限がありません。");
        jdbc.queryForList("SELECT TM_KEIYAKUJIKAN_ID FROM MCM.MCM_TM_KEIYAKUJIKAN WITH (UPDLOCK,HOLDLOCK) WHERE TM_KEIYAKUJIKAN_ID=?",draft.getTmKeiyakujikanId());
        var current=loadScreen(draft.getTmKeiyakujikanId(),draft.getSeniMotoKbn(),draft.isLockReleaseFlg());
        if(!Objects.equals(current.getRevision(),draft.getRevision()))throw new IllegalStateException("他の操作で見積が更新されました。再読み込みして内容を確認してください。");
        if(!current.isViewFlg())throw new IllegalStateException("この見積は参照専用です。");
        var errors=validatePeriods(draft.getTabs());if(!errors.isEmpty())throw new IllegalStateException(String.join(" ",errors));
        var seen=new HashSet<Long>();for(var t:draft.getTabs()){
            Mcm1004uTabForm source;
            if(t.getTmKikanId()==null)throw new IllegalStateException(EXPIRED);
            if(t.getTmKikanId()==0)source=current.getTabs().get(current.getTabs().size()-1);else{if(!seen.add(t.getTmKikanId()))throw new IllegalStateException(EXPIRED);source=current.getTabs().stream().filter(p->Objects.equals(p.getTmKikanId(),t.getTmKikanId())).findFirst().orElseThrow(()->new IllegalStateException(EXPIRED));}
            validateRows(t,source,current);recalculate(t);
        }
        if(current.getTabs().stream().filter(t->t.getTmKikanId()!=0).anyMatch(t->!seen.contains(t.getTmKikanId())))throw new IllegalStateException(EXPIRED);
        Timestamp now=Timestamp.from(Instant.now());
        for(var t:draft.getTabs()){
            boolean fresh=t.getTmKikanId()==0;BigDecimal period=fresh?nextId("MCM_TM_KIKAN","TM_KIKAN_ID"):BigDecimal.valueOf(t.getTmKikanId());
            var p=new LinkedHashMap<String,Object>();p.put("TM_KEIYAKUJIKAN_ID",draft.getTmKeiyakujikanId());p.put("KAISI_DT",t.getKaisiDt());p.put("SYURYO_DT",t.getSyuryoDt());p.put("TM_MITSUMORI_NO",t.getTmMitsumoriNo());p.put("KAITO_DT",t.getKaitoDt());p.put("HYOJUNGOKEI_KIN",t.getHyojungokeiKin());p.put("SIKIRISYOKEI_KIN",t.getSikirisyokeiKin());p.put("SYUSSEINEBIKI_KIN",zero(t.getSyusseinebikiKin()));p.put("SIKIRIGOKEI_KIN",t.getSikirigokeiKin());p.put("BIKO",t.getBiko());persist("MCM_TM_KIKAN","TM_KIKAN_ID",period,fresh,p,user,now);
            for(var r:t.getTankaRows()){
                var values=new LinkedHashMap<String,Object>();values.put("TM_KIKIMEISAI_ID",r.getTmKikimeisaiId());values.put("TM_KIKAN_ID",period);values.put("HYOJUN_KIN",zero(r.getHyojunKin()));values.put("SIKIRI_KIN",zero(r.getSikiriKin()));values.put("PACK_FLG",flag(r.getPackFlg()));values.put("KEIYAKUNAIYO",r.getKeiyakunaiyo());values.put("KEIYAKU_NO",r.getKeiyakuNo());values.put("TORIHOSYUJIKAN_ID",r.getTorihosyujikanId());values.put("TENKENUMU",flag(r.getTenkenumu()));values.put("HOSYUHOHO",r.getHosyuhoho());values.put("SERVICEKEITAI",r.getServicekeitai());values.put("BIKO",r.getBiko());boolean insert=fresh||r.getTmTankaId()==null;persist("MCM_TM_TANKA","TM_TANKA_ID",insert?nextId("MCM_TM_TANKA","TM_TANKA_ID"):r.getTmTankaId(),insert,values,user,now);
            }
            for(var r:t.getTenkenRows()){
                var values=new LinkedHashMap<String,Object>();values.put("TM_KIKIKOSEI_ID",r.getTmKikikoseiId());values.put("TM_KIKAN_ID",period);values.put("OYAKIKIBUNRUI_CD",r.getOyakikibunruiCd());values.put("TENKENKAISU",zero(r.getTenkenkaisu()));values.put("TENKENKANOYOBI",r.getTenkenkanoyobi());values.put("YAKANTAIOUMU",flag(r.getYakantaioumu()));values.put("BIKO",r.getBiko());boolean insert=fresh||r.getTmTenkenId()==null;persist("MCM_TM_TENKEN","TM_TENKEN_ID",insert?nextId("MCM_TM_TENKEN","TM_TENKEN_ID"):r.getTmTenkenId(),insert,values,user,now);
            }
        }
        String status=Set.of("0","9").contains(current.getJotai())?"1":current.getJotai();
        if(jdbc.update("UPDATE MCM.MCM_TM_KEIYAKUJIKAN SET JOTAI=?,LASTUPDATE_DT=?,LASTUPDATE_BY=? WHERE TM_KEIYAKUJIKAN_ID=?",status,now,user,draft.getTmKeiyakujikanId())!=1)throw new IllegalStateException(EXPIRED);
    }
    private void validateRows(Mcm1004uTabForm t,Mcm1004uTabForm source,Mcm1004uForm f){
        length(t.getTmMitsumoriNo(),20);length(t.getBiko(),4000);amount(t.getSyusseinebikiKin());if(t.getKaitoDt()!=null&&(t.getKaitoDt().getYear()<1||t.getKaitoDt().getYear()>9999))throw new IllegalStateException("回答日の年が範囲外です。");
        if(t.getTankaRows().size()!=source.getTankaRows().size()||t.getTenkenRows().size()!=source.getTenkenRows().size())throw new IllegalStateException(EXPIRED);
        var keys=new HashSet<String>();for(var r:t.getTankaRows()){
            if(!keys.add(str(r.getTmKikimeisaiId())))throw new IllegalStateException(EXPIRED);var old=source.getTankaRows().stream().filter(x->same(x.getTmKikimeisaiId(),r.getTmKikimeisaiId())).findFirst().orElseThrow(()->new IllegalStateException(EXPIRED));
            if(t.getTmKikanId()!=0&&!Objects.equals(old.getTmTankaId(),r.getTmTankaId()))throw new IllegalStateException(EXPIRED);r.setSuryoNm(old.getSuryoNm());amount(r.getHyojunKin());amount(r.getSikiriKin());flag(r.getPackFlg());flag(r.getTenkenumu());length(r.getKeiyakunaiyo(),400);length(r.getBiko(),4000);choice(r.getHosyuhoho(),old.getHosyuhoho(),METHODS);choice(r.getServicekeitai(),old.getServicekeitai(),SERVICES);choice(str(r.getTorihosyujikanId()),str(old.getTorihosyujikanId()),f.getHoshuOptions());
        }
        keys.clear();for(var r:t.getTenkenRows()){if(!keys.add(tenkenKey(r)))throw new IllegalStateException(EXPIRED);var old=source.getTenkenRows().stream().filter(x->tenkenKey(x).equals(tenkenKey(r))).findFirst().orElseThrow(()->new IllegalStateException(EXPIRED));if(t.getTmKikanId()!=0&&!Objects.equals(old.getTmTenkenId(),r.getTmTenkenId()))throw new IllegalStateException(EXPIRED);amount(r.getTenkenkaisu());flag(r.getYakantaioumu());choice(r.getTenkenkanoyobi(),old.getTenkenkanoyobi(),DAYS);length(r.getBiko(),4000);}
    }
    private static void choice(String v,String old,Map<String,String> choices){if(v!=null&&!v.isBlank()&&!choices.containsKey(v)&&!Objects.equals(v,old))throw new IllegalStateException("選択項目の値が正しくありません。");}
    private static void length(String s,int n){if(s!=null&&s.length()>n)throw new IllegalStateException("入力文字数が上限を超えています（"+n+"文字）。");}
    private static void amount(BigDecimal n){if(n!=null&&(n.signum()<0||n.stripTrailingZeros().scale()>6||n.abs().compareTo(BigDecimal.TEN.pow(22))>=0))throw new IllegalStateException("金額・回数は0以上、小数6桁以内で入力してください。");}
    private static String flag(String v){if(v==null||v.isBlank())return "0";if(!Set.of("0","1").contains(v))throw new IllegalStateException("チェック項目の値が正しくありません。");return v;}
    private BigDecimal nextId(String table,String id){return jdbc.queryForObject("SELECT COALESCE(MAX("+id+"),0)+1 FROM MCM."+table+" WITH (UPDLOCK,HOLDLOCK)",BigDecimal.class);}
    private void persist(String table,String pk,BigDecimal id,boolean insert,LinkedHashMap<String,Object> values,String user,Timestamp now){
        values.put("LASTUPDATE_DT",now);values.put("LASTUPDATE_BY",user);
        if(insert){values.put(pk,id);values.put("CREATED_DT",now);values.put("CREATED_BY",user);String sql="INSERT INTO MCM."+table+" ("+String.join(",",values.keySet())+") VALUES ("+String.join(",",Collections.nCopies(values.size(),"?"))+")";if(jdbc.update(sql,values.values().toArray())!=1)throw new IllegalStateException(EXPIRED);}
        else{var args=new ArrayList<Object>(values.values());args.add(id);if(jdbc.update("UPDATE MCM."+table+" SET "+String.join(",",values.keySet().stream().map(k->k+"=?").toList())+" WHERE "+pk+"=?",args.toArray())!=1)throw new IllegalStateException(EXPIRED);}
    }
}
