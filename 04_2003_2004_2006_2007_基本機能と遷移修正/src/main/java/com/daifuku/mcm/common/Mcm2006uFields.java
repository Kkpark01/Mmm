package com.daifuku.mcm.common;

import java.math.BigDecimal;
import java.util.*;
import jakarta.servlet.http.HttpServletRequest;
import org.springframework.beans.BeanWrapperImpl;
import org.springframework.jdbc.core.JdbcTemplate;
import com.daifuku.mcm.form.Mcm2006uForm;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;

/** VBの表示列を保持し、編集可能列だけを保存する。ID・確認情報・更新情報は入力対象外。 */
public final class Mcm2006uFields {
    private Mcm2006uFields(){}
    public static final List<String> HEADER=List.of("KEIYAKU_NO","KEIYAKU_DT","SHOKAI_KEIYAKU_DT","KEIYAKUMANRYO_DT","ENTYOKEIYAKUMANRYO_DT","KAIYAKU_DT","AUTO_FLG","BIKO");
    public static final List<String> BRAND=List.of("SOFT_FLG","DREMOS_FLG","DREMOS_NM","REMOTE_FLG","REMOTERENRAKUSAKI","BIKO");
    public static final List<String> DETAIL=List.of("PACK_FLG","KEIYAKUNAIYO","KEIYAKU_NO","DAIFUKUHOSYUJIKAN_ID","TENKENKAISU","TENKENYOBI","HOSYUHOHO","BIKO");
    public static final List<String> SEIBAN=List.of("KAISI_DT","SYURYO_DT","HARD_SEIBAN","SOFT_SEIBAN","BIKO");
    public static final List<String> INSPECTION;
    static {var list=new ArrayList<>(List.of("NAIYO","BIKO"));for(int i=1;i<=12;i++)list.add(String.format("M%02d",i));INSPECTION=List.copyOf(list);}
    @SuppressWarnings("unchecked") public static Map<String,String> fields(Object bean){return (Map<String,String>)new BeanWrapperImpl(bean).getPropertyValue("extra");}
    private static void edited(Object bean){new BeanWrapperImpl(bean).setPropertyValue("extraEdited",true);}
    private static boolean changed(Object bean){return Boolean.TRUE.equals(new BeanWrapperImpl(bean).getPropertyValue("extraEdited"));}
    private static String value(Object x){
        if(x==null)return "";if(x instanceof BigDecimal n)return n.stripTrailingZeros().toPlainString();
        if(x instanceof java.sql.Timestamp t)return t.toLocalDateTime().format(java.time.format.DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss"));
        return x.toString();
    }
    public static String numberText(String text){if(text==null||text.isBlank())return "";try{return new BigDecimal(text).stripTrailingZeros().toPlainString();}catch(NumberFormatException e){return text;}}
    private static void read(JdbcTemplate jdbc,Object bean,String table,String key,BigDecimal id){
        var rows=jdbc.queryForList("SELECT * FROM MCM."+table+" WHERE "+key+"=?",id);if(rows.size()!=1)throw new IllegalStateException(EXPIRED);
        rows.get(0).forEach((k,v)->fields(bean).put(k,k.endsWith("_DT")&&!k.equals("CREATED_DT")&&!k.equals("LASTUPDATE_DT")&&v instanceof java.sql.Timestamp t?format(t.toLocalDateTime().toLocalDate()):value(v)));
    }
    public static void load(JdbcTemplate jdbc,Mcm2006uForm f){
        read(jdbc,f,"MCM_UK_KEIYAKU","UK_KEIYAKU_ID",f.getUkKeiyakuId());
        for(var row:f.getSeibanRows())read(jdbc,row,"MCM_UK_SEIBAN","UK_SEIBAN_ID",row.getUkSeibanId());
        for(var t:f.getKikanTabs()){
            read(jdbc,t,"MCM_UK_KIKAN","UK_KIKAN_ID",t.getUkKikanId());
            for(var r:t.getBrandRows())read(jdbc,r,"MCM_UK_BRAND","UK_BRAND_ID",r.getUkBrandId());
            for(var r:t.getKoseiRows())read(jdbc,r,"MCM_UK_KIKIKOSEI","UK_KIKIKOSEI_ID",r.getUkKikoseiId());
            for(var r:t.getMeisaiRows())read(jdbc,r,"MCM_UK_KIKIMEISAI","UK_KIKIMEISAI_ID",r.getUkKikimeisaiId());
            for(var r:t.getTenkenRows())read(jdbc,r,"MCM_UK_TENKEN","UK_TENKEN_ID",r.getUkTenkenId());
        }
    }
    /** 2007Uで選んだ見積の契約条件を新しい期間へ引き継ぐ。 */
    public static void quoteFields(JdbcTemplate jdbc,Mcm2006uForm.KikanTabForm tab){
        for(var brand:tab.getBrandRows()){
            var sources=jdbc.queryForList("SELECT * FROM MCM.MCM_UM_KIHON_BRAND WHERE UM_KIHON_BRAND_ID=?",brand.getUmKihonBrandId());
            if(sources.size()!=1)throw new IllegalStateException(EXPIRED);
            for(String key:BRAND)if(sources.get(0).containsKey(key))fields(brand).put(key,value(sources.get(0).get(key)));
            edited(brand);
            for(var detail:tab.getMeisaiRows())if(same(brand.getBrandkoseiId(),detail.getBrandkoseiId())){
                var terms=jdbc.queryForList("SELECT T.PACK_FLG,T.BIKO FROM MCM.MCM_UM_TANKA T JOIN MCM.MCM_UM_KIKIMEISAI D ON D.UM_KIKIMEISAI_ID=T.UM_KIKIMEISAI_ID JOIN MCM.MCM_UM_KIKIKOSEI C ON C.UM_KIKIKOSEI_ID=D.UM_KIKIKOSEI_ID WHERE C.UM_KIHON_BRAND_ID=? AND T.UM_MITSUMORI_ID=? AND D.KIKIMEISAI_ID=?",brand.getUmKihonBrandId(),brand.getUmMitsumoriId(),detail.getKikimeisaiId());
                if(terms.size()>1)throw new IllegalStateException("見積明細が重複しています。見積を確認してください。");
                if(!terms.isEmpty()){terms.get(0).forEach((k,v)->fields(detail).put(k,value(v)));edited(detail);}
            }
        }
    }
    public static void initialize(Mcm2006uForm f){
        if(f.getUkKeiyakuId()==null&&!f.getKikanTabs().isEmpty()){
            String first=f.getKikanTabs().get(0).getKaisiDt();
            if(f.getKeiyakuDt()==null)f.setKeiyakuDt(first);
            if(f.getShokaiKeiyakuDt()==null)f.setShokaiKeiyakuDt(first);
            if(f.getJotai()==null)f.setJotai("2");
        }
        var header=fields(f);header.putIfAbsent("KEIYAKU_DT",value(f.getKeiyakuDt()));header.putIfAbsent("SHOKAI_KEIYAKU_DT",value(f.getShokaiKeiyakuDt()));header.putIfAbsent("KEIYAKUMANRYO_DT",value(f.getKeiyakumanryoDt()));header.putIfAbsent("ENTYOKEIYAKUMANRYO_DT",value(f.getEntyokeiyakumanryoDt()));header.putIfAbsent("KAIYAKU_DT",value(f.getKaiyakuDt()));
        fields(f).putIfAbsent("AUTO_FLG",f.getAutoFlg()==null?"1":f.getAutoFlg());
        for(var t:f.getKikanTabs()){
            fields(t).putIfAbsent("IRAITENPO_ID",value(t.getIraitenpoId()));
            for(var r:t.getMeisaiRows()){
                var e=fields(r);e.putIfAbsent("PACK_FLG","0");e.putIfAbsent("KEIYAKUNAIYO",value(r.getKeiyakunaiyo()));e.putIfAbsent("KEIYAKU_NO",value(r.getKeiyakuNo()));
                e.putIfAbsent("DAIFUKUHOSYUJIKAN_ID",value(r.getDaifukuhosyujikanId()));e.putIfAbsent("TENKENKAISU",value(r.getTenkenkaisu()));e.putIfAbsent("TENKENYOBI",value(r.getTenkenyobi()));e.putIfAbsent("HOSYUHOHO",value(r.getHosyuhoho()));
            }
        }
    }
    private static void bindRow(HttpServletRequest request,Object row,String prefix,List<String> allowed){
        for(String key:allowed){String input=request.getParameter(prefix+key);if(input!=null&&prefix.startsWith("seiban.")&&key.endsWith("_DT")&&!input.isBlank()){var day=date(input.length()==7?input+"/01":input);input=format(key.equals("KAISI_DT")?day.withDayOfMonth(1):day.withDayOfMonth(day.lengthOfMonth()));}if(input!=null&&!Objects.equals(input,fields(row).get(key))){fields(row).put(key,input);edited(row);}}
    }
    public static void bind(HttpServletRequest request,Mcm2006uForm f,boolean tabEditable){
        bindRow(request,f,"contract.",HEADER);
        for(int i=0;i<f.getSeibanRows().size();i++)if(!f.getSeibanRows().get(i).isRemoved()){
            var r=f.getSeibanRows().get(i);bindRow(request,r,"seiban."+i+".",SEIBAN);
            if(changed(r)){r.setKaisiDt(fields(r).get("KAISI_DT"));r.setSyuryoDt(fields(r).get("SYURYO_DT"));}
        }
        var t=f.getSelectedTab();if(t!=null&&tabEditable){
            for(int i=0;i<t.getBrandRows().size();i++)bindRow(request,t.getBrandRows().get(i),"brand."+i+".",BRAND);
            for(int i=0;i<t.getMeisaiRows().size();i++)bindRow(request,t.getMeisaiRows().get(i),"detail."+i+".",DETAIL);
            for(int i=0;i<t.getTenkenRows().size();i++)if(!t.getTenkenRows().get(i).isRemoved())bindRow(request,t.getTenkenRows().get(i),"inspection."+i+".",INSPECTION);
        }
    }
    private static boolean numeric(String key){return key.endsWith("_FLG")||key.equals("DREMOS_NM")||key.equals("TENKENKAISU")||key.equals("DAIFUKUHOSYUJIKAN_ID");}
    private static Object checked(String key,String input){
        String v=input==null?"":input;
        int max=key.equals("KEIYAKU_NO")?50:key.endsWith("_SEIBAN")?20:key.equals("REMOTERENRAKUSAKI")?200:key.equals("KEIYAKUNAIYO")?400:4000;
        if(v.length()>max)throw new IllegalStateException("入力が長すぎます（"+key+"）。");
        if(key.matches("M\\d{2}")||key.endsWith("_FLG")){if(!v.isEmpty()&&!v.equals("0")&&!v.equals("1"))throw new IllegalStateException("チェック項目を確認してください。");return numeric(key)?(v.equals("1")?1:0):v;}
        if(key.endsWith("_DT"))return v.isBlank()?null:java.sql.Date.valueOf(date(v));
        if(numeric(key)){if(v.isBlank())return null;if(!v.matches("[0-9]{1,12}(\\.[0-9]{1,6})?"))throw new IllegalStateException("数値項目を確認してください。");return new BigDecimal(v);}
        if(key.equals("TENKENYOBI")&&!v.isBlank()&&!List.of("1","2","3","4").contains(v))throw new IllegalStateException("点検曜日を確認してください。");
        if(key.equals("HOSYUHOHO")&&!v.isBlank()&&!List.of("F","S","C","H","I","T").contains(v))throw new IllegalStateException("保守方法を確認してください。");return v;
    }
    private static void update(JdbcTemplate db,Object row,String table,String pk,BigDecimal id,String ownerKey,BigDecimal owner,List<String> columns,String user){
        if(!changed(row))return;var sets=new ArrayList<String>();var args=new ArrayList<Object>();
        for(String c:columns)if(fields(row).containsKey(c)){sets.add(c+"=?");args.add(checked(c,fields(row).get(c)));}
        if(sets.isEmpty())return;sets.add("LASTUPDATE_DT=GETDATE()");sets.add("LASTUPDATE_BY=?");args.add(user);args.add(id);args.add(owner);
        if(db.update("UPDATE MCM."+table+" SET "+String.join(",",sets)+" WHERE "+pk+"=? AND "+ownerKey+"=?",args.toArray())!=1)throw new IllegalStateException(EXPIRED);
    }
    public static void save(JdbcTemplate db,Mcm2006uForm f,String user){
        update(db,f,"MCM_UK_KEIYAKU","UK_KEIYAKU_ID",f.getUkKeiyakuId(),"UK_KEIYAKU_ID",f.getUkKeiyakuId(),HEADER,user);
        for(var r:f.getSeibanRows()){
            if(r.isRemoved()){if(r.getUkSeibanId()!=null)db.update("DELETE FROM MCM.MCM_UK_SEIBAN WHERE UK_SEIBAN_ID=? AND UK_KEIYAKU_ID=?",r.getUkSeibanId(),f.getUkKeiyakuId());continue;}
            if(changed(r)&&(!date(r.getSyuryoDt()).isAfter(date(r.getKaisiDt()))))throw new IllegalStateException("製番の終了日は開始日より後にしてください。");
            update(db,r,"MCM_UK_SEIBAN","UK_SEIBAN_ID",r.getUkSeibanId(),"UK_KEIYAKU_ID",f.getUkKeiyakuId(),SEIBAN,user);
        }
        for(var t:f.getKikanTabs()){
            for(var r:t.getBrandRows())update(db,r,"MCM_UK_BRAND","UK_BRAND_ID",r.getUkBrandId(),"UK_KIKAN_ID",t.getUkKikanId(),BRAND,user);
            for(var r:t.getMeisaiRows()){
                if(changed(r)&&!fields(r).getOrDefault("DAIFUKUHOSYUJIKAN_ID","").isBlank()){
                    var id=checked("DAIFUKUHOSYUJIKAN_ID",fields(r).get("DAIFUKUHOSYUJIKAN_ID"));
                    if(db.queryForObject("SELECT COUNT(*) FROM MCM.MCM_MA_TORIHOSYUJIKAN H JOIN MCM.MCM_MA_TORIHIKISAKI R ON R.TORIHIKISAKI_ID=H.TORIHIKISAKI_ID WHERE H.TORIHOSYUJIKAN_ID=? AND R.DAIFUKU_FLG=1",Integer.class,id)!=1)throw new IllegalStateException("保守時間を選択し直してください。");
                }
                update(db,r,"MCM_UK_KIKIMEISAI","UK_KIKIMEISAI_ID",r.getUkKikimeisaiId(),"UK_KIKIKOSEI_ID",r.getUkKikoseiId(),DETAIL,user);
            }
            for(var r:t.getTenkenRows()){
                if(r.isRemoved()){if(r.getUkTenkenId()!=null)db.update("DELETE FROM MCM.MCM_UK_TENKEN WHERE UK_TENKEN_ID=? AND UK_KIKAN_ID=?",r.getUkTenkenId(),t.getUkKikanId());continue;}
                if(r.getUkTenkenId()==null&&changed(r)){
                    r.setUkTenkenId(db.queryForObject("SELECT ISNULL(MAX(UK_TENKEN_ID),0)+1 FROM MCM.MCM_UK_TENKEN WITH (UPDLOCK,HOLDLOCK)",BigDecimal.class));
                    db.update("INSERT INTO MCM.MCM_UK_TENKEN (UK_TENKEN_ID,UK_KIKAN_ID,CREATED_DT,CREATED_BY,LASTUPDATE_DT,LASTUPDATE_BY) VALUES (?,?,GETDATE(),?,GETDATE(),?)",r.getUkTenkenId(),t.getUkKikanId(),user,user);
                }
                update(db,r,"MCM_UK_TENKEN","UK_TENKEN_ID",r.getUkTenkenId(),"UK_KIKAN_ID",t.getUkKikanId(),INSPECTION,user);
            }
        }
    }
}
