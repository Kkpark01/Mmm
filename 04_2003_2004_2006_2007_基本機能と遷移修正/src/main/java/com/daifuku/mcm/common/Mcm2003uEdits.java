package com.daifuku.mcm.common;

import java.math.*;
import java.util.*;
import jakarta.servlet.http.HttpServletRequest;
import org.springframework.beans.BeanWrapperImpl;
import com.daifuku.mcm.form.Mcm2003uForm;
import com.daifuku.mcm.form.Mcm2003uForm.BrandRowForm;
import com.daifuku.mcm.form.Mcm2003uForm.KoseiRowForm;
import com.daifuku.mcm.service.Mcm2003uDraftService;
import static com.daifuku.mcm.common.CustomerScreenSupport.same;

/** VB の Leave 順序を再生し、変更列だけを保存する。変更前値はサーバーのセッションから採取する。 */
public final class Mcm2003uEdits {
    private Mcm2003uEdits() {}
    public static final Map<String,String> BRAND = Map.ofEntries(
        Map.entry("keiyakujikantai","KEIYAKUJIKANTAI"),Map.entry("softFlg","SOFT_FLG"),
        Map.entry("dremosFlg","DREMOS_FLG"),Map.entry("remoteFlg","REMOTE_FLG"),
        Map.entry("systemSupportKin","SYSTEMSUPPORT_KIN"),Map.entry("dtsSupportKin","DTSSUPPORT_KIN"),
        Map.entry("softHosyuKin","SOFTHOSHU_KIN"),Map.entry("hoseiSystemSupportKin","HOSEISYSTEMSUPPORT_KIN"),
        Map.entry("hoseiDtsSupportKin","HOSEIDTSSUPPORT_KIN"),Map.entry("hoseiDaifukuGijutsuKin","HOSEIDAIFUKUGIJUTSU_KIN"),
        Map.entry("hoseiSoftHosyuKin","HOSEISOFTHOSHU_KIN"));
    public static final Map<String,String> KOSEI = Map.of("hosyuhoho","HOSYUHOHO","biko","BIKO");
    public static Map<String,Object> originals(Object row) {
        return row instanceof BrandRowForm b?b.getEditedOriginals():((KoseiRowForm)row).getEditedOriginals();
    }
    public static Object get(Object row,String key) { return new BeanWrapperImpl(row).getPropertyValue(key); }
    public static boolean equal(Object a,Object b) {
        return a instanceof BigDecimal x && b instanceof BigDecimal y ? x.compareTo(y)==0 : Objects.equals(a,b);
    }
    public static void change(Object row,String key,Object value) {
        if(!(row instanceof BrandRowForm?BRAND:KOSEI).containsKey(key))throw new IllegalStateException("編集項目が不正です。");
        var bean=new BeanWrapperImpl(row);Object old=bean.getPropertyValue(key);
        if(equal(old,value))return;
        var originals=originals(row);if(!originals.containsKey(key))originals.put(key,old);
        bean.setPropertyValue(key,value);
    }
    public static void overlay(Object old,Object target) {
        for(var e:originals(old).entrySet()){new BeanWrapperImpl(target).setPropertyValue(e.getKey(),get(old,e.getKey()));originals(target).put(e.getKey(),e.getValue());}
    }
    public static Object databaseValue(Object v) { return v instanceof Boolean b?(b?BigDecimal.ONE:BigDecimal.ZERO):v; }
    public static void draftValues(Object row,Map<String,Object> target) {
        var columns=row instanceof BrandRowForm?BRAND:KOSEI;
        for(String key:originals(row).keySet())target.put(columns.get(key),databaseValue(get(row,key)));
    }
    private static BigDecimal number(String value) {
        if(value==null||!value.matches("-?[0-9]{1,12}(\\.[0-9]{1,6})?"))throw new IllegalStateException("金額は12桁以内の数値で入力してください。");
        return new BigDecimal(value);
    }
    private static BigDecimal n(BigDecimal v){return v==null?BigDecimal.ZERO:v;}
    public static void calculate(BrandRowForm b,String key,String value) {
        if(key.equals("keiyakujikantai")) {
            if(!value.isEmpty()&&(!value.matches("[0-9]{1,2}")||new BigDecimal(value).compareTo(new BigDecimal("24"))>0))throw new IllegalStateException("契約時間帯は0～24で入力してください。");
            if(Objects.equals(value,b.getKeiyakujikantai()))return;
            change(b,key,value);
            var system=value.isEmpty()?BigDecimal.ZERO:Mcm2003uDraftService.systemSupport(n(b.getSystemSekkeiKin()).add(n(b.getKihonSekkeiKin())).add(n(b.getProgramSakuseiKin())),value);
            var dts=system.multiply(new BigDecimal("0.25"));change(b,"systemSupportKin",system);change(b,"dtsSupportKin",dts);change(b,"softHosyuKin",system.add(dts).add(n(b.getDaifukuGijutsuKin())));return;
        }
        if(!List.of("hoseiSystemSupportKin","hoseiDaifukuGijutsuKin","hoseiSoftHosyuKin").contains(key))throw new IllegalStateException("金額の編集項目が不正です。");
        change(b,key,number(value.isEmpty()?"0":value));
        if(key.equals("hoseiSystemSupportKin")) {
            var dts=n(b.getHoseiSystemSupportKin()).multiply(new BigDecimal("0.25")).setScale(0,RoundingMode.HALF_EVEN);
            change(b,"hoseiDtsSupportKin",dts);change(b,"hoseiSoftHosyuKin",n(b.getHoseiSystemSupportKin()).add(dts).add(n(b.getHoseiDaifukuGijutsuKin())));
        } else {
            var difference=n(b.getHoseiSoftHosyuKin()).subtract(n(b.getHoseiDaifukuGijutsuKin()));
            if(difference.signum()>0){var system=difference.divideToIntegralValue(new BigDecimal("1.25"));change(b,"hoseiSystemSupportKin",system);change(b,"hoseiDtsSupportKin",difference.subtract(system));}
        }
    }
    public static void bind(HttpServletRequest request,Mcm2003uForm form) {
        if(request.getParameter("brandEditPresent")==null)return;
        BigDecimal id;try{id=new BigDecimal(request.getParameter("editedBrandId"));}catch(Exception ex){throw new IllegalStateException("ブランドを選択し直してください。");}
        var b=form.getBrandRows().stream().filter(r->same(r.getUmKihonBrandId(),id)).findFirst().orElseThrow(()->new IllegalStateException("ブランドを選択し直してください。"));
        String[] events=request.getParameterValues("brandCalculation");
        if(events!=null){if(events.length>2000)throw new IllegalStateException("一度登録してから編集を続けてください。");for(String event:events){String[] pair=event.split("=",-1);if(pair.length!=2)throw new IllegalStateException("計算内容が不正です。");calculate(b,pair[0],pair[1]);}}
        // JS が無効な場合も、実際に変更された値だけを VB の入力順で処理する。
        for(String key:List.of("keiyakujikantai","hoseiSystemSupportKin","hoseiDaifukuGijutsuKin","hoseiSoftHosyuKin")) {
            String value=request.getParameter("brand."+key);if(value==null)continue;
            Object parsed=key.equals("keiyakujikantai")?value:number(value.isEmpty()?"0":value);
            if(!equal(parsed,get(b,key)))calculate(b,key,value);
        }
        for(String key:List.of("softFlg","dremosFlg","remoteFlg"))change(b,key,"true".equals(request.getParameter("brand."+key)));
        for(var k:form.getKoseiRows())if(same(k.getUmKihonBrandId(),id)) {
            String prefix="kosei."+k.getUmKikikoseiId().stripTrailingZeros().toPlainString()+".";
            for(String key:KOSEI.keySet()) {
                String value=request.getParameter(prefix+key);if(value==null)continue;
                if(key.equals("hosyuhoho")&&!List.of("","F","S","C","H","I","T").contains(value))throw new IllegalStateException("保守方法を選択してください。");
                if(key.equals("biko")&&value.length()>4000)throw new IllegalStateException("機器構成の備考は4000文字以内で入力してください。");
                if(!Objects.equals(value,get(k,key)==null?"":get(k,key)))change(k,key,value);
            }
        }
    }
}
