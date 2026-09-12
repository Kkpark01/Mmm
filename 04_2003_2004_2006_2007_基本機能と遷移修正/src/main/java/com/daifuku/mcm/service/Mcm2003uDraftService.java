package com.daifuku.mcm.service;

import java.math.*;
import java.sql.Timestamp;
import java.time.*;
import java.util.*;
import org.springframework.beans.BeanWrapperImpl;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import com.daifuku.mcm.dto.*;
import com.daifuku.mcm.form.*;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;

/** VB Mcm2002uScreen / Mcm2003uScreen.Search / Mcm2003uTabControl の見積生成。 */
@Service
public class Mcm2003uDraftService {
    static final String KH="MCM_UM_KIHON_MITSUMORI",PER="MCM_UM_MITSUMORI",BR="MCM_UM_KIHON_BRAND",
        KO="MCM_UM_KIKIKOSEI",ME="MCM_UM_KIKIMEISAI",IND="MCM_UM_KOTAIMEISAI",PRICE="MCM_UM_TANKA",BP="MCM_UM_BRAND";
    static final List<String> ORDER=List.of(KH,PER,BR,KO,ME,IND,PRICE,BP);
    static final Map<String,String> IDS=Map.of(KH,"UM_KIHON_MITSUMORI_ID",PER,"UM_MITSUMORI_ID",BR,"UM_KIHON_BRAND_ID",KO,"UM_KIKIKOSEI_ID",ME,"UM_KIKIMEISAI_ID",IND,"UM_KOTAIMEISAI_ID",PRICE,"UM_TANKA_ID");
    static final BigDecimal ZERO=BigDecimal.ZERO,ONE=BigDecimal.ONE;
    static final LocalDate LIMIT=LocalDate.of(9999,12,31);
    @org.springframework.beans.factory.annotation.Autowired private McmCustomerTotalsService totals;
    @org.springframework.beans.factory.annotation.Autowired private McmCustomerIntegrityService integrity;
    private final JdbcTemplate jdbc;
    private final Mcm2003uService estimates;
    public Mcm2003uDraftService(JdbcTemplate jdbc,Mcm2003uService estimates){this.jdbc=jdbc;this.estimates=estimates;}
    public String sourceHours(BigDecimal id){return text(one("SELECT KEIYAKUJIKANTAI FROM MCM."+KH+" WHERE UM_KIHON_MITSUMORI_ID=?",id),"KEIYAKUJIKANTAI");}

    @Transactional(readOnly=true)
    public McmEstimateDraft prepare(Mcm2002uDeliveryDto delivery,Mcm2002uService.SearchResult selection,Mcm2002uForm input,String userName){
        if(delivery==null||selection==null||selection.isReadOnly())throw new IllegalStateException(DENIED);
        var draft=new McmEstimateDraft();
        Map<String,Object> old=Map.of();
        if(delivery.getSeniMotoKbn()!=0){
            old=one("SELECT * FROM MCM."+KH+" WHERE UM_KIHON_MITSUMORI_ID=?",delivery.getUmKihonMitsumoriId());
            if(!same(num(old,"PLANT_ID"),delivery.getPlantId()))throw new IllegalStateException(EXPIRED);
            if(delivery.getSeniMotoKbn()==1){
                var f=estimates.load(delivery.getUmKihonMitsumoriId());
                if(estimates.isReadOnly(f.getSyouninJotai(),1)||!"1".equals(f.getJotai()))throw new IllegalStateException(DENIED);
                draft.setOriginalId(delivery.getUmKihonMitsumoriId());draft.setOriginalVersion(f.getLastupdateDt());
            }
        }
        var plant=one("SELECT P.*,N.NONYUSAKI_CD,N.NONYUSAKI_NK,N.JUSYO1_NK NONYUSAKIJUSYO1_NK,N.JUSYO2_NK NONYUSAKIJUSYO2_NK FROM MCM.MCM_MA_PLANT P JOIN MCM.MCM_MA_NONYUSAKI N ON N.NONYUSAKI_ID=P.NONYUSAKI_ID WHERE P.PLANT_ID=?",delivery.getPlantId());
        if(!same(num(plant,"NONYUSAKI_ID"),delivery.getNonyusakiId()))throw new IllegalStateException(EXPIRED);
        var header=row(draft,KH);
        if(!old.isEmpty())header.putAll(old);
        header.put("UM_KIHON_MITSUMORI_ID",BigDecimal.valueOf(-1));
        putFrom(header,plant,"NONYUSAKI_ID,NONYUSAKI_CD,NONYUSAKI_NK,NONYUSAKIJUSYO1_NK,NONYUSAKIJUSYO2_NK,PLANT_ID,SUPPORT_ID,PLANT_NK,NONYUBUSYO_NK,NONYUTANTOSYA_NK,NONYUTEL_NO,NONYUFAX_NO");
        header.put("JOTAI","1");header.put("SYOUNIN_JOTAI",draft.getOriginalId()==null?"0":old.get("SYOUNIN_JOTAI"));
        if(draft.getOriginalId()==null){header.put("UM_MITSUMORI_NO","");header.put("MITSUMORISAKUSEISYA_NK",userName);header.put("MITSUMORI_DT",Timestamp.valueOf(LocalDateTime.now()));
            for(String col:List.of("SHINSA_DT","SHINSA_BY","SYOUNIN_DT","SYOUNIN_BY"))header.put(col,null);}
        if(old.isEmpty()){header.put("MITSUMORIKIGEN","1ヶ月");header.put("MITSUMORILEVEL",ZERO);header.put("MITSUMORI_JOUKEN","　＊この金額はユーザ渡し標準価格（ハード保守費＋ソフトウェア保守費）です。\n　＊設備点検費用及びコントローラ配下の制御機器は本費用には含みません。");}
        List<BigDecimal> tmIds=selection.getTmKeiyakujikanIds()==null?List.of():selection.getTmKeiyakujikanIds();
        String hours=input.getKeiyakuJikantai();
        TreeSet<LocalDate> boundaries=new TreeSet<>();
        List<Map<String,Object>> tmPeriods=new ArrayList<>();
        for(BigDecimal id:tmIds){
            var source=one("SELECT K.KEIYAKUJIKANTAI,M.NONYUSAKI_ID,M.PLANT_ID FROM MCM.MCM_TM_KEIYAKUJIKAN K JOIN MCM.MCM_TM_MITSUMORI M ON M.TM_IRAI_ID=K.TM_IRAI_ID WHERE K.TM_KEIYAKUJIKAN_ID=?",id);
            if(!same(num(source,"PLANT_ID"),delivery.getPlantId())||!same(num(source,"NONYUSAKI_ID"),delivery.getNonyusakiId()))throw new IllegalStateException(EXPIRED);
            String h=text(source,"KEIYAKUJIKANTAI");if(hours==null||hours.isBlank())hours=h;
            if(!Objects.equals(hours,h))throw new IllegalStateException("契約時間帯が異なる見積は同時に選択できません。");
            tmPeriods.addAll(jdbc.queryForList("SELECT * FROM MCM.MCM_TM_KIKAN WHERE TM_KEIYAKUJIKAN_ID=? ORDER BY KAISI_DT",id));
        }
        if(tmIds.isEmpty()){boundaries.add(date(input.getKaisiDt()));boundaries.add(LIMIT);}
        else for(var period:tmPeriods){
            LocalDate start=day(period.get("KAISI_DT")),end=day(period.get("SYURYO_DT"));
            if(start==null)throw new IllegalStateException("取引先見積の契約開始日を確認してください。");
            boundaries.add(start);boundaries.add(end==null||!end.isBefore(LIMIT)?LIMIT:end.plusDays(1));
        }
        if(boundaries.size()<2)throw new IllegalStateException("見積の契約期間を確認してください。");
        header.put("KEIYAKUJIKANTAI",hours);
        var dates=new ArrayList<>(boundaries);
        for(int i=0;i<dates.size()-1;i++){
            var p=row(draft,PER);p.put("UM_KIHON_MITSUMORI_ID",num(header,"UM_KIHON_MITSUMORI_ID"));p.put("KAISI_DT",java.sql.Date.valueOf(dates.get(i)));
            p.put("SYURYO_DT",dates.get(i+1).equals(LIMIT)?null:java.sql.Date.valueOf(dates.get(i+1).minusDays(1)));
            LocalDate periodStart=dates.get(i);
            p.put("SYUSSEINEBIKI_KIN",tmPeriods.stream().filter(r->covers(r,periodStart)).map(r->num(r,"SYUSSEINEBIKI_KIN")).reduce(ZERO,BigDecimal::add));
        }
        validateParents(selection);
        var masters=jdbc.queryForList("SELECT B.*,M.BRAND_NK,M.DREMOS_KIN,M.REMOTE_KIN FROM MCM.MCM_MA_BRAND_KOSEI B JOIN MCM.MCM_MA_BRAND M ON M.BRAND_ID=B.BRAND_ID WHERE B.PLANT_ID=? ORDER BY B.HYOJIJUN,B.BRANDKOSEI_ID",delivery.getPlantId());
        if(masters.isEmpty())throw new IllegalStateException("ブランド構成が登録されていません。ブランド構成を登録してから選定してください。");
        for(var master:masters){
            BigDecimal brandId=num(master,"BRANDKOSEI_ID");
            var allIndividuals=selection.getKotaiRows().stream().filter(k->same(k.getMadBrandkoseiId(),brandId)).toList();
            var chosen=allIndividuals.stream().filter(k->on(k.getCheckFlg())).toList();
            // VBは個体のないブランドもソフト保守対象として生成する。
            if(!allIndividuals.isEmpty()&&chosen.isEmpty())continue;
            var brand=row(draft,BR);brand.put("UM_KIHON_MITSUMORI_ID",num(header,"UM_KIHON_MITSUMORI_ID"));
            putFrom(brand,master,"BRAND_ID,BRAND_NK,BRANDKOSEI_ID,BRANDSYOSAI_NK,SYSTEMSEKKEI_KIN,KIHONSEKKEI_KIN,PROGRAMSAKUSEI_KIN");
            brand.put("KEIYAKUJIKANTAI",hours);brand.put("SOFT_FLG",ONE);brand.put("DREMOS_FLG",ONE);brand.put("REMOTE_FLG",on(num(master,"REMOTE_FLG"))?ZERO:ONE);brand.put("SOFTHOSYUHOHO","・TEL及びオンサイト対応");
            if(!old.isEmpty()){
                var oldBrands=jdbc.queryForList("SELECT * FROM MCM."+BR+" WHERE UM_KIHON_MITSUMORI_ID=? AND BRANDKOSEI_ID=?",delivery.getUmKihonMitsumoriId(),brandId);
                if(!oldBrands.isEmpty())putFrom(brand,oldBrands.get(0),"KEIYAKUJIKANTAI,SOFT_FLG,DREMOS_FLG,DREMOS_NM,REMOTE_FLG,SOFTHOSYUHOHO,HOSEISYSTEMSUPPORT_KIN,HOSEIDTSSUPPORT_KIN,HOSEIDAIFUKUGIJUTSU_KIN,HOSEISOFTHOSHU_KIN,BIKO");
            }
            for(var ko:selection.getKoseiRows()){
                var selectedKo=chosen.stream().filter(k->same(k.getMaeKikikoseiId(),ko.getKikikoseiId())).toList();
                if(selectedKo.isEmpty())continue;
                var k=row(draft,KO);k.put("UM_KIHON_BRAND_ID",num(brand,"UM_KIHON_BRAND_ID"));
                values(k,"KIKIKOSEI_ID,KIKIKOSEI_NK,TANI,TEHAISEIBAN,CONTROLLER_FLG,HYOJIJUN,BIKO",ko.getKikikoseiId(),ko.getKikikoseiNk(),ko.getTani(),ko.getTehaiseiban(),ko.getControllerFlg(),ko.getHyojijun(),ko.getBiko());
                if(!old.isEmpty()) {
                    var previous=jdbc.queryForList("SELECT K.HOSYUHOHO,K.BIKO FROM MCM."+KO+" K JOIN MCM."+BR+" B ON B.UM_KIHON_BRAND_ID=K.UM_KIHON_BRAND_ID WHERE B.UM_KIHON_MITSUMORI_ID=? AND B.BRANDKOSEI_ID=? AND K.KIKIKOSEI_ID=?",delivery.getUmKihonMitsumoriId(),brandId,ko.getKikikoseiId());
                    if(!previous.isEmpty())putFrom(k,previous.get(0),"HOSYUHOHO,BIKO");
                }
                var setCounts=new HashSet<Integer>();
                for(var me:selection.getMeisaiRows()){
                    if(!same(me.getKikikoseiId(),ko.getKikikoseiId()))continue;
                    var selectedMe=selectedKo.stream().filter(v->same(v.getMafKikimeisaiId(),me.getKikimeisaiId())).toList();
                    if(selectedMe.isEmpty())continue;
                    if(on(me.getKotaikanriFlg()))setCounts.add(selectedMe.size());
                    BigDecimal quantity=selectedMe.stream().map(v->nz(v.getMafSuryoNm())).reduce(ZERO,BigDecimal::add);
                    BigDecimal total=allIndividuals.stream().filter(v->same(v.getMafKikimeisaiId(),me.getKikimeisaiId())).map(v->nz(v.getMafSuryoNm())).reduce(ZERO,BigDecimal::add);
                    var m=row(draft,ME);m.put("UM_KIKIKOSEI_ID",num(k,"UM_KIKIKOSEI_ID"));
                    values(m,"KIKIMEISAI_ID,SEIZOMAKER_ID,SEIZOMAKER_NK,KIKIHINMEI_NK,KIKIKATASHIKI,SURYO_NM,SOSU_NM,ATSUKAIKIKI_ID,HYOJIJUN,BIKO",me.getKikimeisaiId(),me.getSeizomakerId(),me.getSeizomakerNk(),me.getKikihinmeiNk(),me.getKikikatashiki(),quantity,total,me.getAtsukaikikiId(),me.getHyojijun(),me.getBiko());
                    for(var individual:selectedMe){var ind=row(draft,IND);ind.put("UM_KIKIMEISAI_ID",num(m,"UM_KIKIMEISAI_ID"));
                        values(ind,"KOTAIKANRI_ID,KOTAI_NK,SERIAL_NO,ITIJINONYU_DT,SETCHIBASYO,KEIYAKUKIGEN_DT,ENCHOKEIYAKUKIGEN_DT",individual.getMagKotaikanriId(),individual.getMagKotaiNk(),individual.getMagSerialNo(),individual.getMagItijinonyuDt(),individual.getMagSetchibasyo(),individual.getMagKeiyakukigenDt(),individual.getMagEnchokeiyakukigenDt());}
                    for(var p:draft.rows(PER)){
                        LocalDate start=day(p.get("KAISI_DT"));
                        var prices=selection.getTankaRows().stream().filter(t->same(t.getKikimeisaiId(),me.getKikimeisaiId())&&same(t.getBrandkoseiId(),brandId)&&t.getKaisiDt()!=null&&!t.getKaisiDt().isAfter(start)&&(t.getSyuryoDt()==null||!t.getSyuryoDt().isBefore(start))).toList();
                        for(var source:prices){
                            BigDecimal q=selectedMe.stream().filter(v->same(v.getTmvTmKikimeisaiId(),source.getTmvTmKikimeisaiId())||tmIds.isEmpty()).map(v->nz(v.getMafSuryoNm())).reduce(ZERO,BigDecimal::add);
                            if(q.signum()==0)continue;
                            var t=price(draft,m,p,q,on(me.getControllerFlg()));
                            values(t,"TM_TANKA_ID,PACK_FLG,KEIYAKUNAIYO,KEIYAKU_NO,HYOJUN_KIN,SIKIRI_KIN,TORIHOSYUJIKAN_ID,TENKENKAISU,TENKENYOBI,HOSYUHOHO,SERVICEKEITAI",source.getTmvTmTankaId(),source.getPackFlg(),source.getKeiyakunaiyo(),source.getKeiyakuNo(),source.getHyojunKin(),source.getSikiriKin(),source.getTorihosyujikanId(),source.getTenkenkaisu(),source.getTenkenkanoyobi(),source.getHosyuhoho(),source.getServicekeitai());
                            var support=jdbc.queryForList("SELECT DAIFUKUHOSYUJIKAN_ID FROM MCM.MCM_MA_TORIHOSYUJIKAN WHERE TORIHOSYUJIKAN_ID=?",source.getTorihosyujikanId());
                            if(!support.isEmpty())t.put("DAIFUKUHOSYUJIKAN_ID",support.get(0).get("DAIFUKUHOSYUJIKAN_ID"));
                        }
                        if(prices.isEmpty()&&on(me.getControllerFlg())){
                            var rates=jdbc.queryForList("SELECT CONTROLLER_KIN FROM MCM.MCM_MA_ATSUKAIKIKI WHERE ATSUKAIKIKI_ID=?",me.getAtsukaikikiId());
                            if(rates.isEmpty())throw new IllegalStateException("コントローラ保守単価を確認してください。");
                            var t=price(draft,m,p,quantity,true);values(t,"TM_TANKA_ID,HYOJUN_KIN,SIKIRI_KIN,TENKENKAISU,TENKENYOBI,HOSYUHOHO,DAIFUKUHOSYUJIKAN_ID",ZERO,num(rates.get(0),"CONTROLLER_KIN"),ZERO,ONE,"1","F",BigDecimal.valueOf(4));
                        }
                    }
                }
                k.put("SET_NM",setCounts.size()==1?BigDecimal.valueOf(setCounts.iterator().next()):ONE);
            }
            calculateBrand(draft,brand,master,hours);
        }
        if(draft.rows(BR).isEmpty())throw new IllegalStateException("見積対象のブランド・機器を選択してください。");
        for(var p:draft.rows(PER)){
            var brands=draft.rows(BP).stream().filter(b->same(num(b,"UM_MITSUMORI_ID"),num(p,"UM_MITSUMORI_ID"))).toList();
            p.put("HARDHOSYU_KIN",brands.stream().map(b->num(b,"HARDHOSYU_KIN")).reduce(ZERO,BigDecimal::add));
            p.put("MITSUMORI_GKIN",brands.stream().map(b->num(b,"_TOTAL")).reduce(ZERO,BigDecimal::add));
        }
        draft.setForm(preview(draft,delivery.getSeniMotoKbn()));return draft;
    }

    /** 親子IDは登録時に採番する。失敗時は一式ロールバックし、画面の下書きは変更しない。 */
    @Transactional
    public BigDecimal save(McmEstimateDraft source,Mcm2003uForm input,Set<Integer> selectedPeriods,String user){
        if(selectedPeriods.isEmpty())throw new IllegalStateException("契約金額変動時期は、1件以上選択してください。");
        if(input.getUmMitsumoriNo()==null||input.getUmMitsumoriNo().isBlank())throw new IllegalStateException("見積NOを入力してください。");
        if(input.getUmMitsumoriNo().length()>20)throw new IllegalStateException("見積NOは20文字以内で入力してください。");
        if(input.getMitsumoriLevel()!=null&&!input.getMitsumoriLevel().isBlank()&&!input.getMitsumoriLevel().matches("[0-9]{1,2}(?:\\.0+)?"))throw new IllegalStateException("見積レベルは0～99の整数で入力してください。");
        if(!Objects.equals(input.getKeiyakujikantai(),text(source.rows(KH).get(0),"KEIYAKUJIKANTAI")))throw new IllegalStateException("契約時間帯は機器選定画面で確認してください。");
        var d=copy(source);
        for(var b:input.getBrandRows())if(b.isSoftHosyuhohoEdited()){
            if(b.getSoftHosyuhoho()!=null && b.getSoftHosyuhoho().length()>4000)throw new IllegalStateException("保守方法は4000文字以内で入力してください。");
            var row=d.rows(BR).stream().filter(r->same(num(r,"UM_KIHON_BRAND_ID"),b.getUmKihonBrandId())).findFirst().orElseThrow(()->new IllegalStateException(EXPIRED));row.put("SOFTHOSYUHOHO",b.getSoftHosyuhoho());
        }
        for(var b:input.getBrandRows())if(!b.getEditedOriginals().isEmpty())com.daifuku.mcm.common.Mcm2003uEdits.draftValues(b,d.rows(BR).stream().filter(r->same(num(r,"UM_KIHON_BRAND_ID"),b.getUmKihonBrandId())).findFirst().orElseThrow(()->new IllegalStateException(EXPIRED)));
        for(var k:input.getKoseiRows())if(!k.getEditedOriginals().isEmpty())com.daifuku.mcm.common.Mcm2003uEdits.draftValues(k,d.rows(KO).stream().filter(r->same(num(r,"UM_KIKIKOSEI_ID"),k.getUmKikikoseiId())).findFirst().orElseThrow(()->new IllegalStateException(EXPIRED)));
        if(d.getOriginalId()!=null){var original=estimates.load(d.getOriginalId());original.setLastupdateDt(d.getOriginalVersion());original.setSeniMotoKbn(1);estimates.requireEditable(original);}
        // 採番と見積NOの重複判定を同じトランザクションで直列化する。
        BigDecimal nextHeader=nextId(KH);
        BigDecimal headerId=d.getOriginalId()==null?nextHeader:d.getOriginalId();
        Integer duplicate=jdbc.queryForObject("SELECT COUNT(*) FROM MCM."+KH+" WITH (UPDLOCK,HOLDLOCK) WHERE UM_MITSUMORI_NO=? AND UM_KIHON_MITSUMORI_ID<>?",Integer.class,input.getUmMitsumoriNo().trim(),headerId);
        if(duplicate!=null&&duplicate>0)throw new IllegalStateException("同じ見積NOが登録されています。見積NOを確認してください。");
        var periods=new ArrayList<>(d.rows(PER));d.rows(PER).clear();
        for(int i=0;i<periods.size();i++)if(selectedPeriods.contains(i))d.rows(PER).add(periods.get(i));
        if(d.rows(PER).size()!=selectedPeriods.size())throw new IllegalStateException(EXPIRED);
        for(int i=0;i<d.rows(PER).size();i++)d.rows(PER).get(i).put("SYURYO_DT",i+1<d.rows(PER).size()?java.sql.Date.valueOf(day(d.rows(PER).get(i+1).get("KAISI_DT")).minusDays(1)):periods.get(periods.size()-1).get("SYURYO_DT"));
        var kept=d.rows(PER).stream().map(p->num(p,"UM_MITSUMORI_ID")).toList();
        d.rows(PRICE).removeIf(r->!kept.contains(num(r,"UM_MITSUMORI_ID")));d.rows(BP).removeIf(r->!kept.contains(num(r,"UM_MITSUMORI_ID")));
        var header=d.rows(KH).get(0);
        values(header,"UM_MITSUMORI_NO,KEIYAKUJIKANTAI,MITSUMORIKIGEN,MITSUMORILEVEL,IRAITANTO_NK,SOFUTANTO_NK,MITSUMORI_JOUKEN,HOSYUHOHO,BIKO",input.getUmMitsumoriNo().trim(),input.getKeiyakujikantai(),input.getMitsumorigiken(),(input.getMitsumoriLevel()==null||input.getMitsumoriLevel().isBlank()?ZERO:new BigDecimal(input.getMitsumoriLevel())),input.getIraitantoNk(),input.getSofutantoNk(),input.getMitsumoriJouken(),input.getHosyuhoho(),input.getBiko());
        setTenpo(header,"IRAI",input.getIraitenpoId());setTenpo(header,"SOFU",input.getSofutenpoId());
        var existing=d.getOriginalId()==null?new LinkedHashMap<String,List<Map<String,Object>>>():loadRows(headerId);
        Map<String,Map<BigDecimal,BigDecimal>> idMaps=new HashMap<>();
        for(String table:ORDER){
            String pk=IDS.get(table);Map<BigDecimal,BigDecimal> ids=new HashMap<>();if(pk!=null)idMaps.put(pk,ids);
            BigDecimal next=table.equals(KH)?nextHeader:pk==null?ZERO:nextId(table);
            var owned=existing.getOrDefault(table,List.of());
            for(var r:d.rows(table)){
                BigDecimal tempId=pk==null?null:num(r,pk);
                for(var foreign:idMaps.entrySet())if(!foreign.getKey().equals(pk)&&r.containsKey(foreign.getKey())){
                    var mapped=foreign.getValue().get(num(r,foreign.getKey()));if(mapped!=null)r.put(foreign.getKey(),mapped);
                }
                var matched=owned.stream().filter(oldRow->naturalKey(table,oldRow).equals(naturalKey(table,r))).findFirst().orElse(null);
                if(pk!=null){BigDecimal assigned=table.equals(KH)?headerId:matched==null?next:num(matched,pk);r.put(pk,assigned);ids.put(tempId,assigned);if(matched==null)next=next.add(ONE);}
                if(matched!=null){var generated=new LinkedHashMap<>(r);r.clear();r.putAll(matched);r.putAll(generated);
                    // 再選定で維持した単価行の手修正を保持する（VBは選定数量だけを更新）。
                    if(table.equals(PRICE))for(String col:List.of("HYOJUN_KIN","SIKIRI_KIN","PACK_FLG","KEIYAKUNAIYO","KEIYAKU_NO","TORIHOSYUJIKAN_ID","DAIFUKUHOSYUJIKAN_ID","TENKENKAISU","TENKENYOBI","HOSYUHOHO","SERVICEKEITAI","BIKO"))r.put(col,matched.get(col));
                }
                r.entrySet().removeIf(e->e.getKey().startsWith("_"));
                if(matched==null){r.put("CREATED_DT",Timestamp.valueOf(LocalDateTime.now()));r.put("CREATED_BY",user);}
                r.put("LASTUPDATE_DT",Timestamp.valueOf(LocalDateTime.now()));r.put("LASTUPDATE_BY",user);
                storeRow(table,r,matched!=null);
            }
        }
        // 保持行のIDは変えず、選定解除した子から順に削除する。
        var reverse=new ArrayList<>(ORDER);Collections.reverse(reverse);
        for(String table:reverse)for(var r:existing.getOrDefault(table,List.of())){
            if(d.rows(table).stream().noneMatch(keptRow->identity(table,keptRow).equals(identity(table,r)))){
                String where=IDS.containsKey(table)?IDS.get(table)+"=?":"UM_KIHON_BRAND_ID=? AND UM_MITSUMORI_ID=?";
                Object[] args=IDS.containsKey(table)?new Object[]{r.get(IDS.get(table))}:new Object[]{r.get("UM_KIHON_BRAND_ID"),r.get("UM_MITSUMORI_ID")};
                jdbc.update("DELETE FROM MCM."+table+" WHERE "+where,args);
            }
        }
        integrity.estimate(headerId);totals.recalculate(headerId,user);return headerId;
    }
    public void preserveBrandMethods(McmEstimateDraft prior,Mcm2003uForm pending,McmEstimateDraft next) {
        preserveEdits(prior,pending,next);
        for(var old:pending.getBrandRows())if(old.isSoftHosyuhohoEdited()) {
            var source=prior.rows(BR).stream().filter(r->same(num(r,"UM_KIHON_BRAND_ID"),old.getUmKihonBrandId())).findFirst().orElse(null);
            if(source==null)continue;
            for(var target:next.rows(BR))if(same(num(source,"BRANDKOSEI_ID"),num(target,"BRANDKOSEI_ID")))next.getForm().getBrandRows().stream().filter(b->same(b.getUmKihonBrandId(),num(target,"UM_KIHON_BRAND_ID"))).findFirst().ifPresent(b->{b.setSoftHosyuhoho(old.getSoftHosyuhoho());b.setSoftHosyuhohoEdited(true);});
        }
    }
    private void preserveEdits(McmEstimateDraft prior,Mcm2003uForm pending,McmEstimateDraft next) {
        for(var old:pending.getBrandRows()) {
            var source=prior.rows(BR).stream().filter(r->same(num(r,"UM_KIHON_BRAND_ID"),old.getUmKihonBrandId())).findFirst().orElse(null);if(source==null)continue;
            for(var target:next.rows(BR))if(same(num(source,"BRANDKOSEI_ID"),num(target,"BRANDKOSEI_ID"))) {
                next.getForm().getBrandRows().stream().filter(b->same(b.getUmKihonBrandId(),num(target,"UM_KIHON_BRAND_ID"))).findFirst().ifPresent(b->com.daifuku.mcm.common.Mcm2003uEdits.overlay(old,b));
                for(var k:pending.getKoseiRows())if(same(k.getUmKihonBrandId(),old.getUmKihonBrandId())) {
                    var sourceKo=prior.rows(KO).stream().filter(r->same(num(r,"UM_KIKIKOSEI_ID"),k.getUmKikikoseiId())).findFirst().orElse(null);if(sourceKo==null)continue;
                    for(var targetKo:next.rows(KO))if(same(num(targetKo,"UM_KIHON_BRAND_ID"),num(target,"UM_KIHON_BRAND_ID"))&&same(num(targetKo,"KIKIKOSEI_ID"),num(sourceKo,"KIKIKOSEI_ID")))next.getForm().getKoseiRows().stream().filter(r->same(r.getUmKikikoseiId(),num(targetKo,"UM_KIKIKOSEI_ID"))).findFirst().ifPresent(r->com.daifuku.mcm.common.Mcm2003uEdits.overlay(k,r));
                }
            }
        }
    }
    private BigDecimal nextId(String table){return jdbc.queryForObject("SELECT COALESCE(MAX("+IDS.get(table)+"),0)+1 FROM MCM."+table+" WITH (UPDLOCK,HOLDLOCK)",BigDecimal.class);}
    private void storeRow(String table,Map<String,Object>r,boolean update){
        String pk=IDS.get(table);var keys=pk==null?List.of("UM_KIHON_BRAND_ID","UM_MITSUMORI_ID"):List.of(pk);
        var columns=r.keySet().stream().filter(c->!update||!keys.contains(c)).toList();
        var args=new ArrayList<Object>();for(String col:columns){Object v=r.get(col);args.add(v instanceof LocalDate date?java.sql.Date.valueOf(date):v);}
        if(update){for(String k:keys)args.add(r.get(k));jdbc.update("UPDATE MCM."+table+" SET "+String.join(",",columns.stream().map(c->c+"=?").toList())+" WHERE "+String.join(" AND ",keys.stream().map(k->k+"=?").toList()),args.toArray());}
        else jdbc.update("INSERT INTO MCM."+table+" ("+String.join(",",columns)+") VALUES ("+String.join(",",Collections.nCopies(columns.size(),"?"))+")",args.toArray());
    }
    private Map<String,List<Map<String,Object>>> loadRows(BigDecimal id){
        var result=new LinkedHashMap<String,List<Map<String,Object>>>();
        String brands="SELECT UM_KIHON_BRAND_ID FROM MCM."+BR+" WHERE UM_KIHON_MITSUMORI_ID=?";
        String kosei="SELECT UM_KIKIKOSEI_ID FROM MCM."+KO+" WHERE UM_KIHON_BRAND_ID IN ("+brands+")";
        String meisai="SELECT UM_KIKIMEISAI_ID FROM MCM."+ME+" WHERE UM_KIKIKOSEI_ID IN ("+kosei+")";
        for(String table:ORDER){String filter=switch(table){case KH,PER,BR->"UM_KIHON_MITSUMORI_ID=?";case KO,BP->"UM_KIHON_BRAND_ID IN ("+brands+")";case ME->"UM_KIKIKOSEI_ID IN ("+kosei+")";default->"UM_KIKIMEISAI_ID IN ("+meisai+")";};result.put(table,jdbc.queryForList("SELECT * FROM MCM."+table+" WHERE "+filter,id));}
        return result;
    }
    private static String naturalKey(String t,Map<String,Object>r){return switch(t){case KH->"header";case PER->String.valueOf(day(r.get("KAISI_DT")));case BR->num(r,"BRANDKOSEI_ID").toPlainString();case KO->num(r,"UM_KIHON_BRAND_ID")+":"+num(r,"KIKIKOSEI_ID");case ME->num(r,"UM_KIKIKOSEI_ID")+":"+num(r,"KIKIMEISAI_ID");case IND->num(r,"UM_KIKIMEISAI_ID")+":"+num(r,"KOTAIKANRI_ID");case PRICE->num(r,"UM_KIKIMEISAI_ID")+":"+num(r,"UM_MITSUMORI_ID")+":"+num(r,"TM_TANKA_ID");default->identity(t,r);};}
    private static String identity(String t,Map<String,Object>r){return IDS.containsKey(t)?num(r,IDS.get(t)).toPlainString():num(r,"UM_KIHON_BRAND_ID")+":"+num(r,"UM_MITSUMORI_ID");}
    public List<Map<String,Object>> tenpoChoices(){
        var rows=jdbc.queryForList("SELECT TENPO_ID,MEISHO1_NK,MEISHO2_NK,MEISHO3_NK,MEISHO4_NK FROM MCM.MCM_MA_TENPO ORDER BY MEISHO1_NK,MEISHO4_NK,TENPO_ID");
        for(var row:rows)row.put("LABEL",List.of("MEISHO1_NK","MEISHO2_NK","MEISHO3_NK","MEISHO4_NK").stream().map(k->text(row,k)).filter(v->!v.isBlank()).collect(java.util.stream.Collectors.joining(" ")));
        return rows;
    }
    private void setTenpo(Map<String,Object>header,String prefix,BigDecimal id){
        if(id==null||id.signum()==0){header.put(prefix+"TENPO_ID",null);for(String c:List.of("MEISHO1_NK","MEISHO2_NK","MEISHO3_NK","MEISHO4_NK","TENPORYAKU_NK"))header.put(prefix+c,"");return;}
        var row=one("SELECT * FROM MCM.MCM_MA_TENPO WHERE TENPO_ID=?",id);header.put(prefix+"TENPO_ID",id);
        for(String c:List.of("MEISHO1_NK","MEISHO2_NK","MEISHO3_NK","MEISHO4_NK","TENPORYAKU_NK"))header.put(prefix+c,row.get(c));
    }

    private void calculateBrand(McmEstimateDraft d,Map<String,Object> b,Map<String,Object> master,String hours){
        BigDecimal max=ZERO;
        for(var p:d.rows(PER)){BigDecimal hard=brandPrice(d,b,p,false);if(hard.compareTo(max)>0)max=hard;}
        BigDecimal cost=num(b,"SYSTEMSEKKEI_KIN").add(num(b,"KIHONSEKKEI_KIN")).add(num(b,"PROGRAMSAKUSEI_KIN"));
        BigDecimal system=systemSupport(cost,text(b,"KEIYAKUJIKANTAI")),technical=technical(max),dts=system.multiply(new BigDecimal("0.25"));
        values(b,"HARDHOSYU_KIN,SYSTEMSUPPORT_KIN,DTSSUPPORT_KIN,DAIFUKUGIJUTSU_KIN,SOFTHOSHU_KIN",max,system,dts,technical,system.add(dts).add(technical));
        for(String name:List.of("SYSTEMSUPPORT_KIN","DTSSUPPORT_KIN","DAIFUKUGIJUTSU_KIN","SOFTHOSHU_KIN"))b.putIfAbsent("HOSEI"+name,b.get(name));
        for(var p:d.rows(PER)){
            var bp=row(d,BP);BigDecimal hard=brandPrice(d,b,p,false),controller=brandPrice(d,b,p,true);
            BigDecimal subtotal=hard.add(num(b,"HOSEISOFTHOSHU_KIN"));
            BigDecimal adjustment=subtotal.remainder(BigDecimal.valueOf(1000));
            if(d.getOriginalId()!=null){
                var previous=jdbc.queryForList("SELECT P.CHOSEI_KIN FROM MCM."+BP+" P JOIN MCM."+BR+" B ON B.UM_KIHON_BRAND_ID=P.UM_KIHON_BRAND_ID JOIN MCM."+PER+" K ON K.UM_MITSUMORI_ID=P.UM_MITSUMORI_ID WHERE B.UM_KIHON_MITSUMORI_ID=? AND B.BRANDKOSEI_ID=? AND K.KAISI_DT=?",d.getOriginalId(),b.get("BRANDKOSEI_ID"),p.get("KAISI_DT"));
                if(!previous.isEmpty())adjustment=num(previous.get(0),"CHOSEI_KIN");
            }
            BigDecimal total=subtotal.subtract(adjustment).add(controller);
            if(on(num(b,"DREMOS_FLG")))total=total.add(num(master,"DREMOS_KIN"));
            if(on(num(b,"REMOTE_FLG")))total=total.add(num(master,"REMOTE_KIN"));
            values(bp,"UM_KIHON_BRAND_ID,UM_MITSUMORI_ID,HARDHOSYU_KIN,CHOSEI_KIN,HOSHU_KIN",num(b,"UM_KIHON_BRAND_ID"),num(p,"UM_MITSUMORI_ID"),hard,adjustment,subtotal.subtract(adjustment));bp.put("_TOTAL",total);
        }
    }
    public static BigDecimal systemSupport(BigDecimal cost,String hours){
        if(hours==null||hours.isEmpty())return ZERO;
        long[] limits={5000000,10000000,15000000,20000000,25000000,30000000,35000000,40000000,45000000,50000000,55000000,60000000,65000000,70000000,80000000,90000000,100000000};
        long[] amounts={150000,300000,450000,700000,1000000,1200000,1400000,1500000,1600000,1700000,1750000,1800000,1850000,1900000,2000000,2100000,2150000};
        BigDecimal result=cost;for(int i=0;i<limits.length;i++)if(cost.compareTo(BigDecimal.valueOf(limits[i]))<0){result=BigDecimal.valueOf(amounts[i]);break;}
        return "24".equals(hours)?result.multiply(new BigDecimal("1.2")):result;
    }
    static BigDecimal technical(BigDecimal hard){long[] bounds={3000000,2500000,2000000,1500000,1100000},amounts={300000,280000,260000,240000,220000};for(int i=0;i<bounds.length;i++)if(hard.compareTo(BigDecimal.valueOf(bounds[i]))>=0)return BigDecimal.valueOf(amounts[i]);return hard.multiply(new BigDecimal("0.2")).setScale(0,RoundingMode.HALF_EVEN);}
    private BigDecimal brandPrice(McmEstimateDraft d,Map<String,Object>b,Map<String,Object>p,boolean controller){
        var ko=d.rows(KO).stream().filter(k->same(num(k,"UM_KIHON_BRAND_ID"),num(b,"UM_KIHON_BRAND_ID"))).map(k->num(k,"UM_KIKIKOSEI_ID")).toList();
        var me=d.rows(ME).stream().filter(m->ko.contains(num(m,"UM_KIKIKOSEI_ID"))).map(m->num(m,"UM_KIKIMEISAI_ID")).toList();
        return d.rows(PRICE).stream().filter(t->me.contains(num(t,"UM_KIKIMEISAI_ID"))&&same(num(t,"UM_MITSUMORI_ID"),num(p,"UM_MITSUMORI_ID"))&&Objects.equals(t.get("_CONTROLLER"),controller)).map(t->num(t,"HYOJUN_KIN").multiply(num(t,"_QUANTITY"))).reduce(ZERO,BigDecimal::add);
    }
    private Map<String,Object> price(McmEstimateDraft d,Map<String,Object>m,Map<String,Object>p,BigDecimal quantity,boolean controller){var t=row(d,PRICE);values(t,"UM_KIKIMEISAI_ID,UM_MITSUMORI_ID,_QUANTITY,_CONTROLLER",num(m,"UM_KIKIMEISAI_ID"),num(p,"UM_MITSUMORI_ID"),quantity,controller);return t;}
    private static void validateParents(Mcm2002uService.SearchResult s){
        for(var m:s.getMeisaiRows())if(on(m.getCheckFlg())&&!Mcm2002uService.selectable(m,s.getSeniMotoKbn()))throw new IllegalStateException("店舗見積・ユーザ契約作成検索画面で、選択した取引先見積に含まれていない為、選択する事は出来ません。");
        for(var k:s.getKotaiRows())if(on(k.getCheckFlg())&&!Mcm2002uService.selectable(k,s.getSeniMotoKbn()))throw new IllegalStateException("店舗見積・ユーザ契約作成検索画面で、選択した取引先見積に含まれていない為、選択する事は出来ません。");
        for(var m:s.getMeisaiRows())if(on(m.getCheckFlg())&&s.getKoseiRows().stream().noneMatch(k->same(k.getKikikoseiId(),m.getKikikoseiId())&&on(k.getCheckFlg())))throw new IllegalStateException("機器明細の親となる機器構成を選択してください。");
        Set<String> seen=new HashSet<>();for(var k:s.getKotaiRows())if(on(k.getCheckFlg())){
            if(s.getMeisaiRows().stream().noneMatch(m->same(m.getKikimeisaiId(),k.getMafKikimeisaiId())&&same(m.getKikikoseiId(),k.getMaeKikikoseiId())&&on(m.getCheckFlg())))throw new IllegalStateException("個体の親となる機器明細を選択してください。");
            if(!seen.add(k.getMadBrandkoseiId()+":"+k.getMaeKikikoseiId()+":"+k.getMafKikimeisaiId()+":"+k.getMagKotaikanriId()))throw new IllegalStateException("同じ個体が重複しています。選定をやり直してください。");
        }
    }
    private Mcm2003uForm preview(McmEstimateDraft d,int mode){
        var f=new Mcm2003uForm();populate(f,d.rows(KH).get(0));f.setSeniMotoKbn(mode);
        for(var p:d.rows(PER)){var r=new Mcm2003uForm.MitsumoriRowForm();populate(r,p);r.setCheckFlg(true);f.getMitsumoriRows().add(r);}
        for(var b:d.rows(BR)){var r=new Mcm2003uForm.BrandRowForm();populate(r,b);r.setSystemSupportSyokeiKin(num(b,"SYSTEMSEKKEI_KIN").add(num(b,"KIHONSEKKEI_KIN")).add(num(b,"PROGRAMSAKUSEI_KIN")));f.getBrandRows().add(r);}
        for(var k:d.rows(KO)){var r=new Mcm2003uForm.KoseiRowForm();populate(r,k);f.getKoseiRows().add(r);}
        for(var m:d.rows(ME)){var r=new Mcm2003uForm.MeisaiRowForm();populate(r,m);f.getMeisaiRows().add(r);}
        return f;
    }
    private static void populate(Object target,Map<String,Object> row){
        var bean=new BeanWrapperImpl(target);for(var p:bean.getPropertyDescriptors()){
            if(!bean.isWritableProperty(p.getName()))continue;
            String key=p.getName().replace("mitsumorigiken","mitsumorikigen").replace("softHosyuKin","softHoshuKin").replace("SoftHosyuKin","SoftHoshuKin").toUpperCase(Locale.ROOT);
            for(var entry:row.entrySet())if(entry.getKey().replace("_","").equals(key)&&entry.getValue()!=null){Object value=entry.getValue();
                if(p.getPropertyType()==boolean.class)value=on(new BigDecimal(value.toString()));
                else if(p.getPropertyType()==String.class)value=value instanceof java.util.Date?format(day(value)):value.toString();
                bean.setPropertyValue(p.getName(),value);break;}
        }
    }
    private Map<String,Object> one(String sql,Object...args){var list=jdbc.queryForList(sql,args);if(list.size()!=1)throw new IllegalStateException("選定元の情報を確認できません。再検索してください。");return list.get(0);}
    private static Map<String,Object> row(McmEstimateDraft d,String table){var r=new LinkedHashMap<String,Object>();if(IDS.containsKey(table))r.put(IDS.get(table),BigDecimal.valueOf(-d.rows(table).size()-1));d.rows(table).add(r);return r;}
    private static void putFrom(Map<String,Object>to,Map<String,Object>from,String cols){for(String col:cols.split(","))if(from.containsKey(col))to.put(col,from.get(col));}
    private static void values(Map<String,Object>m,String cols,Object...values){String[] names=cols.split(",");if(names.length!=values.length)throw new IllegalArgumentException(cols);for(int i=0;i<names.length;i++)m.put(names[i],values[i]);}
    private static BigDecimal num(Map<String,Object>m,String k){Object v=m.get(k);return v==null?ZERO:new BigDecimal(v.toString()).stripTrailingZeros();}
    private static BigDecimal nz(BigDecimal n){return n==null?ZERO:n;}
    private static boolean on(BigDecimal n){return n!=null&&n.compareTo(ONE)==0;}
    private static String text(Map<String,Object>m,String k){return m.get(k)==null?"":m.get(k).toString();}
    private static LocalDate day(Object v){if(v==null)return null;if(v instanceof LocalDate d)return d;if(v instanceof java.sql.Date d)return d.toLocalDate();if(v instanceof Timestamp t)return t.toLocalDateTime().toLocalDate();return LocalDate.parse(v.toString().substring(0,10));}
    private static boolean covers(Map<String,Object>r,LocalDate d){return day(r.get("KAISI_DT"))!=null&&!day(r.get("KAISI_DT")).isAfter(d)&&(day(r.get("SYURYO_DT"))==null||!day(r.get("SYURYO_DT")).isBefore(d));}
}
