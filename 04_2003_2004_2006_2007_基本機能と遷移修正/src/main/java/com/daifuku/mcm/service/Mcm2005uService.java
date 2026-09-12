package com.daifuku.mcm.service;
import java.math.BigDecimal;
import java.util.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import com.daifuku.mcm.form.Mcm2005uForm;
import com.daifuku.mcm.repository.Mcm2005uRepository;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;
@Service
public class Mcm2005uService {
    @Autowired private Mcm2005uRepository repo;
    @Autowired private Mcm2003uService estimates;
    @Autowired private McmCustomerTotalsService totals;
    @Autowired private McmCustomerIntegrityService integrity;
    @Transactional(readOnly=true)
    public Mcm2005uForm load(BigDecimal brand,BigDecimal period){
        if(brand==null&&period==null)throw new IllegalStateException(EXPIRED);
        if(brand==null)brand=repo.findBrandIdByMitsumoriId(period);
        if(period==null)period=repo.findMitsumoriIdByBrandId(brand);
        var f=new Mcm2005uForm();f.setUmKihonBrandId(brand);f.setUmMitsumoriId(period);
        repo.fillVersion(f);repo.fillKihonBrand(f,brand);repo.fillMitsumori(f,period);repo.fillJotai(f,f.getUmKihonMitsumoriId());repo.fillBrand(f,period,brand);f.setTankaRows(repo.findTankaRows(period,brand));for(var row:f.getTankaRows())row.setTorihosyujikanNk(repo.hourName(row.getTorihosyujikanId()));f.setHeader(estimates.load(f.getUmKihonMitsumoriId()));repo.fillLayout(f);return f;
    }
    @Transactional
    public String save(Mcm2005uForm f,String user){
        var h=estimates.load(f.getUmKihonMitsumoriId());h.setLastupdateDt(f.getVersion());estimates.requireEditable(h);
        var current=load(f.getUmKihonBrandId(),f.getUmMitsumoriId());
        if(!same(f.getUmKihonMitsumoriId(),current.getUmKihonMitsumoriId())||isReadOnly(current))throw new IllegalStateException(EXPIRED);
        if(f.getTankaRows()==null||f.getTankaRows().size()!=current.getTankaRows().size())throw new IllegalStateException(EXPIRED);
        var seen=new HashSet<BigDecimal>();
        for(var row:f.getTankaRows()){
            var original=current.getTankaRows().stream().filter(r->same(r.getUmTankaId(),row.getUmTankaId())).findFirst().orElseThrow(()->new IllegalStateException(EXPIRED));
            if(!seen.add(row.getUmTankaId().stripTrailingZeros()))throw new IllegalStateException(EXPIRED);
            // DBの参照単価・数量を採用する。改ざんされた送信値を計算に使わない。
            row.setHyojunKin(original.getHyojunKin());row.setSikiriKin(original.getSikiriKin());row.setSuryoNm(original.getSuryoNm());
            if(row.getDaifukuhosyujikanId()!=null && repo.hours().stream().noneMatch(r->same(row.getDaifukuhosyujikanId(),(BigDecimal)r.get("TORIHOSYUJIKAN_ID"))))throw new IllegalStateException("ダイフクの保守時間を選択してください。");
            totals.checkHours(original.getTorihosyujikanId(),row.getDaifukuhosyujikanId());
            if(row.getTenkenyobi()!=null&&!row.getTenkenyobi().isBlank()&&!List.of("1","2","3","4").contains(row.getTenkenyobi()))throw new IllegalStateException("点検曜日を確認してください。");
            if(row.getHosyuhoho()!=null&&!row.getHosyuhoho().isBlank()&&!List.of("F","S","C","H","I","T").contains(row.getHosyuhoho()))throw new IllegalStateException("保守方法を確認してください。");
            amount(row.getTenkenkaisu(),"点検回数",false);length(row.getTenkenyobi(),1);length(row.getHosyuhoho(),2);length(row.getKeiyakunaiyo(),400);length(row.getKeiyakuNo(),50);length(row.getBiko(),4000);
            repo.updateTanka(row,user);
        }
        if(f.getChoseiKin()==null)f.setChoseiKin(BigDecimal.ZERO);amount(f.getChoseiKin(),"調整費",true);length(f.getMitsumorichuki(),1000);length(f.getBiko(),4000);
        if(f.isHeaderEdited()){length(f.getHeader().getIraitantoNk(),80);length(f.getHeader().getSofutantoNk(),80);length(f.getPeriodBiko(),4000);repo.updateHeader(f,user);}
        repo.mergeBrand(f,user);integrity.estimate(f.getUmKihonMitsumoriId());totals.recalculate(f.getUmKihonMitsumoriId(),user);
        return estimates.load(f.getUmKihonMitsumoriId()).getLastupdateDt();
    }
    /** 全ブランドの保留入力を同一トランザクションで登録する。 */
    @Transactional
    public String saveWorkspace(java.util.List<Mcm2005uForm> forms,String user){
        if(forms.isEmpty())throw new IllegalStateException(EXPIRED);
        String version=forms.get(0).getVersion();BigDecimal parent=forms.get(0).getUmKihonMitsumoriId();
        for(var f:forms)if(!same(parent,f.getUmKihonMitsumoriId())||!java.util.Objects.equals(version,f.getVersion()))throw new IllegalStateException("見積情報が変更されています。開き直して確認してください。");
        for(var f:forms){f.setVersion(version);version=save(f,user);}return version;
    }
    private void length(String v,int max){if(v!=null&&v.length()>max)throw new IllegalStateException("入力文字数が上限（"+max+"文字）を超えています。");}
    private void amount(BigDecimal v,String name,boolean signed){if(v!=null&&((!signed&&v.signum()<0)||v.stripTrailingZeros().scale()>0||v.abs().compareTo(new BigDecimal("999999999999"))>0))throw new IllegalStateException(name+"は範囲内の整数で入力してください。");}
    public boolean isReadOnly(Mcm2005uForm f){return List.of("3","4").contains(String.valueOf(f.getJotai()))||List.of("1","2","3").contains(String.valueOf(f.getSyouninJotai()));}
    public List<Map<String,Object>> tenpoChoices(){return repo.tenpoChoices();}
    public List<Map<String,Object>> hours(){return repo.hours();}
}
