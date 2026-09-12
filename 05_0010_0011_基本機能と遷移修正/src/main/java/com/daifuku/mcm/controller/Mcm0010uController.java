package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.time.format.DateTimeFormatter;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpSession;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.dao.DataAccessException;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import org.springframework.web.util.UriUtils;
import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.common.CustomerScreenSupport;
import com.daifuku.mcm.dto.Mcm0010uBrandKoseiDto;
import com.daifuku.mcm.dto.Mcm0010uKikiKoseiDto;
import com.daifuku.mcm.entity.NonyusakiEntity;
import com.daifuku.mcm.entity.PlantEntity;
import com.daifuku.mcm.form.Mcm0010uForm;
import com.daifuku.mcm.service.Mcm0010uService;

/** VB MCM0010U: 納入先検索、親子一覧、画面遷移、管理者用のプラント一括削除。 */
@Controller
@RequestMapping("/mcm0010u")
public class Mcm0010uController extends BaseController {
    public static final String MSG_0110="ブランド構成が設定されていない為、機器構成の設定画面には遷移することは出来ません。";
    private static final String FORM="mcm0010u_search_form", OWNER="mcm0010u_owner", SCREEN="MCM0010U";
    private static final String NONYUSAKI="mcm0010u_nonyusaki", PLANTS="mcm0010u_plants", BRANDS="mcm0010u_brands", KIKI="mcm0010u_kiki", VERSIONS="mcm0010u_versions";
    private static final DateTimeFormatter DISPLAY_DATE=DateTimeFormatter.ofPattern("yyyy/MM/dd");
    private static final Map<String,String> DAYS=Map.of("1","日","2","月","3","火","4","水","5","木","6","金","7","土");
    private static final Map<String,String> DELIVERY_KINDS=Map.of("1","概算","2","撤去","3","実績無","4","不正ﾃﾞｰﾀ");
    @Autowired private Mcm0010uService service;
    @Override protected String getScreenTitle(){return "納入先マスタ検索";}
    @Override protected String getFunctionId(){return SCREEN;}
    @ModelAttribute("brandKoseiRequiredMessage") public String brandKoseiRequiredMessage(){return MSG_0110;}

    private boolean canView(String screen){try{return hasInspectionAuthority()&&service.canView(getLoginUserId(),screen);}catch(DataAccessException ex){logger.warn("{} 権限を取得できません。",screen,ex);return false;}}
    private boolean canUpdate(String screen){try{return hasUpdateAuthority()&&service.canUpdate(getLoginUserId(),screen);}catch(DataAccessException ex){return false;}}
    private boolean canDelete(){try{return hasUpdateAuthority()&&service.canDelete(getLoginUserId());}catch(DataAccessException ex){return false;}}
    private boolean owned(HttpSession s){return Objects.equals(getLoginUserId(),s.getAttribute(OWNER))&&getLoginUserId()!=null;}
    private void initializeOwner(HttpSession s){
        if(!owned(s)){for(String key:List.of(FORM,NONYUSAKI,PLANTS,BRANDS,KIKI,VERSIONS))s.removeAttribute(key);s.setAttribute(OWNER,getLoginUserId());CustomerScreenSupport.rotate(s,SCREEN);}
    }
    private void common(Model m,HttpSession s){
        setCommonAttributes(m,s);m.addAttribute("isUpdateAuthority",canUpdate(SCREEN));
        m.addAttribute("isMaintenance",canDelete());m.addAttribute("canAddNonyusaki",canUpdate("MCM0011U"));
        m.addAttribute("canAddPlant",canUpdate("MCM0012U"));m.addAttribute("workflowToken",CustomerScreenSupport.token(s,SCREEN));
        long started=System.nanoTime();
        m.addAttribute("searchRows",displayRows(s));
        logger.debug("MCM0010U display_rows_ms={}",(System.nanoTime()-started)/1_000_000.0);
    }
    /** VBと同じ検索時点の全件を保持し、表示用の値だけを一度整形して送る。 */
    private Map<String,List<List<String>>> displayRows(HttpSession s){
        var n=new ArrayList<List<String>>();var p=new ArrayList<List<String>>();
        var b=new ArrayList<List<String>>();var k=new ArrayList<List<String>>();
        for(var row:this.<NonyusakiEntity>rows(s,NONYUSAKI))n.add(List.of(text(row.getNonyusakiId()),text(row.getNonyusakiCd()),text(row.getNonyusakiNk()),text(row.getNonyusakikanaKn()),text(row.getNonyusakieimeiEn()),text(row.getNonyusakikojoNk()),text(row.getKyunonyusakiNk()),text(row.getPlanttsuika())));
        for(var row:this.<PlantEntity>rows(s,PLANTS))p.add(List.of(text(row.getPlantId()),text(row.getNonyusakiId()),text(row.getSupportId()),text(row.getPlantNk()),row.getNonyuDt()==null?"":DISPLAY_DATE.format(row.getNonyuDt()),row.getHosyusyusokuDt()==null?"":DISPLAY_DATE.format(row.getHosyusyusokuDt()),DAYS.getOrDefault(text(row.getSystemkadoyoubi()),text(row.getSystemkadoyoubi())),number(row.getSystemkadonissu()),number(row.getSystemkadojikan()),text(row.getKikikoseilink())));
        for(var row:this.<Mcm0010uBrandKoseiDto>rows(s,BRANDS))b.add(List.of(text(row.getPlantId()),DELIVERY_KINDS.getOrDefault(text(row.getNounyuKbn()),text(row.getNounyuKbn())),text(row.getBrandNk()),text(row.getBrandsyosaiNk()),text(row.getSuryo())));
        for(var row:this.<Mcm0010uKikiKoseiDto>rows(s,KIKI))k.add(List.of(text(row.getPlantId()),text(row.getSeizomakerNk()),text(row.getKikihinmeiNk()),text(row.getKikikatashiki()),number(row.getSuryo())));
        return Map.of("nonyusaki",n,"plant",p,"brand",b,"kiki",k);
    }
    private String text(Object value){return Objects.toString(value,"");}
    private String number(BigDecimal value){return value==null?"":value.stripTrailingZeros().toPlainString();}
    private void empty(Model m,HttpSession s){
        m.addAttribute("nonyusakiList",List.of());m.addAttribute("plantList",List.of());m.addAttribute("brandList",List.of());m.addAttribute("kikiList",List.of());
        for(String key:List.of(NONYUSAKI,PLANTS,BRANDS,KIKI))s.setAttribute(key,List.of());s.setAttribute(VERSIONS,Map.of());
    }
    @GetMapping
    public String index(Model m,HttpSession s){
        synchronized(s){
            if(!canView(SCREEN))return "redirect:/menu";initializeOwner(s);
            var form=(Mcm0010uForm)s.getAttribute(FORM);boolean searched=form!=null;
            if(form==null)form=new Mcm0010uForm();m.addAttribute("form",form);m.addAttribute("searched",searched);
            empty(m,s);if(searched)results(form,m,s);CustomerScreenSupport.rotate(s,SCREEN);common(m,s);return "mcm0010u/index";
        }
    }
    @PostMapping("/search")
    public String search(@ModelAttribute("form") Mcm0010uForm form,Model m,HttpSession s){
        synchronized(s){
            if(!canView(SCREEN))return "redirect:/menu";initializeOwner(s);
            s.setAttribute(FORM,CustomerScreenSupport.copy(form));CustomerScreenSupport.rotate(s,SCREEN);
            m.addAttribute("form",form);m.addAttribute("searched",true);empty(m,s);results(form,m,s);common(m,s);return "mcm0010u/index";
        }
    }
    private void results(Mcm0010uForm f,Model m,HttpSession s){
        if(service.isAllEmpty(f)){m.addAttribute("errorMessage","検索条件は1項目以上選択して下さい。");return;}
        if(tooLong(f)){m.addAttribute("errorMessage","検索条件の文字数が上限を超えています。");return;}
        if(containsZenkaku(f.getNonyusakiCd())||containsZenkaku(f.getSupportId())){m.addAttribute("errorMessage","検索結果が1件も存在しません。");return;}
        try{
            long t0=System.nanoTime();
            var n=service.searchNonyusaki(f);
            long t1=System.nanoTime();
            if(n.isEmpty()){m.addAttribute("errorMessage","検索結果が1件も存在しません。");return;}
            var p=service.searchPlants(f,canUpdate(SCREEN)?"2":"1");
            long t2=System.nanoTime();var b=service.searchBrandKosei(f);
            long t3=System.nanoTime();var k=service.searchKikiKosei(f);long t4=System.nanoTime();
            logger.debug("MCM0010U search_ms nonyusaki={} plant={} brand={} kiki={} rows={}/{}/{}/{}",(t1-t0)/1_000_000.0,(t2-t1)/1_000_000.0,(t3-t2)/1_000_000.0,(t4-t3)/1_000_000.0,n.size(),p.size(),b.size(),k.size());
            m.addAttribute("nonyusakiList",n);m.addAttribute("plantList",p);m.addAttribute("brandList",b);m.addAttribute("kikiList",k);
            s.setAttribute(NONYUSAKI,n);s.setAttribute(PLANTS,p);s.setAttribute(BRANDS,b);s.setAttribute(KIKI,k);
            Map<String,String> versions=new LinkedHashMap<>();for(var row:p)versions.put(row.getPlantId().toBigIntegerExact().toString(),Mcm0010uService.plantVersion(row));s.setAttribute(VERSIONS,versions);
        }catch(DataAccessException ex){logger.error("MCM0010U 検索に失敗しました。",ex);empty(m,s);m.addAttribute("errorMessage","検索に失敗しました。接続先とデータを確認してください。");}
    }
    private boolean tooLong(Mcm0010uForm f){return length(f.getNonyusakiCd())>12||length(f.getSupportId())>7||length(f.getNonyusakiNk())>4000||length(f.getPlantNk())>4000||length(f.getJusyo())>4000;}
    private int length(String s){return s==null?0:s.length();}
    private boolean containsZenkaku(String s){return s!=null&&s.chars().anyMatch(c->c>0x7e&&(c<0xff61||c>0xff9f));}
    @SuppressWarnings("unchecked") private <T> List<T> rows(HttpSession s,String key){return s.getAttribute(key) instanceof List<?> l?(List<T>)l:List.of();}
    private boolean displayedPlant(HttpSession s,BigDecimal id){return owned(s)&&this.<PlantEntity>rows(s,PLANTS).stream().anyMatch(p->CustomerScreenSupport.same(p.getPlantId(),id));}
    private boolean displayedNonyusaki(HttpSession s,BigDecimal id){return owned(s)&&this.<NonyusakiEntity>rows(s,NONYUSAKI).stream().anyMatch(n->CustomerScreenSupport.same(n.getNonyusakiId(),id));}
    @GetMapping("/api/nonyusaki/{nonyusakiId}/plants") @ResponseBody
    public List<PlantEntity> getPlantsByNonyusakiId(@PathVariable BigDecimal nonyusakiId,HttpSession s){
        if(!canView(SCREEN)||!displayedNonyusaki(s,nonyusakiId))return List.of();
        return this.<PlantEntity>rows(s,PLANTS).stream().filter(p->CustomerScreenSupport.same(p.getNonyusakiId(),nonyusakiId)).toList();
    }
    @GetMapping("/api/plant/{plantId}/brands") @ResponseBody
    public List<Mcm0010uBrandKoseiDto> getBrandsByPlantId(@PathVariable BigDecimal plantId,HttpSession s){
        if(!canView(SCREEN)||!displayedPlant(s,plantId))return List.of();
        return this.<Mcm0010uBrandKoseiDto>rows(s,BRANDS).stream().filter(b->CustomerScreenSupport.same(b.getPlantId(),plantId)).toList();
    }
    @GetMapping("/api/plant/{plantId}/kikikosei") @ResponseBody
    public List<Mcm0010uKikiKoseiDto> getKikiByPlantId(@PathVariable BigDecimal plantId,HttpSession s){
        if(!canView(SCREEN)||!displayedPlant(s,plantId))return List.of();
        return this.<Mcm0010uKikiKoseiDto>rows(s,KIKI).stream().filter(k->CustomerScreenSupport.same(k.getPlantId(),plantId)).toList();
    }
    @GetMapping("/api/plant/{plantId}/hasBrand") @ResponseBody
    public boolean hasBrandKosei(@PathVariable BigDecimal plantId,HttpSession s){return canView(SCREEN)&&displayedPlant(s,plantId)&&service.hasBrandKosei(plantId);}

    @PostMapping("/plant/delete")
    public String deletePlant(@RequestParam(required=false) List<BigDecimal> plantId,HttpServletRequest request,HttpSession s,RedirectAttributes ra){
        synchronized(s){
            try{
                if(!owned(s)||!canView(SCREEN)||!canDelete())throw new IllegalStateException("削除権限がありません。");
                CustomerScreenSupport.check(s,SCREEN,request);
                if(plantId==null||plantId.isEmpty())throw new IllegalStateException("削除するプラントを選択してください。");
                for(var id:plantId)if(!displayedPlant(s,id))throw new IllegalStateException(CustomerScreenSupport.EXPIRED);
                @SuppressWarnings("unchecked") var versions=(Map<String,String>)s.getAttribute(VERSIONS);
                service.deletePlants(plantId,versions==null?Map.of():versions,getLoginUserId());
                CustomerScreenSupport.rotate(s,SCREEN);ra.addFlashAttribute("infoMessage","プラントと関連情報を削除しました。");
            }catch(IllegalStateException ex){ra.addFlashAttribute("errorMessage",ex.getMessage());}
             catch(DataAccessException ex){logger.error("MCM0010U 削除に失敗しました。",ex);ra.addFlashAttribute("errorMessage","削除に失敗したため変更を取り消しました。関連データを確認してください。");}
            return "redirect:/mcm0010u";
        }
    }
    /** ID以外の親名等はクライアント値を使わず、現在の納入先マスタから組み立てる。 */
    @GetMapping("/plant")
    public String navigatePlant(@RequestParam(required=false) BigDecimal plantId,@RequestParam(required=false) BigDecimal nonyusakiId,HttpSession s,RedirectAttributes ra){
        try{
            if(!canView(SCREEN)||!canView("MCM0012U"))throw new IllegalStateException(CustomerScreenSupport.DENIED);
            PlantEntity plant=null;
            if(plantId!=null){if(!displayedPlant(s,plantId))throw new IllegalStateException(CustomerScreenSupport.EXPIRED);plant=service.findPlant(plantId);nonyusakiId=plant.getNonyusakiId();}
            else if(!canUpdate("MCM0012U")||!displayedNonyusaki(s,nonyusakiId))throw new IllegalStateException(CustomerScreenSupport.DENIED);
            var n=service.findNonyusaki(nonyusakiId);
            s.removeAttribute("mcm0012u_form_backup");
            return "redirect:/mcm0012u?nonyusakiId="+n.getNonyusakiId().toBigIntegerExact()+"&nonyusakiCd="+encode(n.getNonyusakiCd())+"&nonyusakiNk="+encode(n.getNonyusakiNk())+"&returnUrl=/mcm0010u&updateMode="+(plant==null?"0":"1")+(plant==null?"":"&plantId="+plant.getPlantId().toBigIntegerExact());
        }catch(IllegalStateException ex){ra.addFlashAttribute("errorMessage",ex.getMessage());return "redirect:/mcm0010u";}
    }
    @GetMapping("/kikikosei")
    public String navigateKikikosei(@RequestParam BigDecimal plantId,HttpSession s,RedirectAttributes ra){
        try{
            if(!canView(SCREEN)||!canView("MCM0013U"))throw new IllegalStateException(CustomerScreenSupport.DENIED);
            Mcm0010uService.validId(plantId);if(!displayedPlant(s,plantId))throw new IllegalStateException(CustomerScreenSupport.EXPIRED);
            service.findPlant(plantId);
            if(!service.hasBrandKosei(plantId))throw new IllegalStateException(MSG_0110);
            if(!canUpdate(SCREEN)&&this.<PlantEntity>rows(s,PLANTS).stream().noneMatch(p->CustomerScreenSupport.same(p.getPlantId(),plantId)&&"機器構成".equals(p.getKikikoseilink())))throw new IllegalStateException(CustomerScreenSupport.DENIED);
            // 別ルートで開いた同じプラントの古い戻り先を持ち込まない。
            for(String key:List.of("mcm0013uForm","mcm0013u_form_backup","mcm0013u_return_to_0012_url"))s.removeAttribute(key);
            return "redirect:/mcm0013u?plantId="+plantId.toBigIntegerExact();
        }catch(IllegalStateException ex){ra.addFlashAttribute("errorMessage",ex.getMessage());return "redirect:/mcm0010u";}
    }
    private String encode(String s){return UriUtils.encodeQueryParam(s==null?"":s,StandardCharsets.UTF_8);}
}
