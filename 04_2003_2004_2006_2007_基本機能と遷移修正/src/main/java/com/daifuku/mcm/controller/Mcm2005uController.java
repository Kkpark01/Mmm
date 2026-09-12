package com.daifuku.mcm.controller;
import java.math.BigDecimal;
import java.util.*;
import jakarta.servlet.http.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.form.*;
import com.daifuku.mcm.service.*;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;
@Controller @RequestMapping("/mcm2005u")
public class Mcm2005uController extends BaseController {
    @Autowired private Mcm2005uService service;
    @Autowired private Mcm2004uService permissions;
    private static final String FORM="MCM2005U_FORM",SCREEN="mcm2005u";
    @GetMapping
    public String index(@RequestParam(required=false) BigDecimal umKihonBrandId,@RequestParam(required=false) BigDecimal umMitsumoriId,Model model,HttpSession session,RedirectAttributes ra){
        synchronized(session){try{
            var f=(Mcm2005uForm)session.getAttribute(FORM);
            if(umKihonBrandId!=null||umMitsumoriId!=null||f==null){
                f=service.load(umKihonBrandId,umMitsumoriId);
                final var loaded=f;
                if(pending(session).values().stream().anyMatch(p->!same(p.getUmKihonMitsumoriId(),loaded.getUmKihonMitsumoriId())||!same(p.getUmMitsumoriId(),loaded.getUmMitsumoriId())))session.removeAttribute("MCM2005U_PENDING");
                var held=pending(session).get(key(f));if(held!=null)f=held;
                else if(!pending(session).isEmpty()){var shared=pending(session).values().iterator().next();if(shared.isHeaderEdited()){copyHeader(shared.getHeader(),f.getHeader());f.setPeriodBiko(shared.getPeriodBiko());f.setHeaderEdited(true);}}
                if(!same(f.getUmKihonMitsumoriId(),(BigDecimal)session.getAttribute("mcm2005u.parentId")))throw new IllegalStateException(EXPIRED);
                if(held==null && session.getAttribute("MCM2003U_FORM") instanceof Mcm2003uForm parent && Objects.equals(parent.getLastupdateDt(),f.getVersion())) {
                    if(!f.isHeaderEdited())copyHeader(parent,f.getHeader());
                    for(var brand:parent.getBrandRows())if(same(brand.getUmKihonBrandId(),f.getUmKihonBrandId())) {
                        if(brand.getEditedOriginals().containsKey("dremosFlg"))f.setDremosFlg(brand.isDremosFlg());
                        if(brand.getEditedOriginals().containsKey("remoteFlg"))f.setRemoteFlg(brand.isRemoteFlg());
                    }
                }
                rotate(session,SCREEN);session.setAttribute(FORM,f);
            }
            final var displayed=f;
            model.addAttribute("brand",f.getHeader().getBrandRows().stream().filter(b->same(b.getUmKihonBrandId(),displayed.getUmKihonBrandId())).findFirst().orElseThrow());
            model.addAttribute("configurations",f.getHeader().getKoseiRows().stream().filter(k->same(k.getUmKihonBrandId(),displayed.getUmKihonBrandId())).toList());
            setCommonAttributes(model,session);model.addAttribute("form",f);model.addAttribute("readOnly",readOnly(f,session));model.addAttribute("workflowToken",token(session,SCREEN));model.addAttribute("hours",service.hours());model.addAttribute("tenpoChoices",service.tenpoChoices());
            return "mcm2005u/index";
        }catch(Exception ex){logger.error("MCM2005U 読込失敗",ex);ra.addFlashAttribute("error","詳細を読み込めませんでした。見積から開き直してください。");return "redirect:/mcm2003u";}}
    }
    @PostMapping("/save")
    public String save(HttpServletRequest request,HttpSession session,RedirectAttributes ra){synchronized(session){try{
        check(session,SCREEN,request);
        if(!(session.getAttribute(FORM) instanceof Mcm2005uForm original))throw new IllegalStateException(EXPIRED);
        if(readOnly(original,session))throw new IllegalStateException(DENIED);
        var f=copy(original);bindDetail(request,f);session.setAttribute(FORM,f);
        var workspace=new java.util.LinkedHashMap<>(pending(session));workspace.put(key(f),f);
        var updates=new java.util.ArrayList<Mcm2005uForm>();
        for(var item:workspace.values()){
            var updated=copy(item);
            if(f.isHeaderEdited()){copyHeader(f.getHeader(),updated.getHeader());updated.setPeriodBiko(f.getPeriodBiko());updated.setHeaderEdited(true);}
            updates.add(updated);
        }
        String version=service.saveWorkspace(updates,getLoginUserId());
        var refreshed=service.load(f.getUmKihonBrandId(),f.getUmMitsumoriId());
        // 自分の詳細保存による版更新だけを親画面に反映し、他者更新の上書きを防ぐ。
        if(session.getAttribute("MCM2003U_FORM") instanceof Mcm2003uForm parent&&same(parent.getUmKihonMitsumoriId(),f.getUmKihonMitsumoriId())&&Objects.equals(parent.getLastupdateDt(),original.getVersion())) {
            parent.setLastupdateDt(version);
            if(f.isHeaderEdited())copyHeader(refreshed.getHeader(),parent);
            for(var updated:updates)for(var brand:parent.getBrandRows())if(same(brand.getUmKihonBrandId(),updated.getUmKihonBrandId())) {
                brand.setDremosFlg(updated.getDremosFlg());brand.setRemoteFlg(updated.getRemoteFlg());
                brand.getEditedOriginals().remove("dremosFlg");brand.getEditedOriginals().remove("remoteFlg");
            }
        }
        session.removeAttribute("MCM2005U_PENDING");session.setAttribute("mcm2003u.selectedBrandId",f.getUmKihonBrandId());
        refreshed.setSelectedKoseiId(f.getSelectedKoseiId());session.setAttribute(FORM,refreshed);rotate(session,SCREEN);session.setAttribute("mcm2003u.refreshDetails",true);session.setAttribute("mcm2004u.refresh",true);ra.addFlashAttribute("message","詳細を保存し、見積合計を更新しました。");
    }catch(IllegalStateException ex){ra.addFlashAttribute("error",ex.getMessage());}catch(Exception ex){logger.error("MCM2005U 更新失敗",ex);ra.addFlashAttribute("error","保存できませんでした。入力内容を保持しています。");}return "redirect:/mcm2005u";}}
    @GetMapping("/back") public String back(HttpSession s){s.removeAttribute(FORM);s.removeAttribute("MCM2005U_PENDING");rotate(s,SCREEN);s.setAttribute("mcm2003u.refreshDetails",true);return "redirect:/mcm2003u";}
    private void copyHeader(Mcm2003uForm source,Mcm2003uForm target){
        target.setIraitenpoId(source.getIraitenpoId());target.setSofutenpoId(source.getSofutenpoId());target.setIraitantoNk(source.getIraitantoNk());target.setSofutantoNk(source.getSofutantoNk());
        var from=new org.springframework.beans.BeanWrapperImpl(source);var to=new org.springframework.beans.BeanWrapperImpl(target);
        for(String field:java.util.List.of("iraimeisho1Nk","iraimeisho2Nk","iraimeisho3Nk","iraimeisho4Nk","iraitenporyakuNk","sofumeisho1Nk","sofumeisho2Nk","sofumeisho3Nk","sofumeisho4Nk","sofutenporyakuNk"))to.setPropertyValue(field,from.getPropertyValue(field));
    }
    private void bindDetail(HttpServletRequest request,Mcm2005uForm f){
        bind(request,f,"selectedKoseiId");
        if(request.getParameter("headerInputPresent")!=null){bind(request,f,"header.iraitenpoId","header.sofutenpoId","header.iraitantoNk","header.sofutantoNk","periodBiko");f.setHeaderEdited(true);}
        var allowed=new ArrayList<String>(List.of("mitsumorichuki","choseiKin","biko","dremosFlg","remoteFlg"));
        f.setDremosFlg(false);f.setRemoteFlg(false);
        for(int i=0;i<f.getTankaRows().size();i++){
            f.getTankaRows().get(i).setPackFlg(false);
            for(String field:List.of("daifukuhosyujikanId","tenkenkaisu","tenkenyobi","hosyuhoho","packFlg","keiyakunaiyo","keiyakuNo","biko"))allowed.add("tankaRows["+i+"]."+field);
        }
        bind(request,f,allowed.toArray(String[]::new));
    }
    private static String key(Mcm2005uForm f){return f.getUmKihonBrandId().stripTrailingZeros().toPlainString()+":"+f.getUmMitsumoriId().stripTrailingZeros().toPlainString();}
    @SuppressWarnings("unchecked")
    private java.util.Map<String,Mcm2005uForm> pending(HttpSession session){
        var map=(java.util.Map<String,Mcm2005uForm>)session.getAttribute("MCM2005U_PENDING");if(map==null){map=new java.util.LinkedHashMap<>();session.setAttribute("MCM2005U_PENDING",map);}return map;
    }
    @PostMapping("/navigate")
    public String navigate(@RequestParam BigDecimal brandId,HttpServletRequest request,HttpSession session,RedirectAttributes ra){synchronized(session){try{
        check(session,SCREEN,request);if(!(session.getAttribute(FORM) instanceof Mcm2005uForm original))throw new IllegalStateException(EXPIRED);
        if(original.getHeader().getBrandRows().stream().noneMatch(b->same(b.getUmKihonBrandId(),brandId)))throw new IllegalStateException(EXPIRED);
        var f=copy(original);if(!readOnly(f,session))bindDetail(request,f);pending(session).put(key(f),f);
        if(f.isHeaderEdited())for(var other:pending(session).values()){copyHeader(f.getHeader(),other.getHeader());other.setPeriodBiko(f.getPeriodBiko());other.setHeaderEdited(true);}
        session.setAttribute(FORM,f);
        return "redirect:/mcm2005u?umKihonBrandId="+brandId.toPlainString()+"&umMitsumoriId="+f.getUmMitsumoriId().toPlainString();
    }catch(IllegalStateException ex){ra.addFlashAttribute("error",ex.getMessage());return "redirect:/mcm2005u";}}}
    private boolean readOnly(Mcm2005uForm f,HttpSession s){return service.isReadOnly(f)||Boolean.TRUE.equals(s.getAttribute("mcm2005u.parentReadOnly"))||!update(s,"MCM2003U",hasUpdateAuthority(),()->permissions.getAuthority(getLoginUserId(),"MCM2003U"));}
    protected String getScreenTitle(){return "店舗見積ブランド詳細設定";}
    protected String getFunctionId(){return "MCM2005U";}
}
