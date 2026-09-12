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
import com.daifuku.mcm.dto.Mcm2002uDeliveryDto;
import com.daifuku.mcm.form.Mcm2002uForm;
import com.daifuku.mcm.service.*;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;

@Controller
@RequestMapping("/mcm2002u")
public class Mcm2002uController extends BaseController {
    private static final String SCREEN="mcm2002u",DELIVERY=SCREEN+".delivery",RESULT=SCREEN+".searchResult",FORM=SCREEN+".form";
    private final Mcm2002uService service;
    @Autowired private Mcm2003uDraftService drafts;
    @Autowired private Mcm2004uService permissions;
    public Mcm2002uController(Mcm2002uService service){this.service=service;}
    private boolean editable(HttpSession s){return update(s,"MCM2003U",hasUpdateAuthority(),()->permissions.getAuthority(getLoginUserId(),"MCM2003U"));}

    @GetMapping
    public String index(HttpSession session,Model model,RedirectAttributes ra){
        synchronized(session){try{
            if(!(session.getAttribute(DELIVERY) instanceof Mcm2002uDeliveryDto delivery))return "redirect:/mcm2001u";
            var result=(Mcm2002uService.SearchResult)session.getAttribute(RESULT);
            var form=(Mcm2002uForm)session.getAttribute(FORM);
            if(result==null||form==null){
                result=service.loadScreen(delivery);form=new Mcm2002uForm();
                form.setNonyusakiCd(delivery.getNonyusakiCd());form.setNonyusakiNk(delivery.getNonyusakiNk());form.setPlantNk(delivery.getPlantNk());form.setSupportId(delivery.getSupportId());
                form.setSeniMotoKbn(result.getSeniMotoKbn());form.setKaisiDt(result.getOldKaisiDt());
                form.setKoseiRows(result.getKoseiRows());form.setMeisaiRows(result.getMeisaiRows());form.setKotaiRows(result.getKotaiRows());form.setTankaRows(result.getTankaRows());
                if(result.isShowKeiyakuJikantai()&&delivery.getUmKihonMitsumoriId()!=null)form.setKeiyakuJikantai(drafts.sourceHours(delivery.getUmKihonMitsumoriId()));
                session.setAttribute(RESULT,result);session.setAttribute(FORM,form);rotate(session,SCREEN);
            }
            model.addAttribute("selectableMeisai",java.util.stream.IntStream.range(0,result.getMeisaiRows().size()).filter(i->Mcm2002uService.selectable(((Mcm2002uService.SearchResult)session.getAttribute(RESULT)).getMeisaiRows().get(i),delivery.getSeniMotoKbn())).boxed().toList());
            model.addAttribute("selectableKotai",java.util.stream.IntStream.range(0,result.getKotaiRows().size()).filter(i->Mcm2002uService.selectable(((Mcm2002uService.SearchResult)session.getAttribute(RESULT)).getKotaiRows().get(i),delivery.getSeniMotoKbn())).boxed().toList());
            setCommonAttributes(model,session);model.addAttribute("form",form);model.addAttribute("searchResult",result);
            model.addAttribute("showKaisiDt",result.isShowKaisiDt());model.addAttribute("showKeiyakuJikantai",result.isShowKeiyakuJikantai());
            model.addAttribute("readOnly",result.isReadOnly()||!editable(session));model.addAttribute("kaisiDtReadOnly",result.isKaisiDtReadOnly());model.addAttribute("workflowToken",token(session,SCREEN));
            model.addAttribute("seniMotoLabel",switch(result.getSeniMotoKbn()){case 0->"新規作成";case 1->"見積修正";case 2->"複製";default->"";});
            return "mcm2002u/index";
        }catch(Exception ex){logger.error("MCM2002U 読み込み失敗",ex);ra.addFlashAttribute("error","選定情報を読み込めませんでした。再検索してください。");return "redirect:/mcm2001u";}}
    }
    @PostMapping("/kihon-settei")
    public String next(HttpServletRequest request,HttpSession session,RedirectAttributes ra){
        synchronized(session){try{
            check(session,SCREEN,request);
            if(!(session.getAttribute(DELIVERY) instanceof Mcm2002uDeliveryDto delivery)||!(session.getAttribute(RESULT) instanceof Mcm2002uService.SearchResult result)||!(session.getAttribute(FORM) instanceof Mcm2002uForm form))throw new IllegalStateException(EXPIRED);
            if(!editable(session)||result.isReadOnly())throw new IllegalStateException(DENIED);
            bind(request,form,"kaisiDt","keiyakuJikantai");
            var k=selections(request,"koseiSelection",result.getKoseiRows().size());var m=selections(request,"meisaiSelection",result.getMeisaiRows().size());var i=selections(request,"kotaiSelection",result.getKotaiRows().size());
            for(int n=0;n<result.getKoseiRows().size();n++)result.getKoseiRows().get(n).setCheckFlg(k.contains(n)?BigDecimal.ONE:BigDecimal.ZERO);
            for(int n=0;n<result.getMeisaiRows().size();n++)result.getMeisaiRows().get(n).setCheckFlg(m.contains(n)?BigDecimal.ONE:BigDecimal.ZERO);
            for(int n=0;n<result.getKotaiRows().size();n++)result.getKotaiRows().get(n).setCheckFlg(i.contains(n)?BigDecimal.ONE:BigDecimal.ZERO);
            var errors=service.validateKihonSettei(form,result.getTmKeiyakujikanIds()!=null&&!result.getTmKeiyakujikanIds().isEmpty());
            if(!errors.isEmpty()){ra.addFlashAttribute("errors",errors);return "redirect:/mcm2002u";}
            var user=getLoginUserInfo(session);
            var draft=drafts.prepare(delivery,result,form,user==null?getLoginUserId():user.getUserName());
            if(session.getAttribute("mcm2003u.draft") instanceof com.daifuku.mcm.form.McmEstimateDraft prior
                    && session.getAttribute("MCM2003U_FORM") instanceof com.daifuku.mcm.form.Mcm2003uForm pending
                    && java.util.Objects.equals(prior.getOriginalId(),draft.getOriginalId()) && java.util.Objects.equals(pending.getPlantId(),draft.getForm().getPlantId())) {
                drafts.preserveBrandMethods(prior,pending,draft);
                for(var period:draft.getForm().getMitsumoriRows())pending.getMitsumoriRows().stream().filter(old->java.util.Objects.equals(old.getKaisiDt(),period.getKaisiDt())).findFirst().ifPresent(old->period.setCheckFlg(old.isCheckFlg()));
                var from=new org.springframework.beans.BeanWrapperImpl(pending);var to=new org.springframework.beans.BeanWrapperImpl(draft.getForm());
                for(String field:List.of("umMitsumoriNo","mitsumorigiken","mitsumoriLevel","iraitenpoId","sofutenpoId","iraitantoNk","sofutantoNk","mitsumoriJouken","hosyuhoho","biko"))to.setPropertyValue(field,from.getPropertyValue(field));
            }
            for(String key:Collections.list(session.getAttributeNames()))if(key.startsWith("mcm2003u."))session.removeAttribute(key);
            session.setAttribute("mcm2003u.draft",draft);session.setAttribute("MCM2003U_FORM",draft.getForm());
            session.setAttribute("mcm2003u.seniMotoKbn",delivery.getSeniMotoKbn());rotate(session,"mcm2003u");rotate(session,SCREEN);
            return "redirect:/mcm2003u";
        }catch(IllegalStateException ex){ra.addFlashAttribute("error",ex.getMessage());}
        catch(Exception ex){logger.error("MCM2002U 見積生成失敗",ex);ra.addFlashAttribute("error","見積を生成できませんでした。選定内容を保持しています。");}
        return "redirect:/mcm2002u";}
    }
    @GetMapping("/back")
    public String back(HttpSession s){
        if(s.getAttribute(DELIVERY) instanceof Mcm2002uDeliveryDto d&&d.getSeniMotoKbn()==1)return "redirect:/mcm2004u";
        return "redirect:/mcm2001u";
    }
    @Override protected String getScreenTitle(){return "保守見積機器選定";}
    @Override protected String getFunctionId(){return "MCM2002U";}
}
