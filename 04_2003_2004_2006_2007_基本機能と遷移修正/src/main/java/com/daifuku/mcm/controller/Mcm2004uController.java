package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.*;
import jakarta.servlet.http.HttpSession;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.WebDataBinder;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import com.daifuku.mcm.common.AppConstants;
import com.daifuku.mcm.dto.*;
import com.daifuku.mcm.form.*;
import com.daifuku.mcm.service.Mcm2004uService;
import com.daifuku.mcm.service.Mcm2004uService.SearchResult;

/** VB MCM2004U: search, seven row links, maintenance deletion and screen handoff. */
@Controller @RequestMapping("/mcm2004u")
public class Mcm2004uController {
    private static final String FORM="mcm2004u.form", RESULT="mcm2004u.searchResult", TOKEN="mcm2004u.token", DELIVERY="mcm2004u.delivery";
    private static final String NONE="行が選択されていません。", DENIED="権限がないため遷移できません。";
    private final Mcm2004uService service;
    public Mcm2004uController(Mcm2004uService service){this.service=service;}
    @InitBinder("form") public void bind(WebDataBinder binder){binder.setAllowedFields("umMitsumoriNo","keiyakuNo","mitsumoriDtStart","mitsumoriDtEnd","sofutenpoId","nonyusakiCd","nonyusakiNk","supportId","plantNk","jotaiMitsumori","jotaiKeiyaku","jotaiKaiyaku","jotaiHaki","syouninJotaiSakuseichu","syouninJotaiSinsachu","syouninJotaiShoninchu","syouninJotaiShoninzumi","syouninJotaiSashimodoshichu");}

    @GetMapping public String index(HttpSession session,Model model){synchronized(session){
        try {
            var form=session.getAttribute(FORM) instanceof Mcm2004uForm f?f:new Mcm2004uForm();
            var result=session.getAttribute(RESULT) instanceof SearchResult r?r:permissions(session);
            if(session.getAttribute(DELIVERY) instanceof Mcm2004uDeliveryDto delivery){
                result=service.loadScreen(delivery,authority(session,"MCM2003U"),authority(session,"MCM2006U"),maintenance(session));
                session.setAttribute("mcm2004u.activeDelivery",delivery);
                session.removeAttribute(DELIVERY);rotate(session);
                if(result.getRows().isEmpty())result.setErrors(List.of("検索結果が1件も存在しません。"));
            }
            if (Boolean.TRUE.equals(session.getAttribute("mcm2004u.refresh"))) {
                result=session.getAttribute("mcm2004u.activeDelivery") instanceof Mcm2004uDeliveryDto active
                    ? service.loadScreen(active,authority(session,"MCM2003U"),authority(session,"MCM2006U"),maintenance(session)) : service.search(form);
                session.removeAttribute("mcm2004u.refresh"); rotate(session);
            }
            return display(session,model,form,result);
        }catch(org.springframework.dao.DataAccessException e){return failure(session,model,e);}
    }}
    @PostMapping("/search") public String search(@ModelAttribute("form") Mcm2004uForm form,HttpSession session,Model model){synchronized(session){
        session.removeAttribute("mcm2004u.activeDelivery");
        session.setAttribute(FORM,form);session.setAttribute(RESULT,permissions(session));rotate(session);
        try{return display(session,model,form,service.search(form));}
        catch(org.springframework.dao.DataAccessException e){return failure(session,model,e);}
    }}
    @GetMapping("/re-search") public String reSearch(HttpSession session,Model model){synchronized(session){
        var form=session.getAttribute(FORM) instanceof Mcm2004uForm f?f:new Mcm2004uForm();rotate(session);
        try{return display(session,model,form,service.search(form));}
        catch(org.springframework.dao.DataAccessException e){return failure(session,model,e);}
    }}
    private String display(HttpSession session,Model model,Mcm2004uForm form,SearchResult result){
        var p=permissions(session);result.setShowMitsumoriCopy(p.isShowMitsumoriCopy());result.setShowMitsumoriDel(p.isShowMitsumoriDel());
        result.setShowKeiyakuKeiyaku(p.isShowKeiyakuKeiyaku());result.setShowKeiyakuDel(p.isShowKeiyakuDel());result.setNewButtonEnabled(p.isNewButtonEnabled());result.setDeleteButtonVisible(p.isDeleteButtonVisible());
        form.setRows(result.getRows());session.setAttribute(FORM,form);session.setAttribute(RESULT,result);
        if(session.getAttribute(TOKEN)==null)rotate(session);
        model.addAttribute("form",form);model.addAttribute("sr",result);model.addAttribute("searchToken",session.getAttribute(TOKEN));
        model.addAttribute("sofutenpoList",service.getSofutenpoList());
        if(!model.containsAttribute("errors")&&!result.getErrors().isEmpty())model.addAttribute("errors",result.getErrors());
        if(model.containsAttribute("errors"))model.addAttribute("errorTitle",title(model.getAttribute("errors")));
        return "mcm2004u/index";
    }
    private String failure(HttpSession session,Model model,Exception e){
        org.slf4j.LoggerFactory.getLogger(getClass()).error("MCM2004U search failed",e);
        model.addAttribute("errors",List.of("検索処理を完了できませんでした。条件を保持しています。もう一度検索してください。"));
        return display(session,model,session.getAttribute(FORM) instanceof Mcm2004uForm f?f:new Mcm2004uForm(),permissions(session));
    }
    private static String title(Object errors){String t=String.valueOf(errors);return t.contains("検索")?"検索エラー":t.contains("権限")?"権限エラー":t.contains("使用されている為")||t.contains("他のユーザ")?"排他エラー":"入力エラー";}
    private SearchResult permissions(HttpSession s){return service.loadScreen(null,authority(s,"MCM2003U"),authority(s,"MCM2006U"),maintenance(s));}
    private String authority(HttpSession s,String id){
        Object value=s.getAttribute("authority."+id);
        if(value==null)try{value=service.getAuthority(user(s),id);}catch(org.springframework.dao.DataAccessException e){org.slf4j.LoggerFactory.getLogger(getClass()).warn("MCM2004U authority lookup failed",e);}
        return "2".equals(String.valueOf(value))||"UPDATE".equals(value)?"UPDATE":"INSPECTION";
    }
    private boolean maintenance(HttpSession s){
        if(s.getAttribute("isMaintenance") instanceof Boolean b)return b;
        try{return service.isMaintenance(user(s));}catch(org.springframework.dao.DataAccessException e){return false;}
    }
    private String user(HttpSession s){Object v=s.getAttribute(AppConstants.SESSION_USER_ID);return v==null?"SYSTEM":v.toString();}
    private static void rotate(HttpSession s){s.setAttribute(TOKEN,UUID.randomUUID().toString());}
    private void checkToken(HttpSession s,String token){if(token==null||!token.equals(s.getAttribute(TOKEN)))throw new IllegalStateException(Mcm2004uService.CHANGED);}
    private static boolean text(String s){return s!=null&&!s.isBlank();}
    private Mcm2004uRowDto selected(HttpSession s,BigDecimal id,BigDecimal contract,String operation){
        if(id==null||id.signum()<=0||!(s.getAttribute(RESULT) instanceof SearchResult result))throw new IllegalStateException(Mcm2004uService.CHANGED);
        return result.getRows().stream().filter(r->Mcm2004uService.same(id,r.getUmKihonMitsumoriId()))
            .filter(r->operation.equals("delete-row")||Mcm2004uService.same(contract,r.getUkKeiyakuId()))
            .findFirst().orElseThrow(()->new IllegalStateException(Mcm2004uService.CHANGED));
    }
    private boolean allowed(Mcm2004uRowDto r,String op){return switch(op){
        case "discard-mitsumori" -> text(r.getMitsumoriDel());case "discard-keiyaku" -> text(r.getKeiyakuDel());
        case "mcm2003u" -> text(r.getUmMitsumoriNo());case "mcm2006u" -> text(r.getKeiyakuNo());
        case "mcm2002u-upd" -> text(r.getMitsumoriUpd());case "mcm2002u-copy" -> text(r.getMitsumoriCopy());
        case "mcm2007u" -> text(r.getKeiyakuKeiyaku());case "delete-row" -> true;default -> false;
    };}
    @PostMapping({"/discard-mitsumori","/discard-keiyaku","/delete-row"})
    public String mutate(jakarta.servlet.http.HttpServletRequest request,@RequestParam(value="umKihonMitsumoriId",required=false) List<BigDecimal> ids,
        @RequestParam(value="ukKeiyakuId",required=false) BigDecimal contract,@RequestParam(value="searchToken",required=false) String token,
        HttpSession session,RedirectAttributes ra){synchronized(session){
        String op=request.getRequestURI().substring(request.getRequestURI().lastIndexOf('/')+1);
        try{
            checkToken(session,token);
            if(op.equals("delete-row")?!maintenance(session):!authority(session,op.equals("discard-keiyaku")?"MCM2006U":"MCM2003U").equals("UPDATE"))throw new IllegalStateException(DENIED);
            if(ids==null||ids.isEmpty())throw new IllegalStateException(NONE);
            if(ids.stream().anyMatch(id->id==null||id.signum()<=0))throw new IllegalStateException(Mcm2004uService.CHANGED);
            var unique=ids.stream().map(BigDecimal::stripTrailingZeros).distinct().toList();
            if(!op.equals("delete-row")&&unique.size()!=1)throw new IllegalStateException(Mcm2004uService.CHANGED);
            var rows=new ArrayList<Mcm2004uRowDto>();for(var id:unique){var row=selected(session,id,contract,op);if(!allowed(row,op))throw new IllegalStateException(Mcm2004uService.CHANGED);rows.add(row);}
            rotate(session);
            if(op.equals("delete-row"))service.deleteRows(unique,user(session));
            else if(op.equals("discard-mitsumori")){
                var errors=service.discardMitsumori(unique.get(0),rows.get(0).getUkKeiyakuId(),user(session));
                if(!errors.isEmpty()){ra.addFlashAttribute("errors",errors);return "redirect:/mcm2004u";}
            }else service.discardKeiyaku(unique.get(0),rows.get(0).getUkKeiyakuId(),user(session));
            if(op.equals("delete-row")){
                var result=(SearchResult)session.getAttribute(RESULT);result.setRows(result.getRows().stream().filter(r->unique.stream().noneMatch(id->Mcm2004uService.same(id,r.getUmKihonMitsumoriId()))).toList());result.setErrors(List.of());
                return "redirect:/mcm2004u";
            }
            return "redirect:/mcm2004u/re-search";
        }catch(IllegalStateException e){ra.addFlashAttribute("errors",List.of(e.getMessage()));}
        catch(org.springframework.transaction.TransactionException e){org.slf4j.LoggerFactory.getLogger(getClass()).error("MCM2004U transaction failed",e);ra.addFlashAttribute("errors",List.of("処理結果を確認できませんでした。再検索して状態を確認してください。"));}
        catch(org.springframework.dao.DataAccessException e){org.slf4j.LoggerFactory.getLogger(getClass()).error("MCM2004U operation failed",e);ra.addFlashAttribute("errors",List.of("処理を完了できませんでした。再検索してから、もう一度操作してください。"));}
        return "redirect:/mcm2004u";
    }}
    @GetMapping("/new") public String newRow(HttpSession session,RedirectAttributes ra){
        if(!authority(session,"MCM2003U").equals("UPDATE")){ra.addFlashAttribute("errors",List.of(DENIED));return "redirect:/mcm2004u";}
        clear(session,"mcm2003u.");session.removeAttribute("MCM2003U_FORM");clear(session,"mcm2001u.");return "redirect:/mcm2001u";
    }
    @GetMapping("/link/{destination}") public String link(@PathVariable String destination,@RequestParam BigDecimal umKihonMitsumoriId,
        @RequestParam(required=false) BigDecimal ukKeiyakuId,@RequestParam(required=false) String searchToken,HttpSession session,RedirectAttributes ra){synchronized(session){
        try{
            checkToken(session,searchToken);
            var old=selected(session,umKihonMitsumoriId,ukKeiyakuId,"mcm2006u".equals(destination)?"keiyaku":"link");
            var row=service.requireRow(umKihonMitsumoriId,old.getUkKeiyakuId(),"link");
            if(!allowed(row,destination))throw new IllegalStateException(Mcm2004uService.CHANGED);
            if(Set.of("mcm2002u-upd","mcm2002u-copy","mcm2007u").contains(destination)
                    && (row.getPlantId()==null||row.getPlantId().signum()<=0||row.getNonyusakiId()==null||row.getNonyusakiId().signum()<=0))
                throw new IllegalStateException("納入先・プラント情報を確認できません。再検索してください。");
            if((destination.equals("mcm2002u-copy")&&!authority(session,"MCM2003U").equals("UPDATE"))||(destination.equals("mcm2007u")&&!authority(session,"MCM2006U").equals("UPDATE")))throw new IllegalStateException(DENIED);
            switch(destination){
                case "mcm2003u":
                    clear(session,"mcm2003u.");session.removeAttribute("MCM2003U_FORM");session.setAttribute("mcm2003u.umKihonMitsumoriId",row.getUmKihonMitsumoriId());session.setAttribute("mcm2003u.seniMotoKbn",3);return "redirect:/mcm2003u";
                case "mcm2006u":
                    session.removeAttribute("MCM2006U_DELIVERY"); session.removeAttribute("MCM2007U_FORM");
                    clear(session,"mcm2006u.");session.setAttribute("mcm2006u.from2004",true);session.removeAttribute("MCM2006U_FORM");return "redirect:/mcm2006u?ukKeiyakuId="+row.getUkKeiyakuId().toPlainString()+"&seniMotoKbn=0";
                case "mcm2002u-upd":
                    clear(session,"mcm2003u.");session.removeAttribute("MCM2003U_FORM");var d=new Mcm2002uDeliveryDto();org.springframework.beans.BeanUtils.copyProperties(row,d);d.setSeniMotoKbn(1);clear(session,"mcm2002u.");session.setAttribute("mcm2002u.delivery",d);return "redirect:/mcm2002u";
                case "mcm2002u-copy":
                    clear(session,"mcm2003u.");session.removeAttribute("MCM2003U_FORM");var copy=new Mcm2001uDeliveryDto();copy.setUmKihonMitsumoriId(row.getUmKihonMitsumoriId().intValueExact());copy.setPlantId(row.getPlantId().intValueExact());copy.setNonyusakiId(row.getNonyusakiId().intValueExact());copy.setNonyusakiCd(row.getNonyusakiCd());copy.setNonyusakiNk(row.getNonyusakiNk());copy.setSupportId(row.getSupportId());copy.setPlantNk(row.getPlantNk());copy.setSeniMotoKbn(2);clear(session,"mcm2001u.");session.setAttribute("mcm2001u.delivery",copy);return "redirect:/mcm2001u";
                case "mcm2007u":
                    session.removeAttribute("MCM2006U_DELIVERY");
                    var parent=new Mcm2006uForm();org.springframework.beans.BeanUtils.copyProperties(row,parent,"ukKeiyakuId");parent.setSeniMotoKbn(1);clear(session,"mcm2006u.");clear(session,"mcm2007u.");session.setAttribute("MCM2006U_FORM",parent);session.removeAttribute("MCM2007U_FORM");session.removeAttribute("MCM2007U_STEP1_FORM");session.setAttribute("mcm2007u.from2004",row.getUmKihonMitsumoriId());session.setAttribute("mcm2007u.from2004Cancel",true);session.setAttribute("mcm2006u.from2004",true);return "redirect:/mcm2007u?seniMotoKbn=1";
                default:throw new IllegalStateException(Mcm2004uService.CHANGED);
            }
        }catch(IllegalStateException|ArithmeticException e){ra.addFlashAttribute("errors",List.of(e instanceof ArithmeticException?Mcm2004uService.CHANGED:e.getMessage()));return "redirect:/mcm2004u";}
        catch(org.springframework.dao.DataAccessException e){ra.addFlashAttribute("errors",List.of("画面の情報を取得できませんでした。再検索してから、もう一度操作してください。"));return "redirect:/mcm2004u";}
    }}
    private static void clear(HttpSession s,String prefix){for(var name:Collections.list(s.getAttributeNames()))if(name.startsWith(prefix))s.removeAttribute(name);}
}
