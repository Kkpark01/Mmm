package com.daifuku.mcm.controller;
import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.form.*;
import com.daifuku.mcm.service.*;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;
import static com.daifuku.mcm.repository.Mcm3007uRepository.id;
import jakarta.servlet.http.*;
import java.util.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.WebDataBinder;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import org.springframework.web.util.UriComponentsBuilder;
@Controller @RequestMapping("/mcm3007u") public class Mcm3007uController extends BaseController {
 @Autowired private Mcm3007uService service;
 @Autowired private Mcm2004uService permissions;
 private static final String FORM="MCM3007U_FORM",SCREEN="mcm3007u";
 @InitBinder public void fields(WebDataBinder b){b.setAllowedFields("nonyusakiCd","nonyusakiNk","supportId","plantNk","jotaiUmMitsumori","jotaiUmKeiyaku","jotaiUmKaiyaku","jotaiUmHaki","jotaiTkMitsumori","jotaiTkKeiyaku","jotaiTkKaiyaku","jotaiTkHaki");}
 private String authority(HttpSession s,String screen){Object a=s.getAttribute("authority."+screen);if(a!=null)return a.toString();try{return permissions.getAuthority(getLoginUserId(),screen);}catch(org.springframework.dao.DataAccessException e){return "0";}}
 private boolean canView(HttpSession s,String screen){return hasInspectionAuthority()&&Set.of("1","2","INSPECTION","UPDATE").contains(authority(s,screen));}
 private boolean canUpdate(HttpSession s,String screen){return update(s,screen,hasUpdateAuthority(),()->authority(s,screen));}
 @GetMapping public String index(Model m,HttpSession s){synchronized(s){var f=(Mcm3007uForm)s.getAttribute(FORM);if(f==null){f=new Mcm3007uForm();s.setAttribute(FORM,f);}setCommonAttributes(m,s);m.addAttribute("form",f);m.addAttribute("workflowToken",token(s,SCREEN));var rights=new HashMap<String,Boolean>();for(String k:List.of("2003","2006","1004","1005","0011","0012","0013"))rights.put(k,canView(s,"MCM"+k+"U"));rights.put("umUpdate",canUpdate(s,"MCM2006U"));rights.put("tkUpdate",canUpdate(s,"MCM1005U"));m.addAttribute("rights",rights);return "mcm3007u/index";}}
 @PostMapping("/search") public String search(@ModelAttribute Mcm3007uForm f,HttpSession s,RedirectAttributes ra){synchronized(s){try{service.search(f);if(f.getGrid1Rows().isEmpty())ra.addFlashAttribute("error","対象データが存在しません。");}catch(IllegalStateException e){empty(f);ra.addFlashAttribute("error",e.getMessage());}catch(Exception e){empty(f);logger.error("3007U検索失敗",e);ra.addFlashAttribute("error","検索できませんでした。条件を確認して再検索してください。");}s.setAttribute(FORM,f);rotate(s,SCREEN);return "redirect:/mcm3007u";}}
 private static void empty(Mcm3007uForm f){f.setGrid1Rows(List.of());f.setGrid2Rows(List.of());f.setGrid3Rows(List.of());}
 private Map<String,String> row(HttpSession s,HttpServletRequest req,String kind,int index){check(s,SCREEN,req);var f=(Mcm3007uForm)s.getAttribute(FORM);if(f==null||index<0)throw new IllegalStateException(EXPIRED);try{return switch(kind){case "master"->f.getGrid1Rows().get(index).getDisplay();case "um"->f.getGrid2Rows().get(index).getDisplay();case "tk"->f.getGrid3Rows().get(index).getDisplay();default->throw new IllegalStateException(EXPIRED);};}catch(IndexOutOfBoundsException e){throw new IllegalStateException(EXPIRED);}}
 private Map<String,String> fresh(HttpSession s,String kind,Map<String,String> old){var f=copy((Mcm3007uForm)s.getAttribute(FORM));service.search(f);var rows=switch(kind){case "master"->f.getGrid1Rows().stream().map(Mcm3007uForm.Grid1RowForm::getDisplay).toList();case "um"->f.getGrid2Rows().stream().map(Mcm3007uForm.Grid2RowForm::getDisplay).toList();default->f.getGrid3Rows().stream().map(Mcm3007uForm.Grid3RowForm::getDisplay).toList();};if(rows.stream().noneMatch(old::equals))throw new IllegalStateException(EXPIRED);return old;}
 @PostMapping({"/break-um","/break-tk"}) public String discard(@RequestParam int rowIndex,HttpServletRequest request,HttpSession s,RedirectAttributes ra){synchronized(s){try{boolean um=request.getServletPath().endsWith("break-um")||request.getRequestURI().endsWith("break-um");if(!canUpdate(s,um?"MCM2006U":"MCM1005U"))throw new IllegalStateException(DENIED);String kind=um?"um":"tk";var r=fresh(s,kind,row(s,request,kind,rowIndex));service.discard(um,r,getLoginUserId());service.search((Mcm3007uForm)s.getAttribute(FORM));rotate(s,SCREEN);ra.addFlashAttribute("message","契約を破棄しました。");}catch(IllegalStateException e){ra.addFlashAttribute("error",e.getMessage());}catch(Exception e){logger.error("3007U破棄失敗",e);ra.addFlashAttribute("error","契約を破棄できませんでした。再検索してください。");}return "redirect:/mcm3007u";}}
 @GetMapping("/link/{destination}") public String link(@PathVariable String destination,@RequestParam String kind,@RequestParam int rowIndex,HttpServletRequest request,HttpSession s,RedirectAttributes ra){synchronized(s){try{
  var d=fresh(s,kind,row(s,request,kind,rowIndex));String screen=switch(destination){case "mcm2007u"->"MCM2006U";case "mcm1006u"->"MCM1005U";default->destination.toUpperCase(Locale.ROOT);};
  if(!Set.of("mcm2003u","mcm2006u","mcm2007u","mcm1004u","mcm1005u","mcm1006u","mcm0011u","mcm0012u","mcm0013u").contains(destination)||!canView(s,screen))throw new IllegalStateException(DENIED);
  if((destination.equals("mcm2007u")||destination.equals("mcm1006u"))&&(!canUpdate(s,screen)||d.getOrDefault("KEIYAKU_KEIYAKU","").isBlank()))throw new IllegalStateException(DENIED);
  var u=UriComponentsBuilder.fromPath("/"+destination);
  switch(destination){
   case "mcm2003u": if(!kind.equals("um")||id(d,"UM_KIHON_MITSUMORI_ID")==null)throw new IllegalStateException(EXPIRED);clear(s,"mcm2003u.");s.removeAttribute("MCM2003U_FORM");s.setAttribute("mcm2003u.umKihonMitsumoriId",id(d,"UM_KIHON_MITSUMORI_ID"));s.setAttribute("mcm2003u.seniMotoKbn",3);s.setAttribute("mcm2003u.returnTo","/mcm3007u");break;
   case "mcm2006u": if(!kind.equals("um")||id(d,"UK_KEIYAKU_ID")==null)throw new IllegalStateException(EXPIRED);clear(s,"mcm2006u.");s.removeAttribute("MCM2006U_FORM");s.removeAttribute("MCM2006U_DELIVERY");s.setAttribute("mcm2006u.returnTo","/mcm3007u");u.queryParam("ukKeiyakuId",d.get("UK_KEIYAKU_ID")).queryParam("seniMotoKbn",0);break;
   case "mcm2007u": if(!kind.equals("um")||id(d,"PLANT_ID")==null||id(d,"NONYUSAKI_ID")==null)throw new IllegalStateException(EXPIRED);var parent=new Mcm2006uForm();parent.setPlantId(id(d,"PLANT_ID"));parent.setNonyusakiId(id(d,"NONYUSAKI_ID"));parent.setNonyusakiCd(d.get("NONYUSAKI_CD"));parent.setNonyusakiNk(d.get("NONYUSAKI_NK"));parent.setSupportId(d.get("SUPPORT_ID"));parent.setPlantNk(d.get("PLANT_NK"));parent.setSeniMotoKbn(1);clear(s,"mcm2006u.");clear(s,"mcm2007u.");s.removeAttribute("MCM2006U_DELIVERY");s.removeAttribute("MCM2007U_FORM");s.removeAttribute("MCM2007U_STEP1_FORM");s.setAttribute("MCM2006U_FORM",parent);s.setAttribute("mcm2006u.returnTo","/mcm3007u");s.setAttribute("mcm2007u.from2004",id(d,"UM_KIHON_MITSUMORI_ID"));s.setAttribute("mcm2007u.from2004Cancel",true);s.setAttribute("mcm2007u.cancelTo","/mcm3007u");u.queryParam("seniMotoKbn",1);break;
   case "mcm1004u": if(!kind.equals("tk")||id(d,"TM_KEIYAKUJIKAN_ID")==null)throw new IllegalStateException(EXPIRED);u.queryParam("tmKeiyakujikanId",d.get("TM_KEIYAKUJIKAN_ID")).queryParam("seniMotoKbn",1);break;
   case "mcm1005u": if(!kind.equals("tk")||id(d,"TK_KEIYAKU_ID")==null)throw new IllegalStateException(EXPIRED);u.queryParam("tkKeiyakuId",d.get("TK_KEIYAKU_ID"));break;
   case "mcm1006u": if(!kind.equals("tk")||id(d,"TORIHIKISAKI_ID")==null||id(d,"PLANT_ID")==null)throw new IllegalStateException(EXPIRED);s.removeAttribute("MCM1006U_DELIVERY");s.removeAttribute("MCM1006U_FORM");u.path("/step1").queryParam("seniMotoKbn",1).queryParam("plantId",d.get("PLANT_ID")).queryParam("torihikisakiId",d.get("TORIHIKISAKI_ID")).queryParam("tmKeiyakujikanIdsStr",d.get("TM_KEIYAKUJIKAN_ID"));for(var e:Map.of("nonyusakiCd","NONYUSAKI_CD","nonyusakiNk","NONYUSAKI_NK","supportId","SUPPORT_ID","plantNk","PLANT_NK","torihikisakiCd","TORIHIKISAKI_CD","torihikisakiNk","TORIHIKISAKI_NK").entrySet())u.queryParam(e.getKey(),d.get(e.getValue()));break;
   default: if(!kind.equals("master"))throw new IllegalStateException(EXPIRED);if(destination.equals("mcm0011u"))u.queryParam("nonyusakiId",d.get("NONYUSAKI_ID")).queryParam("returnUrl","/mcm3007u");else u.queryParam("plantId",d.get("PLANT_ID")).queryParam(destination.equals("mcm0013u")?"returnTo0012Url":"returnUrl","/mcm3007u");
  }
  return "redirect:"+u.build().encode().toUriString();
 }catch(IllegalStateException e){ra.addFlashAttribute("error",e.getMessage());}catch(Exception e){logger.error("3007U遷移失敗",e);ra.addFlashAttribute("error","遷移先を確認できませんでした。再検索してください。");}return "redirect:/mcm3007u";}}
 private static void clear(HttpSession s,String prefix){var keys=Collections.list(s.getAttributeNames());for(String key:keys)if(key.startsWith(prefix))s.removeAttribute(key);}
 @Override protected String getScreenTitle(){return "契約・解約";}
 @Override protected String getFunctionId(){return "MCM3007U";}
}
