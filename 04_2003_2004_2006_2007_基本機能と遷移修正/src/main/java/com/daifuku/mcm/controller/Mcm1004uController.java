package com.daifuku.mcm.controller;
import java.math.BigDecimal;
import java.util.*;
import jakarta.servlet.http.*;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import org.springframework.web.bind.ServletRequestDataBinder;
import org.springframework.format.support.DefaultFormattingConversionService;
import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.form.Mcm1004uForm;
import com.daifuku.mcm.service.Mcm1004uService;
import lombok.RequiredArgsConstructor;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;

/** 取引先見積内容。編集値のみをセッション上の対象行へ反映する。 */
@Controller @RequestMapping("/mcm1004u") @RequiredArgsConstructor
public class Mcm1004uController extends BaseController {
 private final Mcm1004uService service;
 private static final String FORM="MCM1004U_FORM",SCREEN="mcm1004u",OWNER="mcm1004u.owner";
 private record Unlock(String user,BigDecimal id,int mode) implements java.io.Serializable {}
 private boolean updateAllowed(){return hasUpdateAuthority()&&service.canUpdate(getLoginUserId());}
 private boolean unlocked(HttpSession s,Mcm1004uForm f){Object grant=s.getAttribute("mcm1004u.unlock");boolean allowed=grant instanceof Unlock u&&u.user().equals(getLoginUserId())&&same(u.id(),f.getTmKeiyakujikanId())&&u.mode()==f.getSeniMotoKbn()&&updateAllowed()&&service.canUnlock(getLoginUserId());if(!allowed)s.removeAttribute("mcm1004u.unlock");return allowed;}
 private Mcm1004uForm form(HttpSession s){if(!(s.getAttribute(FORM) instanceof Mcm1004uForm f)||!getLoginUserId().equals(s.getAttribute(OWNER)))throw new IllegalStateException(EXPIRED);return f;}
 private void store(HttpSession s,Mcm1004uForm f){s.setAttribute(FORM,f);s.setAttribute(OWNER,getLoginUserId());}
 private Mcm1004uForm fresh(HttpSession s,BigDecimal id,int mode,boolean release){var f=service.loadScreen(id,mode,release);f.setViewFlg(f.isViewFlg()&&updateAllowed());store(s,f);rotate(s,SCREEN);return f;}
 @GetMapping
 public String index(@RequestParam(value="tmKeiyakujikanId",required=false) BigDecimal id,@RequestParam(value="seniMotoKbn",required=false) Integer mode,@RequestParam(value="returnTo",required=false) String returnTo,Model model,HttpSession session){synchronized(session){
  try{
   if(!hasInspectionAuthority()||!service.canView(getLoginUserId(),"MCM1004U"))throw new IllegalStateException("この画面を参照する権限がありません。");
   Mcm1004uForm previous=session.getAttribute(FORM) instanceof Mcm1004uForm f?f:null;
   if(mode==null)mode=id==null&&previous!=null?previous.getSeniMotoKbn():1;
   if(id==null&&previous!=null)id=previous.getTmKeiyakujikanId();if(id==null)throw new IllegalStateException("検索画面から見積依頼を選択してください。");
   if(returnTo!=null||previous==null||mode!=previous.getSeniMotoKbn()||!same(id,previous.getTmKeiyakujikanId())||!getLoginUserId().equals(session.getAttribute(OWNER))){session.removeAttribute("mcm1004u.unlock");session.setAttribute("mcm1004u.returnTo","/mcm3007u".equals(returnTo)?"/mcm3007u":"/mcm1003u");previous=fresh(session,id,mode,false);}
   boolean release=unlocked(session,previous);previous.setLockReleaseFlg(release);previous.setViewFlg(updateAllowed()&&service.loadScreen(id,previous.getSeniMotoKbn(),release).isViewFlg());
   model.addAttribute("form",previous);model.addAttribute("canUnlock",updateAllowed()&&service.canUnlock(getLoginUserId())&&previous.getSeniMotoKbn()!=2);model.addAttribute("adminUnlocked",unlocked(session,previous));
  }catch(IllegalStateException e){model.addAttribute("error",e.getMessage());empty(model);}catch(Exception e){logger.error("1004U読込失敗",e);model.addAttribute("error","見積を読み込めませんでした。検索画面から開き直してください。");empty(model);}
  model.addAttribute("workflowToken",token(session,SCREEN));model.addAttribute("methods",Mcm1004uService.METHODS);model.addAttribute("services",Mcm1004uService.SERVICES);model.addAttribute("days",Mcm1004uService.DAYS);return "mcm1004u/index";
 }}
 private void empty(Model model){var f=new Mcm1004uForm();f.setDataNotFound(true);model.addAttribute("form",f);model.addAttribute("canUnlock",false);}
 private Mcm1004uForm input(HttpServletRequest request,HttpSession session){
  check(session,SCREEN,request);var saved=form(session);if(!updateAllowed())throw new IllegalStateException("更新権限がありません。");
  boolean release=unlocked(session,saved);var latest=service.loadScreen(saved.getTmKeiyakujikanId(),saved.getSeniMotoKbn(),release);
  if(!latest.isViewFlg())throw new IllegalStateException("この見積は参照専用です。");if(!Objects.equals(saved.getRevision(),latest.getRevision()))throw new IllegalStateException("他の操作で見積が更新されました。再読み込みして内容を確認してください。");
  var f=copy(saved);f.setLockReleaseFlg(release);var fields=new ArrayList<String>();fields.add("selectedIndex");
  for(int t=0;t<f.getTabs().size();t++){
   String p="tabs["+t+"].";for(String key:List.of("kaisiDt","syuryoDt","tmMitsumoriNo","kaitoDt","biko","syusseinebikiKin"))fields.add(p+key);
   for(int r=0;r<f.getTabs().get(t).getTankaRows().size();r++)for(String key:List.of("hyojunKin","sikiriKin","packFlg","keiyakunaiyo","torihosyujikanId","tenkenumu","hosyuhoho","servicekeitai","biko"))fields.add(p+"tankaRows["+r+"]."+key);
   for(int r=0;r<f.getTabs().get(t).getTenkenRows().size();r++)for(String key:List.of("tenkenkaisu","tenkenkanoyobi","yakantaioumu","biko"))fields.add(p+"tenkenRows["+r+"]."+key);
  }
  var binder=new ServletRequestDataBinder(f);binder.setAutoGrowNestedPaths(false);binder.setConversionService(new DefaultFormattingConversionService());binder.setAllowedFields(fields.toArray(String[]::new));binder.bind(request);
  if(binder.getBindingResult().hasErrors())throw new IllegalStateException("日付・数値の入力形式を確認してください。");if(f.getSelectedIndex()<0||f.getSelectedIndex()>=f.getTabs().size())throw new IllegalStateException(EXPIRED);
  for(var t:f.getTabs())Mcm1004uService.recalculate(t);return f;
 }
 private String fail(Exception e,RedirectAttributes ra){if(e instanceof IllegalStateException)ra.addFlashAttribute("error",e.getMessage());else{logger.error("1004U操作失敗",e);ra.addFlashAttribute("error","処理を完了できませんでした。入力内容を確認してください。");}return "redirect:/mcm1004u";}
 @PostMapping("/update") public String update(HttpServletRequest request,HttpSession session,RedirectAttributes ra){synchronized(session){try{var f=input(request,session);store(session,f);service.updateAll(f,getLoginUserId());var loaded=fresh(session,f.getTmKeiyakujikanId(),f.getSeniMotoKbn(),unlocked(session,f));loaded.setSelectedIndex(Math.min(f.getSelectedIndex(),loaded.getTabs().size()-1));ra.addFlashAttribute("message","登録を完了しました。");return "redirect:/mcm1004u";}catch(Exception e){return fail(e,ra);}}}
 @PostMapping("/addTab") public String addTab(HttpServletRequest request,HttpSession session,RedirectAttributes ra){synchronized(session){try{var f=input(request,session);store(session,f);service.addNewTab(f);rotate(session,SCREEN);return "redirect:/mcm1004u";}catch(Exception e){return fail(e,ra);}}}
 @PostMapping("/dates") public String dates(@RequestParam String field,HttpServletRequest request,HttpSession session,RedirectAttributes ra){synchronized(session){try{var f=input(request,session);try{service.adjustDates(f,f.getSelectedIndex(),field);}catch(IllegalStateException e){var before=form(session);for(int i=0;i<f.getTabs().size();i++){f.getTabs().get(i).setKaisiDt(before.getTabs().get(i).getKaisiDt());f.getTabs().get(i).setSyuryoDt(before.getTabs().get(i).getSyuryoDt());}store(session,f);throw e;}store(session,f);rotate(session,SCREEN);return "redirect:/mcm1004u";}catch(Exception e){return fail(e,ra);}}}
 @PostMapping("/lockRelease") public String lockRelease(HttpServletRequest request,HttpSession session,RedirectAttributes ra){synchronized(session){try{check(session,SCREEN,request);var f=form(session);if(!updateAllowed()||!service.canUnlock(getLoginUserId())||f.getSeniMotoKbn()==2)throw new IllegalStateException("管理者用ロック解除の権限がありません。");session.setAttribute("mcm1004u.unlock",new Unlock(getLoginUserId(),f.getTmKeiyakujikanId(),f.getSeniMotoKbn()));fresh(session,f.getTmKeiyakujikanId(),f.getSeniMotoKbn(),true);ra.addFlashAttribute("message","最新の情報を読み込み、編集制限を解除しました。");return "redirect:/mcm1004u";}catch(Exception e){return fail(e,ra);}}}
 @PostMapping("/reload") public String reload(HttpServletRequest request,HttpSession session,RedirectAttributes ra){synchronized(session){try{check(session,SCREEN,request);var f=form(session);session.removeAttribute("mcm1004u.unlock");fresh(session,f.getTmKeiyakujikanId(),f.getSeniMotoKbn(),false);return "redirect:/mcm1004u";}catch(Exception e){return fail(e,ra);}}}
 @PostMapping("/inquiry") public String inquiry(@RequestParam int rowIndex,HttpServletRequest request,HttpSession session,RedirectAttributes ra){synchronized(session){try{
  check(session,SCREEN,request);if(!hasInspectionAuthority()||!service.canView(getLoginUserId(),"MCM1004U")||!service.canView(getLoginUserId(),"MCM1007U"))throw new IllegalStateException("単価精査の参照権限がありません。");var saved=form(session);Mcm1004uForm f=saved;
  if(saved.isViewFlg()){f=input(request,session);store(session,f);}else{int i=Integer.parseInt(request.getParameter("selectedIndex"));if(i<0||i>=f.getTabs().size())throw new IllegalStateException(EXPIRED);f.setSelectedIndex(i);}
  var rows=f.getTabs().get(f.getSelectedIndex()).getTankaRows();if(rowIndex<0||rowIndex>=rows.size())throw new IllegalStateException(EXPIRED);var row=rows.get(rowIndex);if(row.getAtsukaikikiId()==null)throw new IllegalStateException("取扱機器が登録されていません。");
  session.setAttribute("mcm1007u.atsukaikikiId",row.getAtsukaikikiId());session.setAttribute("mcm1007u.torihikisakiId",f.getTorihikisakiId());session.setAttribute("mcm1007u.keiyakujikantai",f.getKeiyakujikantai());return "redirect:/mcm1007u";
 }catch(Exception e){return fail(e,ra);}}}
 @GetMapping("/back") public String back(HttpSession session){synchronized(session){String target="/mcm3007u".equals(session.getAttribute("mcm1004u.returnTo"))?"/mcm3007u":"/mcm1003u";session.removeAttribute(FORM);session.removeAttribute(OWNER);session.removeAttribute("mcm1004u.unlock");session.removeAttribute("mcm1004u.returnTo");rotate(session,SCREEN);return "redirect:"+target;}}
 @Override protected String getScreenTitle(){return "取引先見積内容";}
 @Override protected String getFunctionId(){return "MCM1004U";}
}
