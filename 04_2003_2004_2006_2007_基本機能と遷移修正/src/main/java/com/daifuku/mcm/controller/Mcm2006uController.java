package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.Map;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpSession;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.form.Mcm2006uForm;
import com.daifuku.mcm.service.Mcm2006uService;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;

/** VB版の契約読込・編集・登録・変更画面への受け渡し。 */
@Controller
@RequestMapping("/mcm2006u")
public class Mcm2006uController extends BaseController {
    @Autowired private Mcm2006uService service;
    @Autowired private com.daifuku.mcm.service.Mcm2004uService permissions;
    private static final String FORM = "MCM2006U_FORM", SCREEN = "mcm2006u";

    @GetMapping
    public String index(@RequestParam(required=false) BigDecimal ukKeiyakuId,
                        @RequestParam(required=false) Integer seniMotoKbn,
                        @RequestParam(required=false) BigDecimal kikanId,
                        @RequestParam(required=false) Integer tabIndex,
                        Model model, HttpSession session, RedirectAttributes ra) {
        synchronized (session) {
            try {
                if (session.getAttribute("MCM2006U_DELIVERY") instanceof Map<?,?> delivery) {
                    ukKeiyakuId = new BigDecimal(delivery.get("ukKeiyakuId").toString());
                    seniMotoKbn = Integer.valueOf(delivery.get("seniMotoKbn").toString());
                    session.setAttribute("mcm2006u.returnTo", "/mcm3007u");
                    session.removeAttribute("MCM2006U_DELIVERY");
                }
                var form = (Mcm2006uForm) session.getAttribute(FORM);
                if (ukKeiyakuId != null) {
                    service.clearUnlock(session);
                    form = service.load(ukKeiyakuId);
                    form.setSeniMotoKbn(seniMotoKbn == null ? 0 : seniMotoKbn);
                    session.setAttribute("mcm2006u.ukKeiyakuId", ukKeiyakuId);
                    session.setAttribute("mcm2006u.seniMotoKbn", form.getSeniMotoKbn());
                    rotate(session, SCREEN);
                } else if (form == null && session.getAttribute("mcm2006u.ukKeiyakuId") instanceof BigDecimal id) {
                    form = service.load(id);
                    Object mode = session.getAttribute("mcm2006u.seniMotoKbn");
                    form.setSeniMotoKbn(mode instanceof Integer i ? i : 0);
                }
                if (form == null || form.getPlantId() == null) throw new IllegalStateException(EXPIRED);
                if (kikanId != null) {
                    final BigDecimal wanted = kikanId;
                    if (form.getKikanTabs().stream().noneMatch(t -> same(wanted, t.getUkKikanId()))) throw new IllegalStateException(EXPIRED);
                    form.setSelectedKikanId(kikanId);
                }
                if (tabIndex != null) {
                    if (tabIndex < 0 || tabIndex >= form.getKikanTabs().size()) throw new IllegalStateException(EXPIRED);
                    form.setSelectedKikanId(form.getKikanTabs().get(tabIndex).getUkKikanId());
                }
                if ((tabIndex != null || kikanId != null) && form.getSelectedTab() != null) {
                    var tab=form.getSelectedTab();
                    form.setNonyubusyoNk(tab.getNonyubusyoNk());form.setNonyutantosyaNk(tab.getNonyutantosyaNk());
                    form.setNonyutelNo(tab.getNonyutelNo());form.setNonyufaxNo(tab.getNonyufaxNo());
                }
                com.daifuku.mcm.common.Mcm2006uFields.initialize(form);
                session.setAttribute(FORM, form);
                boolean unlocked=service.isUnlocked(form,session);
                model.addAttribute("adminUnlocked",unlocked);
                model.addAttribute("canUnlock",form.getUkKeiyakuId()!=null&&form.getSeniMotoKbn()==0&&service.canReleaseLock(session));
                model.addAttribute("periodReadOnly",form.getSelectedTab()==null||readOnly(form,session)||(!unlocked&&service.isTabReadOnly(form.getSelectedTab())));
                model.addAttribute("canChange",form.getSelectedTab()!=null&&!readOnly(form,session)&&(unlocked||form.getSelectedTab().getTabFlg()!=2));
                model.addAttribute("tenpoChoices",service.tenpoChoices());model.addAttribute("hours",service.hours());
                setCommonAttributes(model, session);
                model.addAttribute("form", form);
                model.addAttribute("readOnly", readOnly(form, session));
                model.addAttribute("selectedTab", form.getSelectedTab());
                model.addAttribute("selectedIndex", form.getKikanTabs().indexOf(form.getSelectedTab()));
                model.addAttribute("workflowToken", token(session, SCREEN));
                return "mcm2006u/index";
            } catch (IllegalStateException e) { ra.addFlashAttribute("error", e.getMessage()); }
            catch (org.springframework.dao.DataAccessException e) {
                logger.error("MCM2006U 読み込み失敗", e);
                ra.addFlashAttribute("error", "契約情報を読み込めませんでした。再検索してください。");
            }
            return "redirect:/mcm2004u";
        }
    }


    @PostMapping("/lock-release")
    public String releaseLock(HttpServletRequest request,HttpSession session,RedirectAttributes ra){synchronized(session){
        try{check(session,SCREEN,request);var current=(Mcm2006uForm)session.getAttribute(FORM);
            var latest=service.releaseLock(current,session);session.setAttribute(FORM,latest);
            session.removeAttribute("MCM2007U_FORM");session.removeAttribute("MCM2007U_STEP1_FORM");session.removeAttribute("mcm2007u.parentToken");
            rotate(session,SCREEN);ra.addFlashAttribute("message","契約を読み直し、管理者用の編集制限を解除しました。");
        }catch(IllegalStateException ex){ra.addFlashAttribute("error",ex.getMessage());}
        catch(Exception ex){service.clearUnlock(session);logger.error("MCM2006U 編集制限解除失敗",ex);ra.addFlashAttribute("error","編集制限を解除できませんでした。再検索してください。");}
        return "redirect:/mcm2006u";
    }}
    @PostMapping("/period") @org.springframework.web.bind.annotation.ResponseBody
    public Map<String,Object> period(HttpServletRequest request,HttpSession session){synchronized(session){
        // 古い画面へ新しいトークンや編集中の期間を返さない。
        try{check(session,SCREEN,request);}catch(IllegalStateException ex){return Map.of("error",ex.getMessage(),"periods",java.util.List.of());}
        try{var form=input(request,session);session.setAttribute(FORM,form);rotate(session,SCREEN);return periodResponse(form,session,"");}
        catch(IllegalStateException ex){return periodResponse((Mcm2006uForm)session.getAttribute(FORM),session,ex.getMessage());}
    }}
    private Map<String,Object> periodResponse(Mcm2006uForm form,HttpSession session,String error){
        var result=new java.util.LinkedHashMap<String,Object>();result.put("error",error);result.put("workflowToken",token(session,SCREEN));
        result.put("periods",form==null?java.util.List.of():form.getKikanTabs().stream().map(t->Map.of("start",t.getKaisiDt(),"end",t.getSyuryoDt(),"label",t.getTabLabel())).toList());
        result.put("nextRenewal",form==null?"":java.util.Objects.toString(form.getJikaikosinDt(),""));return result;
    }
    @PostMapping("/save")
    public String save(HttpServletRequest request, HttpSession session, RedirectAttributes ra) {
        synchronized (session) {
            try {
                var form = input(request, session);
                // 保存失敗時のID採番や途中処理が、編集中のフォームに残らないようにする。
                session.setAttribute(FORM, form);
                var saving = copy(form);
                service.save(saving, getLoginUserId());
                session.setAttribute("mcm2006u.ukKeiyakuId", saving.getUkKeiyakuId());
                session.setAttribute("mcm2006u.seniMotoKbn", form.getSeniMotoKbn() == 1 ? 0 : form.getSeniMotoKbn());
                session.removeAttribute(FORM);
                rotate(session, SCREEN);
                session.setAttribute("mcm2004u.refresh", true);
                ra.addFlashAttribute("message", "登録を完了しました。");
            } catch (IllegalStateException e) { ra.addFlashAttribute("error", e.getMessage()); }
            catch (Exception e) {
                logger.error("MCM2006U 登録失敗", e);
                ra.addFlashAttribute("error", "登録できませんでした。入力内容を保持しています。");
            }
            return "redirect:/mcm2006u";
        }
    }

    @PostMapping("/tab")
    public String tab(@RequestParam int tabIndex,HttpServletRequest request,HttpSession session,RedirectAttributes ra) {
        synchronized(session) {
            try {
                var form=input(request,session);
                if(tabIndex<0 || tabIndex>=form.getKikanTabs().size())throw new IllegalStateException(EXPIRED);
                form.setSelectedKikanId(form.getKikanTabs().get(tabIndex).getUkKikanId());
                var selected=form.getSelectedTab();
                form.setNonyubusyoNk(selected.getNonyubusyoNk());form.setNonyutantosyaNk(selected.getNonyutantosyaNk());
                form.setNonyutelNo(selected.getNonyutelNo());form.setNonyufaxNo(selected.getNonyufaxNo());
                session.setAttribute(FORM,form);rotate(session,SCREEN);
            }catch(IllegalStateException ex){ra.addFlashAttribute("error",ex.getMessage());}
            return "redirect:/mcm2006u";
        }
    }
    @PostMapping("/change")
    public String change(HttpServletRequest request, HttpSession session, RedirectAttributes ra) {
        synchronized (session) {
            try {
                var form = input(request, session);
                var tab = form.getSelectedTab();
                if (tab == null || (tab.getTabFlg()==2&&!service.isUnlocked(form,session))) throw new IllegalStateException(DENIED);
                var start = date(request.getParameter("changeStart"));
                if (start.isBefore(date(tab.getKaisiDt())) || start.isAfter(date(tab.getSyuryoDt()).plusDays(1)))
                    throw new IllegalStateException("変更摘要開始日は選択期間の開始日から終了日の翌日までで指定してください。");
                session.setAttribute(FORM, form);
                session.removeAttribute("MCM2007U_FORM"); session.removeAttribute("MCM2007U_STEP1_FORM");
                session.removeAttribute("mcm2007u.from2004"); session.removeAttribute("mcm2007u.from2004Cancel");
                session.setAttribute("mcm2007u.changeStart", format(start));
                return "redirect:/mcm2007u";
            } catch (IllegalStateException e) { ra.addFlashAttribute("error", e.getMessage()); return "redirect:/mcm2006u"; }
        }
    }
    private Mcm2006uForm input(HttpServletRequest request, HttpSession session) {
        check(session, SCREEN, request);
        if (!(session.getAttribute(FORM) instanceof Mcm2006uForm saved)) throw new IllegalStateException(EXPIRED);
        if (readOnly(saved, session)) throw new IllegalStateException(DENIED);
        var form = copy(saved);
        bind(request, form, "nonyubusyoNk", "nonyutantosyaNk", "nonyutelNo", "nonyufaxNo");
        var tab = form.getSelectedTab();
        boolean unlocked=service.isUnlocked(form,session);
        if (tab != null && (unlocked||!service.isTabReadOnly(tab))) {
            String index = request.getParameter("selectedIndex");
            if (index == null || !index.equals(String.valueOf(form.getKikanTabs().indexOf(tab)))) throw new IllegalStateException(EXPIRED);
            bind(request, tab, "biko", "hosyuhoho", "hosyuGkin", "keiyakujikantai", "iraitenpoId", "iraitantoNk");
            tab.setNonyubusyoNk(form.getNonyubusyoNk());tab.setNonyutantosyaNk(form.getNonyutantosyaNk());
            tab.setNonyutelNo(form.getNonyutelNo());tab.setNonyufaxNo(form.getNonyufaxNo());
        }
        service.adjustPeriodDates(form,request.getParameter("kaisiDt"),request.getParameter("syuryoDt"),unlocked);
        com.daifuku.mcm.common.Mcm2006uFields.bind(request,form,tab!=null&&(unlocked||!service.isTabReadOnly(tab)));
        return form;
    }
    @PostMapping("/rows") public String rows(@RequestParam String table,@RequestParam String operation,@RequestParam(required=false) Integer rowIndex,HttpServletRequest request,HttpSession session,RedirectAttributes ra){synchronized(session){try{
        var form=input(request,session);var tab=form.getSelectedTab();
        if(!java.util.Set.of("seiban","inspection").contains(table)||!java.util.Set.of("add","remove").contains(operation))throw new IllegalStateException(EXPIRED);
        if(table.equals("inspection")&&(tab==null||(!service.isUnlocked(form,session)&&service.isTabReadOnly(tab))))throw new IllegalStateException(DENIED);
        if(operation.equals("add")){
            if(table.equals("seiban")){var row=new Mcm2006uForm.SeibanRowForm();row.setKaisiDt(tab==null?form.getKeiyakuDt():tab.getKaisiDt());row.setSyuryoDt(tab==null?form.getJikaikosinDt():tab.getSyuryoDt());row.getExtra().put("KAISI_DT",row.getKaisiDt());row.getExtra().put("SYURYO_DT",row.getSyuryoDt());row.setExtraEdited(true);form.getSeibanRows().add(row);}
            else{var row=new Mcm2006uForm.TenkenRowForm();row.setExtraEdited(true);row.getExtra().put("NAIYO","");row.getExtra().put("BIKO","");for(int i=1;i<=12;i++)row.getExtra().put(String.format("M%02d",i),"0");tab.getTenkenRows().add(row);}
        }else{
            if(rowIndex==null||rowIndex<0)throw new IllegalStateException("削除する行を選択してください。");
            if(table.equals("seiban")){if(rowIndex>=form.getSeibanRows().size())throw new IllegalStateException(EXPIRED);form.getSeibanRows().get(rowIndex).setRemoved(true);}
            else{if(rowIndex>=tab.getTenkenRows().size())throw new IllegalStateException(EXPIRED);tab.getTenkenRows().get(rowIndex).setRemoved(true);}
        }
        session.setAttribute(FORM,form);rotate(session,SCREEN);
    }catch(IllegalStateException ex){ra.addFlashAttribute("error",ex.getMessage());}return "redirect:/mcm2006u";}}
    private boolean readOnly(Mcm2006uForm f, HttpSession s) {
        return !update(s,"MCM2006U",hasUpdateAuthority(),()->permissions.getAuthority(getLoginUserId(),"MCM2006U")) || (!service.isUnlocked(f,s)&&service.isReadOnly(f.getJotai(), f.getSeniMotoKbn()));
    }
    @GetMapping("/back")
    public String back(HttpSession session) {
        String destination = "/mcm3007u".equals(session.getAttribute("mcm2006u.returnTo")) ? "/mcm3007u" : "/mcm2004u";
        service.clearUnlock(session);session.removeAttribute(FORM);
        session.removeAttribute("mcm2006u.ukKeiyakuId"); session.removeAttribute("mcm2006u.pendingKikan");
        rotate(session, SCREEN);
        return "redirect:" + destination;
    }
    @Override protected String getScreenTitle() { return "ユーザ契約内容"; }
    @Override protected String getFunctionId() { return "MCM2006U"; }
}
