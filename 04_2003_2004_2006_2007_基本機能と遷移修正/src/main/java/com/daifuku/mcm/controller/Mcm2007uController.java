package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.List;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpSession;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.common.CustomerScreenSupport;
import com.daifuku.mcm.form.Mcm2006uForm;
import com.daifuku.mcm.form.Mcm2007uForm;
import com.daifuku.mcm.service.Mcm2006uService;
import com.daifuku.mcm.service.Mcm2007uService;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;

/** VB版の見積・ブランド選択→機器選択→契約内容への受け渡し。DB更新は契約の登録時に行う。 */
@Controller
@RequestMapping("/mcm2007u")
public class Mcm2007uController extends BaseController {
    @Autowired private Mcm2007uService service;
    @Autowired private com.daifuku.mcm.service.Mcm2004uService permissions;
    @Autowired private Mcm2006uService mcm2006uService;
    private static final String SCREEN = "mcm2007u";
    private static final String FORM = "MCM2007U_FORM";

    private Mcm2006uForm parent(HttpSession session) {
        if (!(session.getAttribute("MCM2006U_FORM") instanceof Mcm2006uForm parent)
                || parent.getPlantId() == null) throw new IllegalStateException(EXPIRED);
        if (!update(session,"MCM2006U",hasUpdateAuthority(),()->permissions.getAuthority(getLoginUserId(),"MCM2006U"))
                || (!mcm2006uService.isUnlocked(parent,session)&&mcm2006uService.isReadOnly(parent.getJotai(), parent.getSeniMotoKbn())))
            throw new IllegalStateException(DENIED);
        return parent;
    }

    @GetMapping
    public String step1(@RequestParam(required = false) Integer seniMotoKbn, Model model,
                        HttpSession session, RedirectAttributes ra) {
        synchronized (session) {
            try {
                var parent = parent(session);
                var form = (Mcm2007uForm) session.getAttribute(FORM);
                if (form == null) {
                    form = new Mcm2007uForm();
                    BeanUtils.copyProperties(parent, form, "seniMotoKbn");
                    form.setSeniMotoKbn(parent.getUkKeiyakuId() == null ? 1 : 2);
                    service.loadStep1(form);
                    if (session.getAttribute("mcm2007u.from2004") instanceof BigDecimal id) {
                        service.selectInitialEstimate(form, id);
                    } else if (parent.getSelectedTab() != null) {
                        service.selectExisting(form, parent.getSelectedTab());
                        Object start = session.getAttribute("mcm2007u.changeStart");
                        form.setKaisiDt(start == null ? parent.getSelectedTab().getKaisiDt() : start.toString());
                        form.setSyuryoDt(format(date(form.getKaisiDt()).plusYears(1).minusDays(1)));
                    }
                    session.setAttribute("mcm2007u.parentToken", token(session, "mcm2006u"));
                    rotate(session, SCREEN);
                    session.setAttribute(FORM, form);
                }
                return display(form, session, model, false);
            } catch (IllegalStateException e) { ra.addFlashAttribute("error", e.getMessage()); return "redirect:/mcm2004u"; }
        }
    }

    @PostMapping("/step2")
    public String step2(HttpServletRequest request, HttpSession session, Model model, RedirectAttributes ra) {
        synchronized (session) {
            try {
                parent(session); check(session, SCREEN, request);
                var form = stored(session);
                var input = copy(form);
                if (input.getUkKeiyakuId() == null) bind(request, input, "kaisiDt", "syuryoDt");
                else bind(request, input, "syuryoDt");
                if(request.getParameter("selectedEstimateIndex")!=null){
                    int index;
                    try{index=Integer.parseInt(request.getParameter("selectedEstimateIndex"));}catch(NumberFormatException ex){throw new IllegalStateException(EXPIRED);}
                    if(index < -1 || index >= input.getMitsumoriRows().size())throw new IllegalStateException(EXPIRED);
                    input.setSelectedEstimateIndex(index);
                }
                var brands = input.getMitsumoriRows().stream().flatMap(m -> m.getBrandRows().stream()).toList();
                var selected = selections(request, "brandSelection", brands.size());
                for (int i = 0; i < brands.size(); i++) brands.get(i).setCheckFlg(selected.contains(i));
                var errors = service.validateStep1(input);
                session.setAttribute(FORM, input);
                if (!errors.isEmpty()) { model.addAttribute("errors", errors); return display(input, session, model, false); }
                if(service.hasDuplicateCandidates(input))ra.addFlashAttribute("infoMessage","重複した個体を含む見積が選択されています。機器・個体の選定内容を確認してください。");
                boolean sameSelection = form.getMitsumoriRows().stream().flatMap(m -> m.getBrandRows().stream()).map(b -> b.isCheckFlg()).toList()
                    .equals(brands.stream().map(b -> b.isCheckFlg()).toList());
                if (!sameSelection || session.getAttribute("MCM2007U_STEP1_FORM") == null) {
                    service.loadStep2(input);
                    if (input.getUkKeiyakuId() != null) service.selectExistingEquipment(input, parent(session).getSelectedTab());
                }
                session.setAttribute("MCM2007U_STEP1_FORM", Boolean.TRUE);
                rotate(session, SCREEN);
                return "redirect:/mcm2007u/step2";
            } catch (IllegalStateException e) { ra.addFlashAttribute("error", e.getMessage()); return "redirect:/mcm2007u"; }
        }
    }

    @GetMapping("/step2")
    public String step2Get(HttpSession session, Model model, RedirectAttributes ra) {
        synchronized(session){
            try{
                parent(session);
                if (!(session.getAttribute(FORM) instanceof Mcm2007uForm form)
                        || session.getAttribute("MCM2007U_STEP1_FORM") == null) return "redirect:/mcm2007u";
                if(!token(session,"mcm2006u").equals(session.getAttribute("mcm2007u.parentToken")))throw new IllegalStateException(EXPIRED);
                return display(form, session, model, true);
            }catch(IllegalStateException ex){ra.addFlashAttribute("error",ex.getMessage());return "redirect:/mcm2006u";}
        }
    }

    @PostMapping({"/apply", "/previous"})
    public String apply(HttpServletRequest request, HttpSession session, RedirectAttributes ra) {
        synchronized (session) {
            try {
                parent(session); check(session, SCREEN, request);
                if (!token(session, "mcm2006u").equals(session.getAttribute("mcm2007u.parentToken"))
                        || session.getAttribute("MCM2007U_STEP1_FORM") == null) throw new IllegalStateException(EXPIRED);
                var form = copy(stored(session));
                var k = selections(request, "koseiSelection", form.getKoseiRows().size());
                var m = selections(request, "meisaiSelection", form.getMeisaiRows().size());
                var t = selections(request, "kotaiSelection", form.getKotaiRows().size());
                for (int i=0;i<form.getKoseiRows().size();i++) form.getKoseiRows().get(i).setCheckFlg(k.contains(i));
                for (int i=0;i<form.getMeisaiRows().size();i++) form.getMeisaiRows().get(i).setCheckFlg(m.contains(i));
                for (int i=0;i<form.getKotaiRows().size();i++) form.getKotaiRows().get(i).setCheckFlg(t.contains(i));
                bindDisplayIndex(request,"selectedBrandIndex",form.getStep2Brands().size(),form::setSelectedBrandIndex);
                bindDisplayIndex(request,"selectedKoseiIndex",form.getKoseiRows().size(),form::setSelectedKoseiIndex);
                session.setAttribute(FORM, form);
                if (request.getRequestURI().endsWith("/previous")) return "redirect:/mcm2007u";
                var errors = service.validateStep2(form);
                if (!errors.isEmpty()) { ra.addFlashAttribute("errors", errors); return "redirect:/mcm2007u/step2"; }
                var parent = copy(parent(session));
                service.prepareDeliveryQuantities(form);
                var pending = mcm2006uService.buildKikanFromDelivery(form);
                mcm2006uService.applyPendingKikan(parent, pending);
                mcm2006uService.validatePeriods(parent);
                session.setAttribute("MCM2006U_FORM", parent);
                rotate(session, "mcm2006u");
                clear(session);
                return "redirect:/mcm2006u";
            } catch (IllegalStateException e) { ra.addFlashAttribute("error", e.getMessage()); return "redirect:/mcm2007u/step2"; }
        }
    }

    @GetMapping("/back")
    public String back(HttpSession session) {
        String cancelTo="/mcm3007u".equals(session.getAttribute("mcm2007u.cancelTo"))?"/mcm3007u":"/mcm2004u";
        boolean fromSearch = Boolean.TRUE.equals(session.getAttribute("mcm2007u.from2004Cancel"));
        clear(session);
        if (fromSearch) { session.removeAttribute("MCM2006U_FORM"); return "redirect:"+cancelTo; }
        return "redirect:/mcm2006u";
    }
    private Mcm2007uForm stored(HttpSession s) {
        if (s.getAttribute(FORM) instanceof Mcm2007uForm form) return form;
        throw new IllegalStateException(EXPIRED);
    }
    private String display(Mcm2007uForm form, HttpSession session, Model model, boolean second) {
        setCommonAttributes(model, session);
        var parent=(Mcm2006uForm)session.getAttribute("MCM2006U_FORM");
        model.addAttribute("contractNo",parent==null?"":parent.getExtra().getOrDefault("KEIYAKU_NO",""));
        if(!second){
            var supportHours=new java.util.LinkedHashMap<String,String>();
            for(var row:mcm2006uService.hours())supportHours.put(new BigDecimal(row.get("TORIHOSYUJIKAN_ID").toString()).stripTrailingZeros().toPlainString(),row.get("label").toString());
            model.addAttribute("supportHours",supportHours);
        }
        model.addAttribute("form", form);
        model.addAttribute("workflowToken", token(session, SCREEN));
        model.addAttribute("brands", form.getMitsumoriRows().stream().flatMap(m -> m.getBrandRows().stream()).toList());
        return second ? "mcm2007u/step2" : "mcm2007u/index";
    }
    private void bindDisplayIndex(HttpServletRequest request,String name,int size,java.util.function.IntConsumer setter){
        if(request.getParameter(name)==null)return;
        try{int index=Integer.parseInt(request.getParameter(name));if(index < -1 || index >= size)throw new IllegalStateException(EXPIRED);setter.accept(index);}catch(NumberFormatException ex){throw new IllegalStateException(EXPIRED);}
    }
    private void clear(HttpSession s) {
        s.removeAttribute(FORM); s.removeAttribute("MCM2007U_STEP1_FORM");
        for (String key : List.of("from2004", "from2004Cancel", "parentToken", "changeStart", "cancelTo")) s.removeAttribute(SCREEN + "." + key);
        rotate(s, SCREEN);
    }
    @org.springframework.web.bind.annotation.ExceptionHandler(org.springframework.dao.DataAccessException.class)
    public String failure(Exception e, RedirectAttributes ra) {
        logger.error("MCM2007U 読み込み失敗", e);
        ra.addFlashAttribute("error", "画面の情報を取得できませんでした。もう一度操作してください。");
        return "redirect:/mcm2006u";
    }
    @Override protected String getScreenTitle() { return "ユーザ契約内容変更"; }
    @Override protected String getFunctionId() { return "MCM2007U"; }
}
