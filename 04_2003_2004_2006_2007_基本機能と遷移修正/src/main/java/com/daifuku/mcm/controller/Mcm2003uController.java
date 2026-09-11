package com.daifuku.mcm.controller;

import java.math.BigDecimal;
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
import com.daifuku.mcm.form.Mcm2003uForm;
import com.daifuku.mcm.service.Mcm2003uService;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;

/** VB版の見積登録・発行・申請と基本設定、ブランド詳細、検索画面との遷移。 */
@Controller
@RequestMapping("/mcm2003u")
public class Mcm2003uController extends BaseController {
    @Autowired private Mcm2003uService service;
    @Autowired private com.daifuku.mcm.service.Mcm2004uService permissions;
    @Autowired private com.daifuku.mcm.service.Mcm2003uAttachmentService attachments;
    private static final String FORM="MCM2003U_FORM", SCREEN="mcm2003u", ID="mcm2003u.umKihonMitsumoriId";

    @GetMapping
    public String index(@RequestParam(required=false) BigDecimal brandId, Model model,
                        HttpSession session, RedirectAttributes ra) {
        synchronized(session) {
            try {
                BigDecimal id=(BigDecimal)session.getAttribute(ID);
                var form=(Mcm2003uForm)session.getAttribute(FORM);
                Object entry=session.getAttribute("mcm2003u.seniMotoKbn");
                boolean changed=form==null || !same(id,form.getUmKihonMitsumoriId())
                    || entry instanceof Integer mode && form.getSeniMotoKbn()!=mode;
                if(changed) {
                    if(id==null) { ra.addFlashAttribute("error","見積情報を選択してから開いてください。");return "redirect:/mcm2004u"; }
                    form=service.load(id);
                    form.setSeniMotoKbn(entry instanceof Integer mode ? mode : 3);
                    rotate(session,SCREEN);
                }
                if(Boolean.TRUE.equals(session.getAttribute("mcm2003u.refreshDetails"))) {
                    var latest=service.load(form.getUmKihonMitsumoriId());
                    form.setBrandRows(latest.getBrandRows());form.setKoseiRows(latest.getKoseiRows());form.setMeisaiRows(latest.getMeisaiRows());form.setMitsumoriRows(latest.getMitsumoriRows());
                    session.removeAttribute("mcm2003u.refreshDetails");
                }
                session.setAttribute(FORM,form);
                session.setAttribute("mcm2003u.seniMotoKbn",form.getSeniMotoKbn());
                var selected=form.getBrandRows().stream().filter(b->same(b.getUmKihonBrandId(),brandId)).findFirst()
                    .orElse(form.getBrandRows().isEmpty()?null:form.getBrandRows().get(0));
                setCommonAttributes(model,session);
                model.addAttribute("form",form);model.addAttribute("readOnly",readOnly(form,session));
                model.addAttribute("selectedBrand",selected);
                if(selected!=null)session.setAttribute("mcm2003u.selectedBrandId",selected.getUmKihonBrandId());
                model.addAttribute("workflowToken",token(session,SCREEN));
                final Mcm2003uForm displayed = form;
                model.addAttribute("visibleMeisai",form.getMeisaiRows().stream().filter(m->selected==null || displayed.getKoseiRows().stream()
                    .anyMatch(k->same(k.getUmKihonBrandId(),selected.getUmKihonBrandId())&&same(k.getUmKikikoseiId(),m.getUmKikikoseiId()))).toList());
                return "mcm2003u/index";
            } catch(Exception e) { logger.error("MCM2003U 読み込み失敗",e);ra.addFlashAttribute("error","見積情報を読み込めませんでした。再検索してください。");return "redirect:/mcm2004u"; }
        }
    }

    @PostMapping({"/save","/publish","/sinsei"})
    public String save(HttpServletRequest request,HttpSession session,RedirectAttributes ra) {
        synchronized(session) {
            try {
                check(session,SCREEN,request);
                if(!(session.getAttribute(FORM) instanceof Mcm2003uForm saved))throw new IllegalStateException(EXPIRED);
                if(readOnly(saved,session))throw new IllegalStateException(DENIED);
                var form=copy(saved);
                bind(request,form,"keiyakujikantai","mitsumoriLevel","mitsumorigiken","iraitantoNk","sofutantoNk","mitsumoriJouken","hosyuhoho","biko");
                session.setAttribute(FORM,form);
                String op=request.getRequestURI().substring(request.getRequestURI().lastIndexOf('/')+1);
                if(op.equals("publish"))service.publish(form,getLoginUserId());
                else if(op.equals("sinsei"))service.saveAndApply(form,getLoginUserId());
                else service.save(form,getLoginUserId());
                session.removeAttribute(FORM);rotate(session,SCREEN);
                session.setAttribute("mcm2004u.refresh",true);
                ra.addFlashAttribute("message",op.equals("sinsei")?"申請を完了しました。":op.equals("publish")?"見積を発行しました。":"登録を完了しました。");
            } catch(IllegalStateException e) {ra.addFlashAttribute("error",e.getMessage());}
            catch(Exception e) {logger.error("MCM2003U 更新失敗",e);ra.addFlashAttribute("error","登録できませんでした。入力内容を保持しています。");}
            return "redirect:/mcm2003u";
        }
    }
    @PostMapping("/navigate")
    public String navigate(@RequestParam String destination,@RequestParam(required=false) BigDecimal brandId,
            @RequestParam(required=false) BigDecimal mitsumoriId,HttpServletRequest request,HttpSession session,RedirectAttributes ra) {
        synchronized(session) {
            try {
                check(session,SCREEN,request);
                if(!(session.getAttribute(FORM) instanceof Mcm2003uForm saved))throw new IllegalStateException(EXPIRED);
                if(!readOnly(saved,session)) {
                    var input=copy(saved);
                    bind(request,input,"keiyakujikantai","mitsumoriLevel","mitsumorigiken","iraitantoNk","sofutantoNk","mitsumoriJouken","hosyuhoho","biko");
                    session.setAttribute(FORM,input);
                }
                if(destination.equals("detail"))return detail(brandId,mitsumoriId,session,ra);
                if(destination.equals("brand") && saved.getBrandRows().stream().anyMatch(b->same(b.getUmKihonBrandId(),brandId)))
                    return "redirect:/mcm2003u?brandId="+brandId.toPlainString();
                throw new IllegalStateException(EXPIRED);
            }catch(IllegalStateException ex){ra.addFlashAttribute("error",ex.getMessage());return "redirect:/mcm2003u";}
        }
    }
    @GetMapping("/detail")
    public String detail(@RequestParam(required=false) BigDecimal brandId,
                         @RequestParam(required=false) BigDecimal mitsumoriId,HttpSession session,RedirectAttributes ra) {
        if(!(session.getAttribute(FORM) instanceof Mcm2003uForm form))return "redirect:/mcm2004u";
        if(brandId!=null && form.getBrandRows().stream().noneMatch(b->same(b.getUmKihonBrandId(),brandId))
                || mitsumoriId!=null && form.getMitsumoriRows().stream().noneMatch(m->same(m.getUmMitsumoriId(),mitsumoriId))) {
            ra.addFlashAttribute("error",EXPIRED);return "redirect:/mcm2003u";
        }
        if(brandId==null && mitsumoriId==null)return "redirect:/mcm2003u";
        session.setAttribute("mcm2005u.parentReadOnly",readOnly(form,session));
        session.setAttribute("mcm2005u.parentId",form.getUmKihonMitsumoriId());
        BigDecimal selectedBrand=brandId;
        if(selectedBrand==null && session.getAttribute("mcm2003u.selectedBrandId") instanceof BigDecimal selected
                && form.getBrandRows().stream().anyMatch(b->same(b.getUmKihonBrandId(),selected)))selectedBrand=selected;
        return "redirect:/mcm2005u?"+(selectedBrand==null?"":"umKihonBrandId="+selectedBrand.toPlainString()+"&")
            +(mitsumoriId==null?"":"umMitsumoriId="+mitsumoriId.toPlainString());
    }
    @GetMapping("/back") public String back(HttpSession session) {
        session.removeAttribute(FORM);rotate(session,SCREEN);return "redirect:/mcm2004u";
    }
    @GetMapping("/lock-release") public String lockRelease(HttpSession session,RedirectAttributes ra) {
        session.removeAttribute(FORM);rotate(session,SCREEN);
        ra.addFlashAttribute("message","最新の情報を再読み込みしました。");return "redirect:/mcm2003u";
    }

    @PostMapping({"/attachment/add","/attachment/remove"})
    public String attachment(HttpServletRequest request,HttpSession session,RedirectAttributes ra,
            @RequestParam(required=false) org.springframework.web.multipart.MultipartFile file,
            @RequestParam(required=false) BigDecimal attachmentId) {
        synchronized(session) {
            try {
                check(session,SCREEN,request);
                if(!(session.getAttribute(FORM) instanceof Mcm2003uForm saved))throw new IllegalStateException(EXPIRED);
                if(readOnly(saved,session))throw new IllegalStateException(DENIED);
                var form=copy(saved);
                bind(request,form,"keiyakujikantai","mitsumoriLevel","mitsumorigiken","iraitantoNk","sofutantoNk","mitsumoriJouken","hosyuhoho","biko");
                session.setAttribute(FORM,form);
                if(request.getRequestURI().endsWith("/add"))attachments.add(form,file,getLoginUserId());
                else attachments.remove(form,attachmentId,getLoginUserId());
                form.setTenpuRows(service.load(form.getUmKihonMitsumoriId()).getTenpuRows());
                rotate(session,SCREEN);ra.addFlashAttribute("message","添付資料を更新しました。");
            } catch(IllegalStateException ex){ra.addFlashAttribute("error",ex.getMessage());}
            catch(Exception ex){logger.error("MCM2003U 添付更新失敗",ex);ra.addFlashAttribute("error","添付資料を更新できませんでした。");}
            return "redirect:/mcm2003u";
        }
    }
    @GetMapping("/attachment")
    public org.springframework.http.ResponseEntity<org.springframework.core.io.Resource> download(@RequestParam BigDecimal attachmentId,HttpSession session,RedirectAttributes ra,HttpServletRequest request) {
        try {
            if(!(session.getAttribute(FORM) instanceof Mcm2003uForm form))throw new IllegalStateException(EXPIRED);
            var row=attachments.row(form.getUmKihonMitsumoriId(),attachmentId);
            var resource=new org.springframework.core.io.FileSystemResource(attachments.file(row));
            return org.springframework.http.ResponseEntity.ok().contentType(org.springframework.http.MediaType.APPLICATION_OCTET_STREAM)
                .header("Content-Disposition",org.springframework.http.ContentDisposition.attachment().filename(row.getTenpufileNk(),java.nio.charset.StandardCharsets.UTF_8).build().toString())
                .body(resource);
        } catch(Exception ex){ra.addFlashAttribute("error","添付資料を取得できませんでした。保存先を確認してください。");return org.springframework.http.ResponseEntity.status(302).location(java.net.URI.create(request.getContextPath()+"/mcm2003u")).build();}
    }
    private boolean readOnly(Mcm2003uForm f,HttpSession s) {
        return !update(s,"MCM2003U",hasUpdateAuthority(),()->permissions.getAuthority(getLoginUserId(),"MCM2003U")) || service.isReadOnly(f.getSyouninJotai(),f.getSeniMotoKbn())
            || "3".equals(f.getJotai()) || "4".equals(f.getJotai());
    }
    @Override protected String getScreenTitle(){return "店舗見積内容基本設定";}
    @Override protected String getFunctionId(){return "MCM2003U";}
}
