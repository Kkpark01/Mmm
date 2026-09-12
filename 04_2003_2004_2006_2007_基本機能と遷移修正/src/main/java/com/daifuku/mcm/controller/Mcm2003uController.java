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
    @Autowired private com.daifuku.mcm.service.McmEstimateMailService mail;
    @Autowired private com.daifuku.mcm.service.Mcm2003uDraftService drafts;
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
                boolean pending=session.getAttribute("mcm2003u.draft") instanceof com.daifuku.mcm.form.McmEstimateDraft;
                if(pending)service.clearUnlock(session);else service.isUnlocked(id,session);
                boolean changed=!pending && (form==null || !same(id,form.getUmKihonMitsumoriId())
                    || entry instanceof Integer mode && form.getSeniMotoKbn()!=mode);
                if(changed) {
                    if(id==null) { ra.addFlashAttribute("error","見積情報を選択してから開いてください。");return "redirect:/mcm2004u"; }
                    form=service.load(id);
                    form.setSeniMotoKbn(entry instanceof Integer mode ? mode : 3);
                    rotate(session,SCREEN);
                }
                if(Boolean.TRUE.equals(session.getAttribute("mcm2003u.refreshDetails"))) {
                    var latest=service.load(form.getUmKihonMitsumoriId());
                    // 詳細設定から戻っても、この画面で編集したブランド別保守方法を保持する。
                    for(var old:form.getBrandRows())if(old.isSoftHosyuhohoEdited())latest.getBrandRows().stream().filter(b->same(b.getUmKihonBrandId(),old.getUmKihonBrandId())).findFirst().ifPresent(b->{b.setSoftHosyuhoho(old.getSoftHosyuhoho());b.setOriginalSoftHosyuhoho(old.getOriginalSoftHosyuhoho());b.setSoftHosyuhohoEdited(true);});
                    for(var old:form.getBrandRows())latest.getBrandRows().stream().filter(b->same(b.getUmKihonBrandId(),old.getUmKihonBrandId())).findFirst().ifPresent(b->com.daifuku.mcm.common.Mcm2003uEdits.overlay(old,b));
                    for(var old:form.getKoseiRows())latest.getKoseiRows().stream().filter(k->same(k.getUmKikikoseiId(),old.getUmKikikoseiId())).findFirst().ifPresent(k->com.daifuku.mcm.common.Mcm2003uEdits.overlay(old,k));
                    form.setBrandRows(latest.getBrandRows());form.setKoseiRows(latest.getKoseiRows());form.setMeisaiRows(latest.getMeisaiRows());form.setMitsumoriRows(latest.getMitsumoriRows());
                    session.removeAttribute("mcm2003u.refreshDetails");
                }
                session.setAttribute(FORM,form);
                session.setAttribute("mcm2003u.seniMotoKbn",form.getSeniMotoKbn());
                BigDecimal requestedBrand=brandId != null ? brandId : (BigDecimal)session.getAttribute("mcm2003u.selectedBrandId");
                var selected=form.getBrandRows().stream().filter(b->same(b.getUmKihonBrandId(),requestedBrand)).findFirst()
                    .orElse(form.getBrandRows().isEmpty()?null:form.getBrandRows().get(0));
                model.addAttribute("pendingEstimate",pending);
                if(drafts!=null)model.addAttribute("tenpoChoices",drafts.tenpoChoices());
                setCommonAttributes(model,session);
                model.addAttribute("backLabel", "/mcm3007u".equals(session.getAttribute("mcm2003u.returnTo")) ? "契約・解約画面へ" : "店舗見積・契約検索画面へ");
                model.addAttribute("form",form);model.addAttribute("readOnly",readOnly(form,session));
                model.addAttribute("adminUnlocked",!pending&&service.isUnlocked(form.getUmKihonMitsumoriId(),session));
                model.addAttribute("canUnlock",!pending&&service.canReleaseLock(session)&&update(session,"MCM2003U",hasUpdateAuthority(),()->permissions.getAuthority(getLoginUserId(),"MCM2003U")));
                model.addAttribute("selectedBrand",selected);
                if(selected!=null)session.setAttribute("mcm2003u.selectedBrandId",selected.getUmKihonBrandId());
                model.addAttribute("workflowToken",token(session,SCREEN));
                model.addAttribute("reportAvailable",session.getAttribute("mcm2001p.report") instanceof Mcm2001pController.Report report&&same(report.estimateId(),form.getUmKihonMitsumoriId()));
                if(!pending){try{String status=mail.status(form.getUmKihonMitsumoriId(),getLoginUserId());model.addAttribute("mailStatus",status);if(!status.equals("NONE"))model.addAttribute("mailStatusMessage",com.daifuku.mcm.service.McmEstimateMailService.message(status));}catch(org.springframework.dao.DataAccessException ex){model.addAttribute("mailSetupRequired",true);}}
                model.addAttribute("visibleKosei",form.getKoseiRows().stream().filter(k->selected!=null && same(k.getUmKihonBrandId(),selected.getUmKihonBrandId())).toList());
                final Mcm2003uForm displayed = form;
                model.addAttribute("visibleMeisai",form.getMeisaiRows().stream().filter(m->selected!=null && displayed.getKoseiRows().stream()
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
                bindHeader(request,form,session);
                session.setAttribute(FORM,form);
                String op=request.getRequestURI().substring(request.getRequestURI().lastIndexOf('/')+1);
                if(session.getAttribute("mcm2003u.draft") instanceof com.daifuku.mcm.form.McmEstimateDraft draft) {
                    if(!op.equals("save"))throw new IllegalStateException("見積登録を行ってから操作してください。");
                    var selectedPeriods=selections(request,"periodSelection",form.getMitsumoriRows().size());
                    for(int i=0;i<form.getMitsumoriRows().size();i++)form.getMitsumoriRows().get(i).setCheckFlg(selectedPeriods.contains(i));
                    var newId=drafts.save(draft,form,selectedPeriods,getLoginUserId());
                    session.setAttribute(ID,newId);session.setAttribute("mcm2003u.seniMotoKbn",3);session.removeAttribute("mcm2003u.draft");
                }
                else if(op.equals("publish")){
                    byte[] bytes=service.publish(form,getLoginUserId());
                    session.setAttribute("mcm2001p.report",new Mcm2001pController.Report(form.getUmKihonMitsumoriId(),getLoginUserId(),bytes));
                    ra.addFlashAttribute("downloadReport",true);
                }
                else if(op.equals("sinsei"))service.saveAndApply(form,getLoginUserId());
                else service.save(form,getLoginUserId());
                session.removeAttribute(FORM);rotate(session,SCREEN);
                session.setAttribute("mcm2004u.refresh",true);
                String notice=op.equals("publish")?"見積を発行しました。Excel帳票をダウンロードできます。":"登録を完了しました。";
                if(op.equals("sinsei")){
                    try{notice=com.daifuku.mcm.service.McmEstimateMailService.message(mail.dispatchLatest(form.getUmKihonMitsumoriId(),getLoginUserId(),false));}
                    catch(Exception ex){logger.error("申請コミット後のメール送信処理に失敗",ex);notice="申請は完了しました。メール送信状況を確認してください。";}
                }
                ra.addFlashAttribute("message",notice);
            } catch(IllegalStateException e) {ra.addFlashAttribute("error",e.getMessage());}
            catch(Exception e) {logger.error("MCM2003U 更新失敗",e);ra.addFlashAttribute("error","登録できませんでした。入力内容を保持しています。");}
            return "redirect:/mcm2003u";
        }
    }
    @PostMapping("/mail/retry")
    public String retryMail(HttpServletRequest request,HttpSession session,RedirectAttributes ra){synchronized(session){try{
        check(session,SCREEN,request);
        if(!(session.getAttribute(FORM) instanceof Mcm2003uForm f)||!update(session,"MCM2003U",hasUpdateAuthority(),()->permissions.getAuthority(getLoginUserId(),"MCM2003U")))throw new IllegalStateException(DENIED);
        var current=service.load(f.getUmKihonMitsumoriId());if(!"1".equals(current.getSyouninJotai()))throw new IllegalStateException("審査中の申請メールだけ再送できます。");
        ra.addFlashAttribute("message",com.daifuku.mcm.service.McmEstimateMailService.message(mail.dispatchLatest(f.getUmKihonMitsumoriId(),getLoginUserId(),true)));rotate(session,SCREEN);
    }catch(IllegalStateException ex){ra.addFlashAttribute("error",ex.getMessage());}catch(Exception ex){logger.error("申請メール再送失敗",ex);ra.addFlashAttribute("error","メール送信状況を確認してください。");}return "redirect:/mcm2003u";}}
    @PostMapping("/navigate")
    public String navigate(@RequestParam String destination,@RequestParam(required=false) BigDecimal brandId,
            @RequestParam(required=false) BigDecimal mitsumoriId,HttpServletRequest request,HttpSession session,RedirectAttributes ra) {
        synchronized(session) {
            try {
                check(session,SCREEN,request);
                if(!(session.getAttribute(FORM) instanceof Mcm2003uForm saved))throw new IllegalStateException(EXPIRED);
                if(!readOnly(saved,session)) {
                    var input=copy(saved);
                    bindHeader(request,input,session);
                    if(session.getAttribute("mcm2003u.draft")!=null && request.getParameter("periodSelectionPresent")!=null){var periods=selections(request,"periodSelection",input.getMitsumoriRows().size());for(int i=0;i<input.getMitsumoriRows().size();i++)input.getMitsumoriRows().get(i).setCheckFlg(periods.contains(i));}
                    session.setAttribute(FORM,input);
                }
                if(destination.equals("selection") && session.getAttribute("mcm2003u.draft")!=null)return "redirect:/mcm2002u";
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
        if(session.getAttribute("mcm2003u.draft")!=null){ra.addFlashAttribute("error","見積登録を行ってから詳細設定を開いてください。");return "redirect:/mcm2003u";}
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
        String destination="/mcm3007u".equals(session.getAttribute("mcm2003u.returnTo"))?"/mcm3007u":"/mcm2004u";
        session.removeAttribute("mcm2003u.returnTo");clearScreen(session);return "redirect:"+destination;
    }
    @GetMapping({"/reload","/lock-release"}) public String reload(HttpSession session,RedirectAttributes ra) {
        synchronized(session){
        if(session.getAttribute("mcm2003u.draft")!=null)return "redirect:/mcm2003u";
        service.clearUnlock(session);session.removeAttribute(FORM);rotate(session,SCREEN);
        ra.addFlashAttribute("message","最新の情報を再読み込みしました。");return "redirect:/mcm2003u";
        }
    }

    @GetMapping("/close") public String close(HttpSession session){clearScreen(session);return "redirect:/menu";}
    private void clearScreen(HttpSession session){synchronized(session){
        session.removeAttribute("mcm2003u.returnTo");
        service.clearUnlock(session);session.removeAttribute(FORM);session.removeAttribute("mcm2003u.draft");
        session.removeAttribute("mcm2003u.refreshDetails");session.removeAttribute("MCM2005U_FORM");session.removeAttribute("MCM2005U_PENDING");
        session.removeAttribute("mcm2005u.parentId");session.removeAttribute("mcm2005u.parentReadOnly");
        rotate(session,SCREEN);rotate(session,"mcm2005u");
    }}
    /** VBのLockReleaseButton_Click同様、最新情報を読み直して画面の編集制限だけ解除する。 */
    @PostMapping("/lock-release") public String lockRelease(HttpServletRequest request,HttpSession session,RedirectAttributes ra){synchronized(session){
        try{
            check(session,SCREEN,request);
            if(!(session.getAttribute(FORM) instanceof Mcm2003uForm saved)
                ||!same(saved.getUmKihonMitsumoriId(),(BigDecimal)session.getAttribute(ID)))throw new IllegalStateException(EXPIRED);
            if(!update(session,"MCM2003U",hasUpdateAuthority(),()->permissions.getAuthority(getLoginUserId(),"MCM2003U")))throw new IllegalStateException(DENIED);
            var latest=service.releaseLock(saved.getUmKihonMitsumoriId(),saved.getSeniMotoKbn(),session);
            session.setAttribute(FORM,latest);session.removeAttribute("mcm2003u.refreshDetails");session.removeAttribute("MCM2005U_FORM");session.removeAttribute("MCM2005U_PENDING");
            rotate(session,SCREEN);rotate(session,"mcm2005u");
            logger.info("MCM2003U 編集制限解除 user={} estimate={}",getLoginUserId(),latest.getUmKihonMitsumoriId());
            ra.addFlashAttribute("message","最新情報を読み込み、管理者用の編集制限を解除しました。");
        }catch(IllegalStateException ex){ra.addFlashAttribute("error",ex.getMessage());}
        catch(Exception ex){service.clearUnlock(session);logger.error("MCM2003U 編集制限解除失敗",ex);ra.addFlashAttribute("error","編集制限を解除できませんでした。再検索してください。");}
        return "redirect:/mcm2003u";
    }}

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
                bindHeader(request,form,session);
                session.setAttribute(FORM,form);
                if(session.getAttribute("mcm2003u.draft")!=null)throw new IllegalStateException("見積登録を行ってから添付資料を追加してください。");
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
    private void bindHeader(HttpServletRequest request,Mcm2003uForm input,HttpSession session) {
        bind(request,input,"keiyakujikantai","mitsumoriLevel","mitsumorigiken","iraitantoNk","sofutantoNk","mitsumoriJouken","hosyuhoho","biko");
        if(session.getAttribute("mcm2003u.draft")!=null||service.isUnlocked(input.getUmKihonMitsumoriId(),session))bind(request,input,"umMitsumoriNo","iraitenpoId","sofutenpoId");
        com.daifuku.mcm.common.Mcm2003uEdits.bind(request,input);
        String method=request.getParameter("brandSoftHosyuhoho");
        if(method!=null) {
            BigDecimal edited;
            try { edited=new BigDecimal(request.getParameter("editedBrandId")); } catch(Exception ex) { throw new IllegalStateException(EXPIRED); }
            var brand=input.getBrandRows().stream().filter(b->same(b.getUmKihonBrandId(),edited)).findFirst().orElseThrow(()->new IllegalStateException(EXPIRED));
            if(method.length()>4000)throw new IllegalStateException("保守方法は4000文字以内で入力してください。");
            if(!java.util.Objects.equals(method,brand.getSoftHosyuhoho()==null?"":brand.getSoftHosyuhoho())){if(!brand.isSoftHosyuhohoEdited())brand.setOriginalSoftHosyuhoho(brand.getSoftHosyuhoho());brand.setSoftHosyuhoho(method);brand.setSoftHosyuhohoEdited(true);}
        }
    }
    private boolean readOnly(Mcm2003uForm f,HttpSession s) {
        return !update(s,"MCM2003U",hasUpdateAuthority(),()->permissions.getAuthority(getLoginUserId(),"MCM2003U")) || (!service.isUnlocked(f.getUmKihonMitsumoriId(),s) && (service.isReadOnly(f.getSyouninJotai(),f.getSeniMotoKbn())
            || "3".equals(f.getJotai()) || "4".equals(f.getJotai())));
    }
    @Override protected String getScreenTitle(){return "店舗見積内容基本設定";}
    @Override protected String getFunctionId(){return "MCM2003U";}
}
