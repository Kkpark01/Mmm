package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;
import jakarta.servlet.http.HttpSession;
import org.springframework.beans.BeanWrapperImpl;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.WebDataBinder;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.common.Mcm0013uConstants;
import com.daifuku.mcm.dto.Mcm0014uDeliveryDto;
import com.daifuku.mcm.dto.UserInfo;
import com.daifuku.mcm.form.Mcm0013uForm;
import com.daifuku.mcm.service.ComboBoxDataService;
import com.daifuku.mcm.service.Mcm0013uService;

/** 納入機器マスタ。画面操作はセッション内、DB更新は登録時の一括処理のみ。 */
@Controller
@RequestMapping("/mcm0013u")
public class Mcm0013uController extends BaseController {
    private static final String FORM="mcm0013uForm";
    private static final String SESSION_ERROR="セッションが切れました。画面を再読み込みしてください。";
    private final Mcm0013uService service;
    @Autowired private ComboBoxDataService comboBoxDataService;
    public Mcm0013uController(Mcm0013uService service){this.service=service;}
    @Override protected String getScreenTitle(){return Mcm0013uConstants.SCREEN_TITLE;}
    @Override protected String getFunctionId(){return Mcm0013uConstants.SCREEN_ID;}
    @InitBinder("input")
    public void bindInput(WebDataBinder binder) {
        binder.setAutoGrowCollectionLimit(100000);
        binder.setAllowedFields("plantId","selectedKikikoseiId","returnTo0012Url",
            "koseiRows[*].kikikoseiId","koseiRows[*].kikikoseiNk","koseiRows[*].setNm","koseiRows[*].tani","koseiRows[*].tehaiseiban","koseiRows[*].controllerFlg","koseiRows[*].biko",
            "kotaikanriRows[*].kotaikanriId","kotaikanriRows[*].brandkoseiId","kotaikanriRows[*].atsukaikikikoseiId","kotaikanriRows[*].hyojijun","kotaikanriRows[*].kotaiNk","kotaikanriRows[*].serialNo","kotaikanriRows[*].itijinonyuDt","kotaikanriRows[*].setchibasyo","kotaikanriRows[*].tekkyoDt","kotaikanriRows[*].keiyakukigenDt","kotaikanriRows[*].enchokeiyakukigenDt","kotaikanriRows[*].keiyakumanryoyoteiDt","kotaikanriRows[*].upskokanDt","kotaikanriRows[*].biko");
    }
    private Mcm0013uForm current(HttpSession s,Long plant) {
        Object value=s.getAttribute(FORM);
        return value instanceof Mcm0013uForm f&&Objects.equals(plant,f.getPlantId())?f:null;
    }
    private static String encode(String v){return URLEncoder.encode(v==null?"":v,StandardCharsets.UTF_8);}
    private String back(Mcm0013uForm f){return "redirect:/mcm0013u?plantId="+f.getPlantId();}
    @GetMapping
    public String index(@RequestParam Long plantId,@RequestParam(required=false) String returnTo0012Url,@RequestParam(required=false) Long selectedKikikoseiId,Model m,HttpSession s) {
        Mcm0013uForm f=current(s,plantId);
        if(f==null) {
            clear(s); f=service.loadForm(plantId); s.setAttribute(FORM,f);
            // #301: 登録直後のリダイレクトでは、セッション破棄によりloadForm()が
            // HYOJIJUN先頭行を選択してしまう。登録直前に選択していた行がまだ存在し
            // 有効な場合のみ、その選択状態を復元する（存在しない場合は従来通り先頭行）。
            if(selectedKikikoseiId!=null) {
                var row=f.findKoseiRow(selectedKikikoseiId);
                if(row!=null&&Mcm0013uService.active(row.getRowStatus())) f.setSelectedKikikoseiId(selectedKikikoseiId);
            }
        }
        if(returnTo0012Url!=null&&returnTo0012Url.startsWith("/mcm0012u")&&!returnTo0012Url.contains("\r")&&!returnTo0012Url.contains("\n")) f.setReturnTo0012Url(returnTo0012Url);
        return render(f,m,s);
    }
    private String render(Mcm0013uForm f,Model m,HttpSession s) {
        if(f.findKoseiRow(f.getSelectedKikikoseiId())==null||!Mcm0013uService.active(f.findKoseiRow(f.getSelectedKikikoseiId()).getRowStatus()))
            f.setSelectedKikikoseiId(f.getActiveKoseiRows().isEmpty()?null:f.getActiveKoseiRows().get(0).getKikikoseiId());
        List<Map<String,Object>> parents=new ArrayList<>();
        for(var r:f.getActiveKoseiRows()) {var values=new HashMap<String,Object>(f.getKoseiDisplay().getOrDefault(r.getKikikoseiId(),Map.of()));values.putAll(toMap(r));values.put("CONTROLLER_FLG","1".equals(r.getControllerFlg())?1:0);parents.add(values);}
        var visible=service.activeIndividuals(f,f.getSelectedKikikoseiId());
        Mcm0013uForm view=f.copy();view.setKotaikanriRows(new ArrayList<>(visible));
        List<Map<String,Object>> individualDisplay=new ArrayList<>();
        for(var k:visible)individualDisplay.add(f.getKotaiDisplay().getOrDefault(k.getKotaikanriId(),Map.of()));
        List<Map<String,Object>> details=new ArrayList<>();
        Map<String,String> names=new HashMap<>();
        for(var n:comboBoxDataService.getNounyukubunList())names.put(n.get("key"),n.get("value"));
        for(var d:service.activeDetails(f,f.getSelectedKikikoseiId())) {
            var data=toMap(d);data.put("KOTAIKANRI_FLG",d.getKotaikanriFlg());data.put("NOUNYU_KBN_NAME",names.getOrDefault(d.getNounyuKbn(),""));details.add(data);
        }
        m.addAttribute("form",view);m.addAttribute("koseiList",parents);m.addAttribute("meisaiList",details);m.addAttribute("kotaikanriList",individualDisplay);
        m.addAttribute("brandKoseiCombo",service.getBrandKoseiCombo(f.getPlantId()));m.addAttribute("atsukaikikiKoseiCombo",service.getAtsukaikikiKoseiCombo());m.addAttribute("taniCombo",service.getTaniCombo());
        // 全候補を各個体行で繰り返し展開しない。必要な機器の候補だけを描画する。
        Map<Long,List<Map<String,Object>>> partnersByEquipment=new HashMap<>();
        @SuppressWarnings("unchecked") var partners=(List<Map<String,Object>>)m.getAttribute("atsukaikikiKoseiCombo");
        for(var partner:partners) {
            Object equipment=partner.get("ATSUKAIKIKI_ID");
            if(equipment instanceof Number n)partnersByEquipment.computeIfAbsent(n.longValue(),key->new ArrayList<>()).add(partner);
        }
        m.addAttribute("partnersByEquipment",partnersByEquipment);
        m.addAttribute("returnTo0012Url",f.getReturnTo0012Url());m.addAttribute("hasChanges",service.hasChanges(f));
        setCommonAttributes(m,s);m.addAttribute("screenId",Mcm0013uConstants.SCREEN_ID);
        return "mcm0013u/index";
    }
    private static Map<String,Object> toMap(Object bean) {
        var w=new BeanWrapperImpl(bean);Map<String,Object> map=new HashMap<>();
        for(var p:w.getPropertyDescriptors())if(!"class".equals(p.getName()))map.put(p.getName().replaceAll("([a-z0-9])([A-Z])","$1_$2").toUpperCase(Locale.ROOT),w.getPropertyValue(p.getName()));
        return map;
    }
    private static boolean equalValue(Object a,Object b){return Objects.equals(a,b)||(a==null&&"".equals(b))||(b==null&&"".equals(a));}
    private boolean copyFields(Object from,Object to,String... fields) {
        var a=new BeanWrapperImpl(from);var b=new BeanWrapperImpl(to);boolean changed=false;
        for(String key:fields)if(!equalValue(a.getPropertyValue(key),b.getPropertyValue(key))){b.setPropertyValue(key,a.getPropertyValue(key));changed=true;}
        return changed;
    }
    private void merge(Mcm0013uForm f,Mcm0013uForm input) {
        for(var r:input.getKoseiRows()) {
            var target=f.findKoseiRow(r.getKikikoseiId());if(target==null||!Mcm0013uService.active(target.getRowStatus()))continue;
            Integer oldSet=target.getSetNm();
            if(copyFields(r,target,"kikikoseiNk","setNm","tani","tehaiseiban","controllerFlg","biko"))Mcm0013uService.mark(target);
            if(!Objects.equals(oldSet,target.getSetNm()))service.syncWorkingIndividuals(f,target.getKikikoseiId());
        }
        for(var r:input.getKotaikanriRows()) {
            var target=f.findKotaiRow(r.getKotaikanriId());if(target==null||!Mcm0013uService.active(target.getRowStatus()))continue;
            if(copyFields(r,target,"brandkoseiId","atsukaikikikoseiId","hyojijun","kotaiNk","serialNo","itijinonyuDt","setchibasyo","tekkyoDt","keiyakukigenDt","enchokeiyakukigenDt","keiyakumanryoyoteiDt","upskokanDt","biko")&&!"new".equals(target.getRowStatus())) target.setRowStatus("modified");
        }
    }
    private List<String> bindingErrors(BindingResult b) {
        if(!b.hasErrors())return List.of();
        String field=b.getFieldErrors().isEmpty()?"":b.getFieldErrors().get(0).getField();
        String name=field.endsWith("setNm")?"セット数":field.endsWith("hyojijun")?"No":field.endsWith("brandkoseiId")?"ブランド":"取引先";
        return List.of(name+"は数値で入力してください。");
    }
    private Map<String,Object> response(List<String> errors) {
        String title=!errors.isEmpty()&&errors.get(0).contains("既に使用されている為")?"削除エラー":"入力エラー";
        return Map.of("errors",errors,"title",title);
    }
    // 確認ダイアログより前の検証。ここでは入力保存のみ、DB更新・削除マークは行わない。
    @PostMapping("/validate") @ResponseBody
    public Map<String,Object> validate(@ModelAttribute("input") Mcm0013uForm input,BindingResult b,@RequestParam String action,@RequestParam(required=false) Long targetId,HttpSession s) {
        synchronized(s) {
            var f=current(s,input.getPlantId());if(f==null)return response(List.of(SESSION_ERROR));
            var errors=bindingErrors(b);if(!errors.isEmpty())return response(errors);merge(f,input);
            return response(switch(action){case "update"->service.validateWorking(f);case "deleteKosei"->service.checkDeleteKosei(f,targetId);case "deleteKotai"->service.checkDeleteKotai(f,targetId);default->List.of(Mcm0013uConstants.MSG_NO_SELECTION);});
        }
    }
    @PostMapping("/update")
    public String update(@ModelAttribute("input") Mcm0013uForm input,BindingResult b,Model m,HttpSession s,RedirectAttributes ra) {
        synchronized(s) {
            var f=current(s,input.getPlantId());if(f==null){ra.addFlashAttribute("errors",List.of(SESSION_ERROR));return "redirect:/mcm0013u?plantId="+input.getPlantId();}
            var errors=bindingErrors(b);
            Long selected=f.getSelectedKikikoseiId();
            if(errors.isEmpty()) {
                merge(f,input);
                // 採番や保存の途中で例外になっても、元の未登録セッションを変更しない。
                try {
                    var saving=f.copy();
                    var selectedRow=saving.findKoseiRow(selected);
                    errors=service.saveWorking(saving,login(s));
                    // #301: 新規構成も、保存用コピーで採番された正式IDを使って復元する。
                    if(errors.isEmpty()&&selectedRow!=null) selected=selectedRow.getKikikoseiId();
                }
                catch(IllegalStateException e){errors=List.of(e.getMessage());}
                catch(org.springframework.dao.DataAccessException e){errors=List.of("他のユーザがデータを変更した可能性があります。処理をやり直してください。");}
            }
            if(!errors.isEmpty()){m.addAttribute("errors",errors);return render(f,m,s);}
            // #301: 登録済み・新規ともに正式IDを引き継ぎ、同じ構成の個体を再表示する。
            String dest=back(f)+(selected!=null&&!Mcm0013uForm.isTempId(selected)?"&selectedKikikoseiId="+selected:"")
                +(f.getReturnTo0012Url()==null?"":"&returnTo0012Url="+encode(f.getReturnTo0012Url()));
            clear(s);
            ra.addFlashAttribute("message",Mcm0013uConstants.MSG_UPDATE_SUCCESS);return dest;
        }
    }
    @PostMapping({"/addKoseiRow","/deleteKoseiRow","/deleteKotaiRow","/selectKoseiRow"})
    public String edit(@ModelAttribute("input") Mcm0013uForm input,BindingResult b,@RequestParam(required=false) Long kikikoseiId,@RequestParam(required=false) Long kotaikanriId,jakarta.servlet.http.HttpServletRequest request,HttpSession s,RedirectAttributes ra) {
        synchronized(s) {
            var f=current(s,input.getPlantId());if(f==null){ra.addFlashAttribute("errors",List.of(SESSION_ERROR));return "redirect:/mcm0013u?plantId="+input.getPlantId();}
            var errors=bindingErrors(b);if(!errors.isEmpty()){ra.addFlashAttribute("errors",errors);return back(f);}merge(f,input);
            String action=request.getRequestURI();
            if(action.endsWith("/addKoseiRow")) {
                for(var r:f.getActiveKoseiRows()){errors=service.checkRequired(r);if(!errors.isEmpty())break;}
                if(errors.isEmpty()) {
                    var r=new Mcm0013uForm.KoseiRowForm();r.setKikikoseiId(f.nextTempId());r.setPlantId(f.getPlantId());r.setSetNm(1);r.setKikikoseiNk("");r.setTani("");r.setControllerFlg("0");r.setRowStatus("new");
                    r.setHyojijun(f.getActiveKoseiRows().stream().mapToInt(x->x.getHyojijun()==null?0:x.getHyojijun()).max().orElse(0)+1);f.getKoseiRows().add(r);f.setSelectedKikikoseiId(r.getKikikoseiId());
                }
            } else if(action.endsWith("/deleteKoseiRow")) {
                errors=service.checkDeleteKosei(f,kikikoseiId);
                if(errors.isEmpty()){var r=f.findKoseiRow(kikikoseiId);if("new".equals(r.getRowStatus())){f.getKoseiRows().remove(r);f.getDetails().remove(kikikoseiId);}else r.setRowStatus("deleted");}
            } else if(action.endsWith("/deleteKotaiRow")) {
                errors=service.checkDeleteKotai(f,kotaikanriId);
                if(errors.isEmpty()){var r=f.findKotaiRow(kotaikanriId);if("new".equals(r.getRowStatus()))f.getKotaikanriRows().remove(r);else r.setRowStatus("deleted");}
            } else if(f.findKoseiRow(kikikoseiId)!=null) f.setSelectedKikikoseiId(kikikoseiId);
            if(!errors.isEmpty())ra.addFlashAttribute("errors",errors);
            return back(f);
        }
    }
    @PostMapping("/syncKotaiForSetNum") @ResponseBody
    public Map<String,Object> syncKotaiForSetNum(@ModelAttribute("input") Mcm0013uForm input,BindingResult b,@RequestParam Long kikikoseiId,@RequestParam Integer setNm,HttpSession s) {
        synchronized(s) {
            var f=current(s,input.getPlantId());if(f==null)return response(List.of(SESSION_ERROR));
            var errors=bindingErrors(b);if(!errors.isEmpty())return response(errors);
            if(setNm<1)return response(List.of(Mcm0013uConstants.MSG_SETNUM_INVALID));
            if(setNm>9999)return response(List.of("セット数は4桁以下で入力してください。"));
            var r=f.findKoseiRow(kikikoseiId);if(r==null||!Mcm0013uService.active(r.getRowStatus()))return response(List.of(Mcm0013uConstants.MSG_NO_SELECTION));
            merge(f,input);if(!Objects.equals(r.getSetNm(),setNm)){r.setSetNm(setNm);Mcm0013uService.mark(r);}
            service.syncWorkingIndividuals(f,kikikoseiId);return response(List.of());
        }
    }
    @PostMapping("/goToMcm0014u")
    public String goToMcm0014u(@ModelAttribute("input") Mcm0013uForm input,BindingResult b,@RequestParam Long kikikoseiId,HttpSession s,RedirectAttributes ra) {
        synchronized(s) {
            var f=current(s,input.getPlantId());if(f==null){ra.addFlashAttribute("errors",List.of(SESSION_ERROR));return "redirect:/mcm0013u?plantId="+input.getPlantId();}
            var errors=bindingErrors(b);if(errors.isEmpty())merge(f,input);
            var row=f.findKoseiRow(kikikoseiId);if(errors.isEmpty())errors=service.checkRequired(row);
            if(!errors.isEmpty()){ra.addFlashAttribute("errors",errors);return back(f);}
            f.setSelectedKikikoseiId(kikikoseiId);
            var d=new Mcm0014uDeliveryDto();d.setPlantId(f.getPlantId());d.setKikikoseiId(kikikoseiId);d.setKikikoseiNk(row.getKikikoseiNk());d.setSetNum(row.getSetNm());
            d.setNonyusakiCd(f.getNonyusakiCd());d.setNonyusakiNk(f.getNonyusakiNk());d.setSupportId(f.getSupportId());d.setPlantNk(f.getPlantNk());
            d.setMeisaiRows(new ArrayList<>(f.copy().getDetails().getOrDefault(kikikoseiId,List.of())));
            s.setAttribute("mcm0014u_delivery",d);s.setAttribute("mcm0014u_plantId",f.getPlantId());
            return "redirect:/mcm0014u?kikikoseiId="+kikikoseiId+"&plantId="+f.getPlantId()+"&kikikoseiNk="+encode(row.getKikikoseiNk())+"&setNm="+row.getSetNm()+"&returnTo0013Url="+encode("/mcm0013u?plantId="+f.getPlantId())+(f.getReturnTo0012Url()==null?"":"&returnTo0012Url="+encode(f.getReturnTo0012Url()));
        }
    }
    @GetMapping("/returnFromMcm0014u")
    public String returnFromMcm0014u(HttpSession s) {
        synchronized(s) {
            Object value=s.getAttribute(FORM);if(!(value instanceof Mcm0013uForm f))return "redirect:/mcm0010u";
            Object d=s.getAttribute("mcm0014u_return_data");if(d instanceof Mcm0014uDeliveryDto delivery)service.acceptDelivery(f,delivery);
            s.removeAttribute("mcm0014u_return_data");s.removeAttribute("mcm0014u_delivery");return back(f);
        }
    }
    // 従来の参照APIも、表示中の未登録データを含む同じ状態を返す。
    @GetMapping("/api/meisai") @ResponseBody
    public List<Map<String,Object>> getMeisaiApi(@RequestParam Long kikikoseiId,HttpSession s) {
        if(!(s.getAttribute(FORM) instanceof Mcm0013uForm f)||f.findKoseiRow(kikikoseiId)==null)return List.of();
        return service.activeDetails(f,kikikoseiId).stream().map(d->{var m=toMap(d);m.put("KOTAIKANRI_FLG",d.getKotaikanriFlg());return m;}).toList();
    }
    @GetMapping("/api/kotaikanri") @ResponseBody
    public List<Map<String,Object>> getKotaikanriApi(@RequestParam Long kikikoseiId,HttpSession s) {
        if(!(s.getAttribute(FORM) instanceof Mcm0013uForm f)||f.findKoseiRow(kikikoseiId)==null)return List.of();
        return service.activeIndividuals(f,kikikoseiId).stream().map(k->{var m=new HashMap<String,Object>(f.getKotaiDisplay().getOrDefault(k.getKotaikanriId(),Map.of()));m.putAll(toMap(k));return (Map<String,Object>)m;}).toList();
    }
    @PostMapping("/discard")
    public String discard(@RequestParam String destination,HttpSession s) {
        var f=(Mcm0013uForm)s.getAttribute(FORM);
        String dest="back".equals(destination)&&f!=null&&f.getReturnTo0012Url()!=null?f.getReturnTo0012Url():"back".equals(destination)?"/mcm0010u":"/menu";
        clear(s);return "redirect:"+dest;
    }
    private void clear(HttpSession s) {
        for(String key:List.of(FORM,"mcm0013u_form_backup","mcm0014u_pending_map","mcm0014u_return_data","mcm0014u_delivery","mcm0014u_plantId"))s.removeAttribute(key);
    }
    private String login(HttpSession s){return s.getAttribute("userInfo") instanceof UserInfo u?u.getLoginId():"SYSTEM";}
}
