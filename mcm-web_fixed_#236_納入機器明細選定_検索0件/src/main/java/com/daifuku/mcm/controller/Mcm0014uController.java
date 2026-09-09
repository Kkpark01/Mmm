package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.*;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpSession;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.util.MultiValueMap;
import org.springframework.web.bind.annotation.*;
import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.dto.Mcm0014uDeliveryDto;
import com.daifuku.mcm.form.Mcm0014uForm;
import com.daifuku.mcm.form.Mcm0014uForm.MeisaiRowForm;
import com.daifuku.mcm.service.ComboBoxDataService;
import com.daifuku.mcm.service.Mcm0014uService;

/** 14U内の編集を保持し、設定に成功した時だけ13Uへ渡す。DB更新は13U登録時のみ。 */
@Controller @RequestMapping("/mcm0014u")
public class Mcm0014uController extends BaseController {
    private static final String WORK="mcm0014u_work";
    private static final String SESSION_ERROR="画面の情報が更新されたか、有効期限が切れています。13Uから明細画面を開き直してください。";
    private static final String FAILURE="処理を完了できませんでした。入力内容を保持しています。もう一度操作してください。";
    private static final String REF_FAILURE="見積・契約情報の使用状況を確認できませんでした。削除せず、入力内容を保持しています。もう一度操作してください。";
    private final Mcm0014uService service;
    @Autowired private ComboBoxDataService comboBoxDataService;
    public Mcm0014uController(Mcm0014uService service){this.service=service;}
    @Override protected String getScreenTitle(){return "納入機器明細選定";}
    @Override protected String getFunctionId(){return "MCM0014U";}
    private static class Work implements java.io.Serializable {
        Mcm0014uDeliveryDto source;
        Mcm0014uForm form=new Mcm0014uForm();
        List<Map<String,Object>> results=new ArrayList<>(),partners=new ArrayList<>();
        boolean edited;
    }
    @GetMapping
    public String index(Model model,HttpSession s) {
        synchronized(s) {
            Object source=s.getAttribute("mcm0014u_delivery");
            if(!(source instanceof Mcm0014uDeliveryDto delivery))return "redirect:/mcm0013u/returnFromMcm0014u";
            Work w=s.getAttribute(WORK) instanceof Work old && old.source==delivery ? old : null;
            if(w==null) {
                w=new Work();w.source=delivery;var f=w.form;
                f.setPlantId(delivery.getPlantId());f.setKikikoseiId(delivery.getKikikoseiId());f.setKikikoseiNk(delivery.getKikikoseiNk());
                f.setSetNm(delivery.getSetNum());f.setSetNmText(str(delivery.getSetNum()));rotate(w);
                if(delivery.getMeisaiRows()!=null)for(var d:delivery.getMeisaiRows())f.getMeisaiRows().add(fromDto(d));
                s.setAttribute(WORK,w);
                for(String key:List.of("mcm0014u_return_data","mcm0014u_session_meisaiRows","mcm0014u_session_searchResults","mcm0014u_session_partners"))s.removeAttribute(key);
            }
            model.addAttribute("form",w.form);model.addAttribute("atsukaikikiList",w.results);model.addAttribute("atsukaikikiPartnerList",w.partners);
            model.addAttribute("hasChanges",w.edited);model.addAttribute("nounyukubunList",comboBoxDataService.getNounyukubunList());
            model.addAttribute("activeMeisaiCount",w.form.getMeisaiRows().stream().filter(Mcm0014uController::active).count());
            model.addAttribute("seizomakerCombo",service.getSeizomakerCombo());model.addAttribute("kikibunruiCombo",service.getKikibunruiCombo());model.addAttribute("torihikisakiCombo",service.getTorihikisakiCombo());
            return "mcm0014u/index";
        }
    }
    @PostMapping({"/search","/selectAtsukaikiki","/deleteMeisaiRow","/validateDelete","/settei","/discard"}) @ResponseBody
    public Map<String,Object> action(@RequestParam MultiValueMap<String,String> input,HttpSession s,HttpServletRequest request) {
        synchronized(s) {
            Work w=s.getAttribute(WORK) instanceof Work value?value:null;
            if(w==null||w.source!=s.getAttribute("mcm0014u_delivery")||!Objects.equals(w.form.getEditToken(),input.getFirst("editToken")))return answer(null,List.of(SESSION_ERROR));
            String op=request.getRequestURI().substring(request.getRequestURI().lastIndexOf('/')+1);
            if("discard".equals(op)) {
                s.removeAttribute(WORK);s.removeAttribute("mcm0014u_return_data");
                var a=answer(w,List.of());a.put("redirect",request.getContextPath()+"/mcm0013u/returnFromMcm0014u");return a;
            }
            List<String> errors=merge(w,input);
            if(!errors.isEmpty())return answer(w,errors);
            rotate(w); // 古い画面・二重送信が、別の行を編集しないようにする。
            try {
                switch(op) {
                    case "search": errors=search(w);break;
                    case "selectAtsukaikiki": errors=select(w,input.getFirst("atsukaikikiId"));break;
                    case "validateDelete": return answer(w,delete(w,input.getFirst("rowIndex"),false));
                    case "deleteMeisaiRow": errors=delete(w,input.getFirst("rowIndex"),true);break;
                    case "settei":
                        errors=validate(w);
                        if(errors.isEmpty()) {
                            s.setAttribute("mcm0014u_return_data",toDelivery(w));s.removeAttribute(WORK);
                            var a=answer(w,List.of());a.put("redirect",request.getContextPath()+"/mcm0013u/returnFromMcm0014u");return a;
                        }
                        break;
                    default: errors=List.of(SESSION_ERROR);
                }
            } catch(org.springframework.dao.DataAccessException ex) {
                org.slf4j.LoggerFactory.getLogger(getClass()).error("MCM0014U {} failed",op,ex);
                errors=List.of(op.contains("Delete")||op.contains("delete")||"settei".equals(op)?REF_FAILURE:FAILURE);
            }
            var a=answer(w,errors);
            if(!errors.isEmpty()) {if("search".equals(op))a.put("clearSearch",true);return a;}
            a.put("redirect",request.getContextPath()+"/mcm0014u");return a;
        }
    }
    /** 編集可能な値だけ取り込む。ID・分類・行状態は画面から受け取らない。 */
    private List<String> merge(Work w,MultiValueMap<String,String> p) {
        var f=w.form;var keys=p.getOrDefault("editRowKey",List.of());var expected=new ArrayList<String>();
        for(int i=0;i<f.getMeisaiRows().size();i++)if(active(f.getMeisaiRows().get(i)))expected.add(Integer.toString(i));
        if(!keys.equals(expected)||p.getOrDefault("editSuryoNm",List.of()).size()!=keys.size()
            ||p.getOrDefault("editHyojijun",List.of()).size()!=keys.size()
            ||p.getOrDefault("editNounyuKbn",List.of()).size()!=keys.size()||p.getOrDefault("editBiko",List.of()).size()!=keys.size()||!p.containsKey("setNm"))return List.of(SESSION_ERROR);
        w.edited|=!Objects.equals(f.getSetNmText(),p.getFirst("setNm"));f.setSetNmText(p.getFirst("setNm"));
        for(int i=0;i<keys.size();i++) {
            var r=f.getMeisaiRows().get(Integer.parseInt(keys.get(i)));String qty=p.get("editSuryoNm").get(i),kbn=p.get("editNounyuKbn").get(i),biko=p.get("editBiko").get(i),order=p.get("editHyojijun").get(i);
            if(!Objects.equals(str(r.getBiko()),biko)||!Objects.equals(str(r.getNounyuKbn()),kbn)||!Objects.equals(r.getSuryoNmText(),qty)||!Objects.equals(r.getHyojijunText(),order)) {
                w.edited=true;if(!"new".equals(r.getRowStatus()))r.setRowStatus("modified");
            }
            r.setSuryoNmText(qty);r.setBiko(biko);r.setNounyuKbn(kbn);r.setHyojijunText(order);
        }
        try {
            f.setSeizomakerId(id(p.getFirst("seizomakerId")));f.setKikibunruiId(id(p.getFirst("kikibunruiId")));f.setTorihikisakiId(id(p.getFirst("torihikisakiId")));
            f.setKikihinmeiNk(p.getFirst("kikihinmeiNk"));f.setKikikatashiki(p.getFirst("kikikatashiki"));
            f.setSeizomakerText(p.getFirst("seizomakerText"));
        }catch(IllegalArgumentException|ArithmeticException e){return List.of("検索条件を選択し直してください。");}
        return List.of();
    }
    private List<String> search(Work w) {
        var f=w.form;w.results=new ArrayList<>();w.partners=new ArrayList<>();f.setSelectedAtsukaikikiId(null);
        if(f.getSeizomakerId()==null&&f.getKikibunruiId()==null&&f.getTorihikisakiId()==null&&str(f.getKikihinmeiNk()).isBlank()&&str(f.getKikikatashiki()).isBlank()&&str(f.getSeizomakerText()).isBlank())return List.of("検索条件は、1項目以上選択して下さい。");
        var results=f.getSeizomakerId()==null&&!str(f.getSeizomakerText()).isBlank()
            ?service.searchAtsukaikiki(null,f.getKikihinmeiNk(),f.getKikikatashiki(),f.getKikibunruiId(),f.getTorihikisakiId(),f.getSeizomakerText())
            :service.searchAtsukaikiki(f.getSeizomakerId(),f.getKikihinmeiNk(),f.getKikikatashiki(),f.getKikibunruiId(),f.getTorihikisakiId());
        var partners=service.getAtsukaikikiPartners(results,f.getTorihikisakiId());w.results=results;w.partners=partners;
        if(results.isEmpty())return List.of("検索結果が１件も存在しません。");
        return List.of();
    }
    private List<String> select(Work w,String raw) {
        Long id;try{id=id(raw);}catch(RuntimeException e){return List.of("検索結果から機器を選択してください。");}
        if(id==null||w.results.stream().noneMatch(m->same(m.get("ATSUKAIKIKI_ID"),id)))return List.of("検索結果から機器を選択してください。");
        var m=service.getAtsukaikikiById(id);
        if(m==null||!one(m.get("SELECTABLE")))return List.of("有効な取引先がある取扱機器を選択してください。");
        var rows=w.form.getMeisaiRows();String cd=str(m.get("KIKIBUNRUI_CD"));
        if(one(m.get("KOTAIKANRI_FLG"))&&rows.stream().anyMatch(r->active(r)&&Objects.equals(cd,str(r.getKikibunruiCd()))))return List.of("同じ分類となる機器は複数登録できません。");
        if(rows.stream().anyMatch(r->active(r)&&Objects.equals(id,r.getAtsukaikikiId())))return List.of("既に同じ型式が指定されています。");
        var r=new MeisaiRowForm();r.setAtsukaikikiId(id);r.setKikikoseiId(w.form.getKikikoseiId());r.setSeizomakerId(id(str(m.get("SEIZOMAKER_ID"))));
        r.setSeizomakerNk(str(m.get("SEIZOMAKER_NK")));r.setKikihinmeiNk(str(m.get("ATSUKAIKIKI_NK")));r.setKikikatashiki(str(m.get("KATASHIKI")));
        r.setKikibunruiCd(cd);r.setOyakikibunruiCd(str(m.get("OYAKIKIBUNRUI_CD")));r.setKikibunruiNk(str(m.get("KIKIBUNRUI_NK")));
        r.setKotaikanriFlg(one(m.get("KOTAIKANRI_FLG"))?"1":"0");r.setControllerFlg(one(m.get("CONTROLLER_FLG"))?"1":"0");
        r.setSuryoNm(1);r.setSuryoNmText("1");r.setNounyuKbn("");r.setBiko("");r.setRowStatus("new");
        r.setHyojijun(rows.stream().filter(Mcm0014uController::active).map(Mcm0014uController::orderForAddition).max(Integer::compare).orElse(0)+1);
        r.setHyojijunText(str(r.getHyojijun()));
        rows.add(r);w.form.setSelectedAtsukaikikiId(id);w.edited=true;return List.of();
    }
    private List<String> delete(Work w,String raw,boolean execute) {
        int at;try{at=Integer.parseInt(raw);}catch(RuntimeException e){return List.of("行が選択されていません。");}
        var rows=w.form.getMeisaiRows();if(at<0||at>=rows.size()||!active(rows.get(at)))return List.of("行が選択されていません。");
        var row=rows.get(at);var errors=service.checkDeleteReference(row.getKikimeisaiId());if(!errors.isEmpty())return errors;
        if(execute) {if(row.getKikimeisaiId()==null)rows.remove(at);else row.setRowStatus("deleted");w.edited=true;}
        return List.of();
    }
    private List<String> validate(Work w) {
        String error=numberError("セット数",w.form.getSetNmText());if(error!=null)return List.of(error);
        var active=w.form.getMeisaiRows().stream().filter(Mcm0014uController::active).toList();
        Set<String> kbns=new HashSet<>(List.of(""));for(var item:comboBoxDataService.getNounyukubunList())kbns.add(item.get("key"));
        for(var row:active) {
            error=orderError(row.getHyojijunText());if(error!=null)return List.of(error);
            error=numberError("数量",row.getSuryoNmText());if(error!=null)return List.of(error);
            // ★修正(#253): バイト数(Shift_JIS換算)ではなく、業務仕様どおり文字数(全角/半角問わず1文字=1カウント)で判定する
            if(str(row.getBiko()).length()>4000)return List.of("備考は4000桁以下で入力してください。");
            if(!kbns.contains(str(row.getNounyuKbn())))return List.of("納入区分を選択し直してください。");
            String parent=str(row.getOyakikibunruiCd());
            if(!parent.equals(str(row.getKikibunruiCd()))&&active.stream().noneMatch(r->parent.equals(str(r.getKikibunruiCd()))))
                return List.of("品名：\""+row.getKikihinmeiNk()+"\"\n型式：\""+row.getKikikatashiki()+"\"\nに該当する親機器が見つかりません。");
        }
        for(var row:w.form.getMeisaiRows())if(!active(row)) {var errors=service.checkDeleteReference(row.getKikimeisaiId());if(!errors.isEmpty())return errors;}
        return List.of();
    }
    private static String numberError(String label,String raw) {
        String v=str(raw).trim();
        if(v.isEmpty())return label+"は1以上の値を入力して下さい。";
        if(!v.matches("-?\\d+"))return label+"は数値で入力してください。";
        BigDecimal n=new BigDecimal(v);if(n.signum()<=0)return label+"は1以上の値を入力して下さい。";
        if(n.compareTo(BigDecimal.valueOf(9999))>0)return label+"は4桁以下で入力してください。";
        return null;
    }
    private Mcm0014uDeliveryDto toDelivery(Work w) {
        var d=new Mcm0014uDeliveryDto();BeanUtils.copyProperties(w.source,d,"meisaiRows");d.setSetNum(Integer.valueOf(w.form.getSetNmText().trim()));
        var rows=new ArrayList<Mcm0014uDeliveryDto.MeisaiRow>();
        for(var r:w.form.getMeisaiRows()) {
            var x=new Mcm0014uDeliveryDto.MeisaiRow();BeanUtils.copyProperties(r,x);
            x.setKikimeisaiId(bd(r.getKikimeisaiId()));x.setKikikoseiId(bd(w.form.getKikikoseiId()));x.setAtsukaikikiId(bd(r.getAtsukaikikiId()));x.setSeizomakerId(bd(r.getSeizomakerId()));
            x.setHyojijun(active(r)?new BigDecimal(r.getHyojijunText().trim()):bd(r.getHyojijun()));x.setSuryoNm(active(r)?new BigDecimal(r.getSuryoNmText().trim()):bd(r.getSuryoNm()));
            x.setKotaikanriFlg(bd(r.getKotaikanriFlg()));x.setControllerFlg(bd(r.getControllerFlg()));x.setDeleteFlg(active(r)?0:1);rows.add(x);
        }
        rows.sort(Comparator.comparing(Mcm0014uDeliveryDto.MeisaiRow::getHyojijun,Comparator.nullsLast(Comparator.naturalOrder())));
        d.setMeisaiRows(rows);return d;
    }
    private static MeisaiRowForm fromDto(Mcm0014uDeliveryDto.MeisaiRow d) {
        var r=new MeisaiRowForm();BeanUtils.copyProperties(d,r);
        r.setKikimeisaiId(asLong(d.getKikimeisaiId()));r.setKikikoseiId(asLong(d.getKikikoseiId()));r.setAtsukaikikiId(asLong(d.getAtsukaikikiId()));r.setSeizomakerId(asLong(d.getSeizomakerId()));
        r.setHyojijun(d.getHyojijun()==null?null:d.getHyojijun().intValue());r.setSuryoNm(d.getSuryoNm()==null?null:d.getSuryoNm().intValue());r.setSuryoNmText(d.getSuryoNm()==null?"":d.getSuryoNm().stripTrailingZeros().toPlainString());
        r.setHyojijunText(d.getHyojijun()==null?"":d.getHyojijun().stripTrailingZeros().toPlainString());
        r.setKotaikanriFlg(one(d.getKotaikanriFlg())?"1":"0");r.setControllerFlg(one(d.getControllerFlg())?"1":"0");
        String status=str(d.getRowStatus()).toLowerCase(Locale.ROOT);
        r.setRowStatus(d.getDeleteFlg()==1||List.of("delete","deleted").contains(status)?"deleted":List.of("new","added").contains(status)?"new":List.of("upd","updated","modified").contains(status)?"modified":"unchanged");return r;
    }
    private static Map<String,Object> answer(Work w,List<String> errors) {
        var a=new HashMap<String,Object>();a.put("errors",errors);if(w!=null)a.put("editToken",w.form.getEditToken());
        if(!errors.isEmpty())a.put("title",errors.get(0).startsWith("検索")?"検索エラー":errors.get(0).contains("既に使用されている為")||errors.get(0).startsWith("見積・契約")?"削除エラー":"入力エラー");return a;
    }
    private static void rotate(Work w){w.form.setEditToken(UUID.randomUUID().toString());}
    private static String orderError(String raw) {
        String v=str(raw).trim();
        if(v.isEmpty())return "Noは必ず入力してください。";
        if(!v.matches("-?\\d+"))return "Noは数値で入力してください。";
        return new BigDecimal(v).abs().compareTo(BigDecimal.valueOf(999999))>0?"Noは6桁以下で入力してください。":null;
    }
    private static int orderForAddition(MeisaiRowForm r) {
        return orderError(r.getHyojijunText())==null?Integer.parseInt(r.getHyojijunText().trim()):r.getHyojijun()==null?0:r.getHyojijun();
    }
    private static String str(Object o){return o==null?"":o.toString();}
    private static Long id(String s){return s==null||s.isBlank()?null:new BigDecimal(s).longValueExact();}
    private static boolean same(Object value,Long id){return value!=null&&new BigDecimal(value.toString()).compareTo(BigDecimal.valueOf(id))==0;}
    private static boolean one(Object value){return value!=null&&new BigDecimal(value.toString()).compareTo(BigDecimal.ONE)==0;}
    private static BigDecimal bd(Object value){return value==null||str(value).isBlank()?null:new BigDecimal(value.toString());}
    private static Long asLong(BigDecimal value){return value==null?null:value.longValueExact();}
    private static boolean active(MeisaiRowForm r){return !"deleted".equals(r.getRowStatus());}
}
