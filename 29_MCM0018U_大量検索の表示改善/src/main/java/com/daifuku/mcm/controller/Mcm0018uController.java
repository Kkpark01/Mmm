package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jakarta.servlet.http.HttpSession;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.constants.Mcm0018uConstants;
import com.daifuku.mcm.dto.UserInfo;
import com.daifuku.mcm.form.Mcm0018uForm;
import com.daifuku.mcm.form.Mcm0018uForm.AtsukaikikiRowForm;
import com.daifuku.mcm.form.Mcm0018uForm.KoseiRowForm;
import com.daifuku.mcm.service.Mcm0018uService;
import com.daifuku.mcm.service.Mcm0018uService.DataNotChangedException;

@Controller
@RequestMapping("/mcm0018u")
public class Mcm0018uController extends BaseController {

    private static final Logger log = LoggerFactory.getLogger(Mcm0018uController.class);

    @Autowired
    private Mcm0018uService service;

    @GetMapping
    public String index(@ModelAttribute("form") Mcm0018uForm form, Model model, HttpSession session) {
        if (form == null) { form = new Mcm0018uForm(); }
        ensurePlaceholderRows(form);
        model.addAttribute("form", form);
        // 【移植】元VB: DataGridView・登録・行削除ボタンはAuthorityIsThrough=False（参照権限のみで非活性）。
        //   検索系（ATUKAIKIKI_NKTextBox, SearchButton, SEIZOUMAKER_NKCombobox, KIKIBUNRUI_NKCombobox）は
        //   AuthorityIsThrough=Trueのため権限に関わらず常に活性のまま。
        setCommonAttributes(model, session);
        loadComboBoxData(model);
        return "mcm0018u/index";
    }

    /**
     * 【新規】行削除ボタン押下時の参照整合性チェック（軽量API）。
     * 従来はチェックが登録ボタン押下時（updateAll/validate）にしか行われず、
     * 行削除ボタン押下時点では参照ありのデータもそのまま削除保留状態になっていた。
     * 削除ボタン押下時にその場でチェックし、参照がある場合は即座にエラー表示する。
     *
     * 【修正】取引先（構成）グリッドの行を先に削除保留（rowStatus="deleted"）にしていても、
     *   DB上はまだ削除されていないため件数チェックに引っかかり、取扱機器（親）行が
     *   削除できない不具合があった。画面上で削除保留中の構成IDを excludeKoseiIds として
     *   受け取り、件数から除外する。
     */
    @GetMapping("/checkDeleteAtsukaikiki")
    @ResponseBody
    public ResponseEntity<Map<String, Object>> checkDeleteAtsukaikiki(
            @RequestParam("atsukaikikiId") BigDecimal atsukaikikiId,
            @RequestParam(value = "excludeKoseiIds", required = false) List<BigDecimal> excludeKoseiIds) {
    	boolean hasRef = service.countKoseiByAtsukaikikiId(atsukaikikiId, excludeKoseiIds) > 0;
    	Map<String, Object> body = new HashMap<>();
    	body.put("hasReference", hasRef);
    	if (hasRef) body.put("message", Mcm0018uConstants.MSG_ERR_DELETE_HAS_KOSEI);
        return ResponseEntity.ok(body);
    }

    @GetMapping("/checkDeleteKosei")
    @ResponseBody
    public ResponseEntity<Map<String, Object>> checkDeleteKosei(
            @RequestParam("atsukaikikikoseiId") BigDecimal atsukaikikikoseiId) {
        boolean hasRef = service.hasKoseiReference(atsukaikikikoseiId);
        Map<String, Object> body = new HashMap<>();
        body.put("hasReference", hasRef);
        if (hasRef) body.put("message", Mcm0018uConstants.MSG_ERR_DELETE_REFERENCED_KOTAIKANRI);
        return ResponseEntity.ok(body);
    }

    @PostMapping("/search")
    public String search(@ModelAttribute Mcm0018uForm form,
                         @RequestParam(value = "discardConfirmed", required = false, defaultValue = "false")
                         boolean discardConfirmed,
                         Model model, HttpSession session) {
        setCommonAttributes(model, session);
        if (form.isChangeStatus() && !discardConfirmed) {
            model.addAttribute("errors", List.of(Mcm0018uConstants.MSG_CONFIRM_DISCARD_CHANGES));
            model.addAttribute("errorType", "input");
            ensurePlaceholderRows(form);
            model.addAttribute("form", form);
            loadComboBoxData(model);
            return "mcm0018u/index";
        }
        if (form.isAllSearchConditionEmpty()) {
            model.addAttribute("errors", List.of(Mcm0018uConstants.MSG_ERR_SEARCH_CONDITION_REQUIRED));
            model.addAttribute("errorType", "search");
            ensurePlaceholderRows(form);
            model.addAttribute("form", form);
            loadComboBoxData(model);
            return "mcm0018u/index";
        }
        List<AtsukaikikiRowForm> parentRows = service.searchAtsukaikiki(form);
        form.setAtsukaikikiRows(parentRows);
        if (parentRows.isEmpty()) {
            model.addAttribute("message", Mcm0018uConstants.MSG_NO_DATA);
            model.addAttribute("messageType", "info");
            form.setKoseiRows(new ArrayList<>());
            form.setKoseiRowsMap(new HashMap<>());
            form.setSelectedParentIndex(-1);
            form.setSelectedAtsukaikikiId(null);
        } else {
            Map<BigDecimal, List<KoseiRowForm>> koseiMap = service.searchKoseiForParents(parentRows);
            form.setKoseiRowsMap(koseiMap);
            BigDecimal firstId = parentRows.get(0).getAtsukaikikiId();
            form.setSelectedParentIndex(0);
            form.setSelectedAtsukaikikiId(firstId);
            form.setKoseiRows(koseiMap.getOrDefault(firstId, new ArrayList<>()));
        }
        form.setChangeStatus(false);
        ensurePlaceholderRows(form);
        model.addAttribute("form", form);
        loadComboBoxData(model);
        return "mcm0018u/index";
    }

    @PostMapping("/loadKosei")
    public String loadKosei(@ModelAttribute Mcm0018uForm form,
                            @RequestParam("parentIndex") int parentIndex,
                            Model model, HttpSession session) {
        setCommonAttributes(model, session);
        form.setSelectedParentIndex(parentIndex);
        if (parentIndex >= 0 && parentIndex < form.getAtsukaikikiRows().size()) {
            BigDecimal id = form.getAtsukaikikiRows().get(parentIndex).getAtsukaikikiId();
            form.setSelectedAtsukaikikiId(id);
            List<KoseiRowForm> cached = form.getKoseiRowsMap() != null
                    ? form.getKoseiRowsMap().get(id) : null;
            if (cached != null) form.setKoseiRows(cached);
            else form.setKoseiRows(service.findKoseiByAtsukaikikiId(id));
        }
        ensurePlaceholderRows(form);
        model.addAttribute("form", form);
        loadComboBoxData(model);
        return "mcm0018u/index";
    }

    /**
     * 【性能改善（MCM0015U準拠）】従来は application/x-www-form-urlencoded 形式の
     *   通常フォーム送信で送信しており、行数の多いフォームでSpring MVCのバインディングが
     *   極端に遅くなる問題があった。@RequestBody + JSON化（Ajax送信）に変更し、
     *   Jacksonの線形時間デシリアライズに切り替えることでこの遅延を解消する。
     *   画面側はJSONレスポンスを受け取った後、画面を再読込する（location.reload()）。
     */
    @PostMapping(value = "/update", consumes = "application/json")
    @ResponseBody
    public ResponseEntity<Map<String, Object>> update(@RequestBody Mcm0018uForm form,
                                                        HttpSession session) {
        requireUpdateAuthority();
        autoMarkAddedRows(form);
        List<String> errors = service.validate(form);
        if (!errors.isEmpty()) {
            return jsonError(errors, "warning");
        }
        String loginUser = getLoginUser(session);
        try {
            service.updateAll(form, loginUser);
        } catch (DataNotChangedException e) {
            return jsonError(List.of(e.getMessage()), "input");
        } catch (IllegalStateException e) {
            // 想定外のエラーもwarningダイアログで表示（500ページ防止）
            log.warn("IllegalStateException in updateAll: {}", e.getMessage(), e);
            return jsonError(List.of(e.getMessage()), "warning");
        }
        Map<String, Object> body = new HashMap<>();
        body.put("message", Mcm0018uConstants.MSG_UPDATE_COMPLETE);
        return ResponseEntity.ok(body);
    }

    /** エラー応答（HTTP 400 + JSON）を生成する。【MCM0015U準拠】 */
    private ResponseEntity<Map<String, Object>> jsonError(List<String> errors, String errorType) {
        Map<String, Object> body = new HashMap<>();
        body.put("message", String.join("\n", errors));
        body.put("errorType", errorType);
        return ResponseEntity.badRequest().body(body);
    }

    @PostMapping("/deleteAtsukaikiki")
    public String deleteAtsukaikiki(@ModelAttribute Mcm0018uForm form,
                                    @RequestParam("deleteIndex") int deleteIndex,
                                    RedirectAttributes ra) {
        requireUpdateAuthority();
        if (deleteIndex < 0 || deleteIndex >= form.getAtsukaikikiRows().size()) {
            ra.addFlashAttribute("errors", List.of(Mcm0018uConstants.MSG_ERR_NO_ROW_SELECTED));
            ra.addFlashAttribute("errorType", "input");
            ra.addFlashAttribute("form", form);
            return "redirect:/mcm0018u";
        }
        AtsukaikikiRowForm targetRow = form.getAtsukaikikiRows().get(deleteIndex);
        if ("added".equals(targetRow.getRowStatus())) {
            form.getAtsukaikikiRows().remove(deleteIndex);
        } else {
            if (targetRow.getAtsukaikikiId() != null) {
                List<BigDecimal> excludeKoseiIds = new ArrayList<>();
                if (form.getKoseiRows() != null) {
                    for (KoseiRowForm k : form.getKoseiRows()) {
                        if ("deleted".equals(k.getRowStatus()) && k.getAtsukaikikikoseiId() != null) {
                            excludeKoseiIds.add(k.getAtsukaikikikoseiId());
                        }
                    }
                }
                long koseiCount = service.countKoseiByAtsukaikikiId(targetRow.getAtsukaikikiId(), excludeKoseiIds);
                if (koseiCount > 0) {
                    ra.addFlashAttribute("errors", List.of(Mcm0018uConstants.MSG_ERR_DELETE_HAS_KOSEI));
                    ra.addFlashAttribute("errorType", "warning");
                    ra.addFlashAttribute("form", form);
                    return "redirect:/mcm0018u";
                }
            }
            targetRow.setRowStatus("deleted");
        }
        ensurePlaceholderRows(form);
        ra.addFlashAttribute("form", form);
        return "redirect:/mcm0018u";
    }

    @PostMapping("/deleteKosei")
    public String deleteKosei(@ModelAttribute Mcm0018uForm form,
                              @RequestParam("deleteKoseiIndex") int deleteKoseiIndex,
                              RedirectAttributes ra) {
        requireUpdateAuthority();
        if (deleteKoseiIndex < 0 || deleteKoseiIndex >= form.getKoseiRows().size()) {
            ra.addFlashAttribute("errors", List.of(Mcm0018uConstants.MSG_ERR_NO_ROW_SELECTED));
            ra.addFlashAttribute("errorType", "input");
            ra.addFlashAttribute("form", form);
            return "redirect:/mcm0018u";
        }
        KoseiRowForm targetRow = form.getKoseiRows().get(deleteKoseiIndex);
        if ("added".equals(targetRow.getRowStatus())) {
            form.getKoseiRows().remove(deleteKoseiIndex);
        } else {
            if (targetRow.getAtsukaikikikoseiId() != null) {
                int refCount = service.countKikikotaikanriByKoseiId(targetRow.getAtsukaikikikoseiId());
                if (refCount > 0) {
                    ra.addFlashAttribute("errors", List.of(Mcm0018uConstants.MSG_ERR_DELETE_REFERENCED_KOTAIKANRI));
                    ra.addFlashAttribute("errorType", "warning");
                    ra.addFlashAttribute("form", form);
                    return "redirect:/mcm0018u";
                }
            }
            targetRow.setRowStatus("deleted");
        }
        ensurePlaceholderRows(form);
        ra.addFlashAttribute("form", form);
        return "redirect:/mcm0018u";
    }

    // ================================================================
    // ヘルパー
    // ================================================================
    private void autoMarkAddedRows(Mcm0018uForm form) {
        if (form.getAtsukaikikiRows() != null) {
            for (AtsukaikikiRowForm r : form.getAtsukaikikiRows()) {
                if (r.getAtsukaikikiId() == null
                        && isUntouchedStatus(r.getRowStatus())
                        && !isEffectivelyEmptyParent(r)) {
                    r.setRowStatus("added");
                }
            }
        }
        if (form.getKoseiRows() != null) {
            for (KoseiRowForm k : form.getKoseiRows()) {
                if (k.getAtsukaikikikoseiId() == null
                        && isUntouchedStatus(k.getRowStatus())
                        && !isEffectivelyEmptyKosei(k)) {
                    k.setRowStatus("added");
                }
            }
        }
    }

    private static boolean isUntouchedStatus(String rs) {
        return rs == null || rs.isEmpty() || "unchanged".equals(rs);
    }

    private void populateDisplayFields(Mcm0018uForm form) {
        Map<String, String> seizomakerMap = new HashMap<>();
        for (Map<String, Object> m : service.getSeizomakerComboList()) {
            Object id = m.get("SEIZOMAKER_ID");
            Object nk = m.get("SEIZOMAKER_NK");
            if (id != null) seizomakerMap.put(id.toString(), nk != null ? nk.toString() : "");
        }
        Map<String, String> kikibunruiMap = new HashMap<>();
        for (Map<String, Object> m : service.getKikibunruiComboList()) {
            Object id = m.get("KIKIBUNRUI_ID");
            Object nk = m.get("KIKIBUNRUI_NK");
            if (id != null) kikibunruiMap.put(id.toString(), nk != null ? nk.toString() : "");
        }
        if (form.getAtsukaikikiRows() != null) {
            for (AtsukaikikiRowForm r : form.getAtsukaikikiRows()) {
                if (r.getSeizomakerId() != null
                        && (r.getSeizomakerNk() == null || r.getSeizomakerNk().isEmpty())) {
                    r.setSeizomakerNk(seizomakerMap.get(r.getSeizomakerId().toString()));
                }
                if (r.getKikibunruiId() != null
                        && (r.getKikibunruiNk() == null || r.getKikibunruiNk().isEmpty())) {
                    r.setKikibunruiNk(kikibunruiMap.get(r.getKikibunruiId().toString()));
                }
            }
        }
        Map<String, String> torihikisakiMap = new HashMap<>();
        for (Map<String, Object> m : service.getTorihikisakiComboList()) {
            Object id = m.get("TORIHIKISAKI_ID");
            Object nk = m.get("TORIHIKISAKI_NK");
            if (id != null) torihikisakiMap.put(id.toString(), nk != null ? nk.toString() : "");
        }
        if (form.getKoseiRows() != null) {
            for (KoseiRowForm k : form.getKoseiRows()) {
                if (k.getTorihikisakiId() != null
                        && (k.getTorihikisakiNk() == null || k.getTorihikisakiNk().isEmpty())) {
                    k.setTorihikisakiNk(torihikisakiMap.get(k.getTorihikisakiId().toString()));
                }
            }
        }
    }

    private void ensurePlaceholderRows(Mcm0018uForm form) {
        List<AtsukaikikiRowForm> pRows = form.getAtsukaikikiRows();
        if (pRows == null) {
            pRows = new ArrayList<>();
            form.setAtsukaikikiRows(pRows);
        }
        while (!pRows.isEmpty() && isEffectivelyEmptyParent(pRows.get(pRows.size() - 1))) {
            pRows.remove(pRows.size() - 1);
        }
        pRows.add(new AtsukaikikiRowForm());
        List<KoseiRowForm> kRows = form.getKoseiRows();
        if (kRows == null) {
            kRows = new ArrayList<>();
            form.setKoseiRows(kRows);
        }
        while (!kRows.isEmpty() && isEffectivelyEmptyKosei(kRows.get(kRows.size() - 1))) {
            kRows.remove(kRows.size() - 1);
        }
        kRows.add(new KoseiRowForm());
    }

    private boolean isEffectivelyEmptyParent(AtsukaikikiRowForm r) {
        if (r == null) return true;
        if (r.getAtsukaikikiId() != null) return false;
        if (r.getSeizomakerId() != null) return false;
        if (r.getAtsukaikikiNk() != null && !r.getAtsukaikikiNk().trim().isEmpty()) return false;
        if (r.getKatashiki() != null && !r.getKatashiki().trim().isEmpty()) return false;
        if (r.getKikibunruiId() != null) return false;
        if (r.getUpskokanshuki() != null && !r.getUpskokanshuki().trim().isEmpty()) return false;
        if (r.getControllerFlg() != null && r.getControllerFlg().intValue() != 0) return false;
        if (r.getControllerKin() != null) return false;
        if (r.getYukoFlg() != null && !r.getYukoFlg().isEmpty() && !"0".equals(r.getYukoFlg())) return false;
        if (r.getMitsumorihyojiFlg() != null && r.getMitsumorihyojiFlg().intValue() != 0) return false;
        if (r.getDummyFlg() != null && r.getDummyFlg().intValue() != 0) return false;
        if (r.getBiko() != null && !r.getBiko().trim().isEmpty()) return false;
        String rs = r.getRowStatus();
        if (rs != null && !rs.isEmpty() && !"unchanged".equals(rs)) return false;
        return true;
    }

    private boolean isEffectivelyEmptyKosei(KoseiRowForm k) {
        if (k == null) return true;
        if (k.getAtsukaikikikoseiId() != null) return false;
        if (k.getTorihikisakiId() != null) return false;
        if (k.getMakerhosyuDt() != null && !k.getMakerhosyuDt().trim().isEmpty()) return false;
        if (k.getKeiyakukanokikan() != null && !k.getKeiyakukanokikan().trim().isEmpty()) return false;
        if (k.getKisanbiKbn() != null && !k.getKisanbiKbn().trim().isEmpty()) return false;
        if (k.getHosyukeiyakutaisyoFlg() != null && k.getHosyukeiyakutaisyoFlg().intValue() != 0) return false;
        if (k.getBiko() != null && !k.getBiko().trim().isEmpty()) return false;
        String rs = k.getRowStatus();
        if (rs != null && !rs.isEmpty() && !"unchanged".equals(rs)) return false;
        return true;
    }

    private void executeSearch(Mcm0018uForm form) {
        if (form.isAllSearchConditionEmpty()) {
            if (form.getAtsukaikikiRows() != null && !form.getAtsukaikikiRows().isEmpty()) {
                log.warn("executeSearch: search conditions empty but existing rows found - preserving");
                return;
            }
            form.setSelectedParentIndex(-1);
            form.setSelectedAtsukaikikiId(null);
            return;
        }
        List<AtsukaikikiRowForm> parentRows = service.searchAtsukaikiki(form);
        form.setAtsukaikikiRows(parentRows);
        if (parentRows.isEmpty()) {
            form.setKoseiRows(new ArrayList<>());
            form.setKoseiRowsMap(new HashMap<>());
            form.setSelectedParentIndex(-1);
            form.setSelectedAtsukaikikiId(null);
        } else {
            Map<BigDecimal, List<KoseiRowForm>> koseiMap = service.searchKoseiForParents(parentRows);
            form.setKoseiRowsMap(koseiMap);
            BigDecimal selId = form.getSelectedAtsukaikikiId();
            int idx = -1;
            if (selId != null) {
                for (int i = 0; i < parentRows.size(); i++) {
                    if (selId.equals(parentRows.get(i).getAtsukaikikiId())) { idx = i; break; }
                }
            }
            if (idx < 0) { idx = 0; selId = parentRows.get(0).getAtsukaikikiId(); }
            form.setSelectedParentIndex(idx);
            form.setSelectedAtsukaikikiId(selId);
            form.setKoseiRows(koseiMap.getOrDefault(selId, new ArrayList<>()));
        }
    }

    private void loadComboBoxData(Model model) {
        var makers = service.getSeizomakerComboList();
        model.addAttribute("seizomakerList", makers);
        model.addAttribute("seizomakerNames", comboNames(makers,"SEIZOMAKER_ID","SEIZOMAKER_NK"));
        var categories = service.getKikibunruiComboList();
        model.addAttribute("kikibunruiList", categories);
        model.addAttribute("kikibunruiNames", comboNames(categories,"KIKIBUNRUI_ID","KIKIBUNRUI_NK"));
        model.addAttribute("torihikisakiList", service.getTorihikisakiComboList());
        model.addAttribute("kisanbiKbnList", service.getKisanbiKbnComboList());
    }

    /**
     * ログイン中ユーザーのログインIDを取得する。
     * 【修正】従来は session.getAttribute("loginUser") という誤ったキー名で取得しており、
     *   実際のセッション格納キー（AppConstants.SESSION_USER_INFO="loginUserInfo"）と
     *   一致していなかったため常にnullとなり、"SYSTEM"にフォールバックしていた
     *   （更新者が常にSYSTEMとして記録される不具合の原因。MCM0016Uと同一パターン）。
     *   BaseController.getLoginUserInfo(session)（正しいキーで取得）を使用する。
     */
    private String getLoginUser(HttpSession session) {
        UserInfo userInfo = getLoginUserInfo(session);
        if (userInfo != null && userInfo.getUserId() != null && !userInfo.getUserId().trim().isEmpty()) {
            return userInfo.getUserId();
        }
        return "SYSTEM";
    }

    @Override protected String getScreenTitle() { return Mcm0018uConstants.SCREEN_NAME; }
    @Override protected String getFunctionId() { return Mcm0018uConstants.SCREEN_ID; }
    private Map<String,String> comboNames(List<Map<String,Object>> rows,String id,String label) {
        Map<String,String> names=new HashMap<>();
        for(var row:rows) if(row.get(id)!=null) names.put(new BigDecimal(row.get(id).toString()).stripTrailingZeros().toPlainString(), java.util.Objects.toString(row.get(label),""));
        return names;
    }
}
