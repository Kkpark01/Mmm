package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpSession;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.dao.DataAccessException;
import org.springframework.http.HttpStatus;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.ServletRequestDataBinder;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.server.ResponseStatusException;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import org.springframework.web.util.UriComponentsBuilder;
import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.common.CustomerScreenSupport;
import com.daifuku.mcm.entity.NonyusakiEntity;
import com.daifuku.mcm.form.Mcm0011uForm;
import com.daifuku.mcm.form.Mcm3007uForm;
import com.daifuku.mcm.service.Mcm0011uService;
import com.daifuku.mcm.service.Mcm3007uService;

/** 納入先マスタ。ID・DTS連携状態・排他情報をセッションで保持する。 */
@Controller
@RequestMapping("/mcm0011u")
public class Mcm0011uController extends BaseController {
    @Autowired private Mcm0011uService service;
    @Autowired private Mcm3007uService contractSearch;
    private static final String SCREEN = "MCM0011U";
    private static final String FORM = "mcm0011u.form";
    private static final String OWNER = "mcm0011u.owner";
    private static final String RETURN = "mcm0011u.returnUrl";
    private static final String ENTITY = "mcm0011u.entity";
    private static final String MISSING = "mcm0011u.dataNotFound";

    @Override protected String getScreenTitle() { return "納入先マスタ"; }
    @Override protected String getFunctionId() { return SCREEN; }
    private String returnTo(HttpSession s) { return Set.of("/mcm0010u", "/mcm3007u")
        .contains(Objects.toString(s.getAttribute(RETURN), "")) ? s.getAttribute(RETURN).toString() : "/mcm0010u"; }
    private boolean canView() {
        try { return hasInspectionAuthority() && service.canView(getLoginUserId(), SCREEN); }
        catch (DataAccessException ex) { logger.warn("MCM0011U 参照権限の取得失敗", ex); return false; }
    }
    private boolean canUpdate() {
        try { return hasUpdateAuthority() && service.canUpdate(getLoginUserId(), SCREEN); }
        catch (DataAccessException ex) { logger.warn("MCM0011U 更新権限の取得失敗", ex); return false; }
    }
    private boolean canPlant() {
        try { return canUpdate() && service.canUpdate(getLoginUserId(), "MCM0012U"); }
        catch (DataAccessException ex) { logger.warn("MCM0012U 更新権限の取得失敗", ex); return false; }
    }
    private Mcm0011uForm current(HttpSession s) {
        if (!Objects.equals(s.getAttribute(OWNER), getLoginUserId()) || !(s.getAttribute(FORM) instanceof Mcm0011uForm f)
                || Boolean.TRUE.equals(s.getAttribute(MISSING))) throw new IllegalStateException(CustomerScreenSupport.EXPIRED);
        return f;
    }
    private void load(HttpSession s, BigDecimal id) {
        s.removeAttribute(ENTITY);
        s.setAttribute(MISSING, false);
        Mcm0011uForm f;
        if (id == null || id.signum() == 0) f = service.initNewForm();
        else {
            NonyusakiEntity entity = service.findByNonyusakiId(id).orElse(null);
            if (entity == null) {
                f = service.initNewForm(); f.setNonyusakiId(id); f.setNewMode(false);
                s.setAttribute(MISSING, true);
            } else { f = service.entityToForm(entity); s.setAttribute(ENTITY, entity); }
        }
        s.setAttribute(FORM, f);
        s.setAttribute(OWNER, getLoginUserId());
        CustomerScreenSupport.rotate(s, SCREEN);
    }
    private String render(Model m, HttpSession s) {
        setCommonAttributes(m, s);
        boolean visible = Objects.equals(s.getAttribute(OWNER), getLoginUserId()) && canView();
        Mcm0011uForm f = visible && s.getAttribute(FORM) instanceof Mcm0011uForm value ? value : service.initNewForm();
        boolean missing = !visible || Boolean.TRUE.equals(s.getAttribute(MISSING));
        boolean update = !missing && canUpdate();
        m.addAttribute("form", f);
        m.addAttribute("entity", visible ? s.getAttribute(ENTITY) : null);
        m.addAttribute("dataNotFound", missing);
        m.addAttribute("canUpdate", update);
        m.addAttribute("hasUnsavedChanges", update && service.hasChanges(f,
            s.getAttribute(ENTITY) instanceof NonyusakiEntity original ? original : new NonyusakiEntity()));
        m.addAttribute("isUpdateAuthority", update);
        m.addAttribute("canAddPlant", update && !f.isNewMode() && canPlant());
        m.addAttribute("returnUrl", returnTo(s));
        m.addAttribute("workflowToken", CustomerScreenSupport.token(s, SCREEN));
        return "mcm0011u/index";
    }
    @GetMapping
    public String index(@RequestParam(required = false) BigDecimal nonyusakiId,
            @RequestParam(required = false) String returnUrl, @RequestParam(defaultValue = "false") boolean resume,
            Model model, HttpSession session) {
        synchronized (session) {
            try {
                if (!canView()) throw new IllegalStateException("参照権限がありません。");
                boolean restore = resume && Objects.equals(session.getAttribute(OWNER), getLoginUserId())
                    && session.getAttribute(FORM) instanceof Mcm0011uForm f
                    && (nonyusakiId == null || CustomerScreenSupport.same(nonyusakiId, f.getNonyusakiId()));
                if (!restore) {
                    session.setAttribute(RETURN, Set.of("/mcm0010u", "/mcm3007u").contains(
                        Objects.toString(returnUrl, "")) ? returnUrl : "/mcm0010u");
                    load(session, nonyusakiId);
                }
                if (Boolean.TRUE.equals(session.getAttribute(MISSING)))
                    model.addAttribute("errorMessage", "指定された納入先データが見つかりません。");
            } catch (IllegalStateException | DataAccessException ex) {
                session.removeAttribute(ENTITY);
                session.setAttribute(FORM, service.initNewForm());
                session.setAttribute(MISSING, true);
                model.addAttribute("errorMessage", error(ex));
            }
            return render(model, session);
        }
    }
    private Mcm0011uForm input(HttpServletRequest request, HttpSession session) {
        CustomerScreenSupport.check(session, SCREEN, request);
        if (!canUpdate()) throw new IllegalStateException("更新権限がありません。");
        Mcm0011uForm f = CustomerScreenSupport.copy(current(session));
        ServletRequestDataBinder binder = new ServletRequestDataBinder(f);
        String[] fields = java.util.Arrays.copyOf(Mcm0011uService.INPUT_FIELDS, Mcm0011uService.INPUT_FIELDS.length + 1);
        fields[fields.length - 1] = "nonyusakiSakujo";
        binder.setAllowedFields(fields);
        binder.setAutoGrowNestedPaths(false);
        // チェックを外した場合も確実に反映する（ID・新規区分・DTSフラグは受け付けない）。
        f.setNonyusakiSakujo(false);
        binder.bind(request);
        if (binder.getBindingResult().hasErrors()) throw new IllegalStateException("入力内容を確認してください。");
        session.setAttribute(FORM, f);
        return f;
    }
    @PostMapping("/save")
    public String save(HttpServletRequest request, HttpSession session, Model model, RedirectAttributes ra) {
        synchronized (session) {
            try {
                Mcm0011uForm f = input(request, session);
                if (f.isNonyusakiSakujo()) {
                    if (!service.delete(f, getLoginUserId())) {
                        f.setNonyusakiSakujo(false);
                        throw new IllegalStateException("既にプラント情報が登録されている為、削除する事が出来ません。");
                    }
                    load(session, BigDecimal.ZERO);
                } else load(session, service.save(f, getLoginUserId()));
                ra.addFlashAttribute("successMessage", "登録が完了しました。");
                return "redirect:/mcm0011u?resume=true";
            } catch (IllegalStateException | DataAccessException ex) {
                model.addAttribute("errorMessage", error(ex));
                return render(model, session);
            }
        }
    }
    @PostMapping("/plantAdd")
    public String plantAdd(HttpServletRequest request, HttpSession session, Model model) {
        synchronized (session) {
            try {
                Mcm0011uForm f = input(request, session);
                if (!canPlant()) throw new IllegalStateException("プラント追加の権限がありません。");
                if (f.isNewMode()) throw new IllegalStateException("先に納入先を登録してください。");
                if (f.isNonyusakiSakujo()) throw new IllegalStateException("納入先削除のチェックを外してください。");
                // 変更ありなら保存、変更なしなら排他確認のみ。検証失敗時は遷移しない。
                load(session, service.save(f, getLoginUserId()));
                Mcm0011uForm saved = current(session);
                // 0012UはLong型で受け取るため、DBのDECIMALの小数部をURLに含めない。
                String id = saved.getNonyusakiId().toBigIntegerExact().toString();
                String back = UriComponentsBuilder.fromPath("/mcm0011u").queryParam("nonyusakiId", id)
                    .queryParam("resume", true).build().encode().toUriString();
                return "redirect:" + UriComponentsBuilder.fromPath("/mcm0012u")
                    .queryParam("nonyusakiId", "{id}").queryParam("nonyusakiCd", "{code}")
                    .queryParam("nonyusakiNk", "{name}").queryParam("returnUrl", "{back}")
                    .encode().buildAndExpand(Map.of("id", id, "code", Objects.toString(saved.getNonyusakiCd(), ""),
                        "name", Objects.toString(saved.getNonyusakiNk(), ""), "back", back)).toUriString();
            } catch (IllegalStateException | DataAccessException ex) {
                model.addAttribute("errorMessage", error(ex));
                return render(model, session);
            }
        }
    }
    @GetMapping("/checkPlantExists") @ResponseBody
    public Map<String, Boolean> checkPlantExists(@RequestParam(required = false) BigDecimal nonyusakiId, HttpSession session) {
        synchronized (session) {
            try {
                if (!canView()) throw new IllegalStateException("参照権限がありません。");
                Mcm0011uForm f = current(session);
                if (nonyusakiId != null && !CustomerScreenSupport.same(nonyusakiId, f.getNonyusakiId()))
                    throw new IllegalStateException(CustomerScreenSupport.EXPIRED);
                return Map.of("hasPlant", !f.isNewMode() && service.countPlantByNonyusakiId(f.getNonyusakiId()) > 0);
            } catch (IllegalStateException ex) { throw new ResponseStatusException(HttpStatus.CONFLICT, ex.getMessage()); }
        }
    }
    @GetMapping("/back")
    public String back(HttpSession session, RedirectAttributes ra) {
        synchronized (session) {
            String target = returnTo(session);
            if ("/mcm3007u".equals(target) && session.getAttribute("MCM3007U_FORM") instanceof Mcm3007uForm previous) {
                Mcm3007uForm refreshed = CustomerScreenSupport.copy(previous);
                try { contractSearch.search(refreshed); }
                catch (RuntimeException ex) {
                    logger.warn("MCM0011Uから戻る際の契約一覧再検索失敗", ex);
                    refreshed.setGrid1Rows(java.util.List.of());
                    refreshed.setGrid2Rows(java.util.List.of());
                    refreshed.setGrid3Rows(java.util.List.of());
                    ra.addFlashAttribute("error", "契約一覧を再表示できませんでした。条件を確認して再検索してください。");
                }
                session.setAttribute("MCM3007U_FORM", refreshed);
                CustomerScreenSupport.rotate(session, "mcm3007u");
            }
            clear(session);
            return "redirect:" + target;
        }
    }
    @GetMapping("/close")
    public String close(HttpSession session) { synchronized (session) { clear(session); return "redirect:/menu"; } }
    private void clear(HttpSession session) {
        for (String key : new String[]{FORM, ENTITY, MISSING, OWNER, RETURN}) session.removeAttribute(key);
        CustomerScreenSupport.rotate(session, SCREEN);
    }
    private String error(RuntimeException ex) {
        if (ex instanceof DataAccessException) {
            logger.error("MCM0011U データ処理失敗", ex);
            return "データを処理できませんでした。入力内容を確認し、画面を開き直して再実行してください。";
        }
        return ex.getMessage();
    }
}
