package com.daifuku.mcm.controller;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.form.Mcm3005uForm;
import com.daifuku.mcm.service.Mcm3005uService;
import jakarta.servlet.http.HttpSession;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.WebDataBinder;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import java.math.BigDecimal;
import java.util.Map;
import java.util.Objects;
import java.util.UUID;

/** 作業予定登録。初回の受け渡しと登録後の再表示を区別し、エラー時は入力を保持する。 */
@Controller
@RequestMapping("/mcm3005u")
public class Mcm3005uController extends BaseController {
    @Autowired private Mcm3005uService service;
    private static final String FORM = "MCM3005U_FORM";
    private static final String TOKEN = "MCM3005U_EDIT_TOKEN";
    private static final String DELIVERY = "MCM3005U_DELIVERY";
    private static final String STALE = "他のユーザがデータを変更した可能性があります。処理をやり直してください。";

    @InitBinder("form")
    public void bindEditableFields(WebDataBinder binder) {
        binder.setAllowedFields("yoteiDt", "kanryoFlg", "naiyo", "biko", "deleteFlg");
    }

    @GetMapping
    public String index(@RequestParam(required=false) BigDecimal torokuId, Model model, HttpSession session) {
        Object delivery = session.getAttribute(DELIVERY);
        session.removeAttribute(DELIVERY);
        if (delivery instanceof Map<?, ?> values && values.get("torokuId") instanceof Number n) {
            torokuId = new BigDecimal(n.toString());
        }
        session.removeAttribute(FORM);
        session.removeAttribute(TOKEN);
        try {
            Mcm3005uForm form = service.load(torokuId == null ? BigDecimal.ZERO : torokuId);
            session.setAttribute(FORM, form);
            session.setAttribute(TOKEN, UUID.randomUUID().toString());
            return render(form, model, session);
        } catch (IllegalStateException ex) {
            model.addAttribute("error", ex.getMessage());
            model.addAttribute("messageTitle", "検索エラー");
            return render(new Mcm3005uForm(), model, session);
        }
    }

    private String render(Mcm3005uForm form, Model model, HttpSession session) {
        model.addAttribute("form", form);
        // 入力エラー後も、画面表示値ではなく編集開始時の値と比較する。
        model.addAttribute("originalForm", session.getAttribute(FORM));
        model.addAttribute("editToken", session.getAttribute(TOKEN));
        model.addAttribute("canEdit", hasUpdateAuthority() && session.getAttribute(FORM) != null);
        return "mcm3005u/index";
    }

    private Mcm3005uForm prepare(Mcm3005uForm form, BindingResult binding, String token, HttpSession session) {
        requireUpdateAuthority();
        Object stored = session.getAttribute(FORM);
        if (!(stored instanceof Mcm3005uForm original) || token == null || !token.equals(session.getAttribute(TOKEN))) {
            throw new IllegalStateException(STALE);
        }
        // ID・新規区分・添付一覧は、画面から送信された値を使用しない。
        form.setTorokuId(original.getTorokuId());
        form.setNewRecord(original.isNewRecord());
        form.setTenpuList(original.getTenpuList());
        if (binding.hasErrors()) throw new IllegalArgumentException("入力値の形式を確認してください。");
        if (form.isDeleteFlg()) {
            if (original.isNewRecord()) throw new IllegalArgumentException("削除対象のレコードが不正です。");
            service.validateDelete(form.getTorokuId());
        } else {
            service.validate(form);
            if (!original.isNewRecord() && Objects.equals(original.getYoteiDt(), form.getYoteiDt())
                    && original.getKanryoFlg() == form.getKanryoFlg()
                    && Objects.equals(original.getNaiyo(), form.getNaiyo())
                    && Objects.equals(original.getBiko(), form.getBiko())) {
                throw new IllegalArgumentException("値が変更されていません。");
            }
        }
        return form;
    }

    @PostMapping("/validate")
    @ResponseBody
    public Map<String, String> validate(@ModelAttribute("form") Mcm3005uForm form, BindingResult binding,
            @RequestParam(required=false) String editToken, HttpSession session) {
        synchronized (session) {
            try {
                prepare(form, binding, editToken, session);
                return Map.of("error", "");
            } catch (IllegalArgumentException | IllegalStateException ex) {
                return Map.of("error", ex.getMessage(), "title", form.isDeleteFlg() ? "削除エラー" : "入力エラー");
            }
        }
    }

    @PostMapping("/save")
    public String save(@ModelAttribute("form") Mcm3005uForm form, BindingResult binding,
            @RequestParam(required=false) String editToken, HttpSession session, Model model, RedirectAttributes ra) {
        requireUpdateAuthority();
        synchronized (session) {
            try {
                prepare(form, binding, editToken, session);
                if (form.isDeleteFlg()) service.delete(form.getTorokuId());
                else service.save(form, getLoginUserId(), form.isNewRecord());
                session.removeAttribute(FORM);
                session.removeAttribute(TOKEN);
                ra.addFlashAttribute("message", "登録を完了しました。");
                ra.addFlashAttribute("messageTitle", "完了");
                return form.isDeleteFlg() ? "redirect:/mcm3005u" : "redirect:/mcm3005u?torokuId=" + form.getTorokuId().toPlainString();
            } catch (IllegalArgumentException | IllegalStateException ex) {
                model.addAttribute("error", ex.getMessage());
                model.addAttribute("messageTitle", form.isDeleteFlg() ? "削除エラー" : "入力エラー");
            } catch (Exception ex) {
                logger.error("作業予定登録に失敗しました", ex);
                model.addAttribute("error", "登録処理に失敗しました。入力内容を確認して再度実行してください。");
                model.addAttribute("messageTitle", "エラー");
            }
            // トークンが一致しない古い画面では、別レコードへの再送を許さない。
            if (editToken == null || !editToken.equals(session.getAttribute(TOKEN))) {
                model.addAttribute("form", form);
                model.addAttribute("canEdit", false);
                model.addAttribute("editToken", "");
                return "mcm3005u/index";
            }
            return render(form, model, session);
        }
    }

    // 旧画面からの削除要求も同じ権限・編集トークン・添付チェックを通す。
    @PostMapping("/delete")
    public String delete(@ModelAttribute("form") Mcm3005uForm form, BindingResult binding,
            @RequestParam(required=false) String editToken, HttpSession session, Model model, RedirectAttributes ra) {
        form.setDeleteFlg(true);
        return save(form, binding, editToken, session, model, ra);
    }

    @Override protected String getScreenTitle() { return "作業予定登録"; }
    @Override protected String getFunctionId() { return "MCM3005U"; }
}
