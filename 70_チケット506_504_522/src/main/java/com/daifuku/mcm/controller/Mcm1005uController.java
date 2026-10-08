/**
 * 【変換元】Mcm1005uScreen.vb / Mcm1005uTabControl.vb
 * 【説明】取引先契約内容（MCM1005U）コントローラ
 *
 * 元イベント対応:
 *   Mcm1005uScreen_Load             → index (GET)
 *   UpdateButton_Click               → save (POST)
 *   SinseiButton_Click               → sinsei (POST)
 *   FileUpButton_Click / FileDelButton_Click → tenpuUpload / tenpuDelete (POST)
 *   KikanTabControl SelectedIndexChanged → loadKikan (GET, AJAX)
 */
package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpSession;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import org.springframework.web.util.UriComponentsBuilder;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.entity.McmTkKeiyakuEntity;
import com.daifuku.mcm.entity.McmTkKikanEntity;
import com.daifuku.mcm.entity.McmTkSiharaiEntity;
import com.daifuku.mcm.entity.McmTkTenkenEntity;
import com.daifuku.mcm.entity.McmTkTenpuEntity;
import com.daifuku.mcm.exception.InputCheckException;
import com.daifuku.mcm.form.Mcm1006uForm;
import com.daifuku.mcm.service.Mcm1005uService;

@Controller
@RequestMapping("/mcm1005u")
public class Mcm1005uController extends BaseController {
    private static String draftKey(BigDecimal id) { return "MCM1005U_FORM_" + (id == null ? "0" : id.stripTrailingZeros().toPlainString()); }
    private Mcm1005uService.SelectionDraft draft(HttpSession session, BigDecimal id) {
        Object value = session.getAttribute(draftKey(id));
        return value instanceof Mcm1005uService.SelectionDraft d && d.user().equals(getLoginUserId()) ? d : null;
    }

    @PostMapping("/discardSelection")
    public String discardSelection(@RequestParam BigDecimal tkKeiyakuId, @RequestParam String draftToken,
            HttpSession session) {
        synchronized (session) {
            var pending = draft(session, tkKeiyakuId);
            if (pending == null || !pending.token().equals(draftToken)) throw new InputCheckException("画面の情報が更新されています。開き直してください。");
            session.removeAttribute(draftKey(tkKeiyakuId));
        }
        return tkKeiyakuId.signum() == 0 ? close(tkKeiyakuId, session) : backUrl(tkKeiyakuId, null, null, -1);
    }

    @PostMapping("/releaseLock")
    public String releaseLock(@RequestParam BigDecimal tkKeiyakuId, HttpSession session, RedirectAttributes ra) {
        if (!canUpdate()) throw new com.daifuku.mcm.exception.AuthorityException("権限がないため実行できません。");
        session.removeAttribute(draftKey(tkKeiyakuId));
        service.releaseLock(tkKeiyakuId, getLoginUserId(), session);
        ra.addFlashAttribute("message", "管理者用ロックを解除しました。最新の登録内容を表示します。");
        return backUrl(tkKeiyakuId, null, null, 0);
    }

    @GetMapping("/close")
    public String close(@RequestParam BigDecimal tkKeiyakuId, HttpSession session) {
        service.clearUnlock(session);
        session.removeAttribute("MCM1005U_ENTRY");
        session.removeAttribute(draftKey(tkKeiyakuId));
        Object target = session.getAttribute("MCM1005U_RETURN_" + tkKeiyakuId);
        session.removeAttribute("MCM1005U_RETURN_" + tkKeiyakuId);
        return "redirect:" + (target == null ? "/menu" : target);
    }
    @GetMapping("/report")
    public org.springframework.http.ResponseEntity<?> contractReport(@RequestParam BigDecimal tkKeiyakuId,
            @RequestParam BigDecimal tkKikanId, HttpSession session) {
        try {
            if (!hasUpdateAuthority()) throw new com.daifuku.mcm.exception.AuthorityException("権限がないため実行できません。");
            var user = getLoginUserInfo(session);
            byte[] bytes = service.contractReport(tkKeiyakuId, tkKikanId, getLoginUserId(), user == null ? "" : user.getUserLastName());
            return org.springframework.http.ResponseEntity.ok()
                    .header("Cache-Control", "no-store")
                    .header("Content-Disposition", org.springframework.http.ContentDisposition.attachment().filename("契約手続依頼書.xls", java.nio.charset.StandardCharsets.UTF_8).build().toString())
                    .contentType(org.springframework.http.MediaType.parseMediaType("application/vnd.ms-excel")).body(bytes);
        } catch (com.daifuku.mcm.exception.McmBusinessException | com.daifuku.mcm.exception.InputCheckException ex) {
            return org.springframework.http.ResponseEntity.badRequest().body(java.util.Map.of("message", ex.getMessage()));
        } catch (Exception ex) {
            logger.error("MCM1003P 帳票出力失敗", ex);
            return org.springframework.http.ResponseEntity.internalServerError().body(java.util.Map.of("message", "契約手続依頼書を出力できませんでした。管理者に確認してください。"));
        }
    }

    @Autowired
    private Mcm1005uService service;

    @Autowired
    private com.daifuku.mcm.service.Mcm1005uAttachmentService attachments;

    @GetMapping("/attachment/download")
    public org.springframework.http.ResponseEntity<org.springframework.core.io.Resource> downloadAttachment(
            @RequestParam BigDecimal tkKeiyakuId, @RequestParam BigDecimal tkTenpuId) throws java.io.IOException {
        if (!hasInspectionAuthority()) throw new com.daifuku.mcm.exception.AuthorityException("権限がないため実行できません。");
        var row = attachments.find(tkKeiyakuId, tkTenpuId, getLoginUserId());
        var file = attachments.file(row);
        return org.springframework.http.ResponseEntity.ok()
                .contentType(org.springframework.http.MediaType.APPLICATION_OCTET_STREAM)
                .header(org.springframework.http.HttpHeaders.CONTENT_DISPOSITION,
                        org.springframework.http.ContentDisposition.attachment().filename(row.getTenpufileNk(), java.nio.charset.StandardCharsets.UTF_8).build().toString())
                .contentLength(java.nio.file.Files.size(file)).body(new org.springframework.core.io.FileSystemResource(file));
    }

    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("uuuu/MM/dd").withResolverStyle(java.time.format.ResolverStyle.STRICT);

    /** Preserve raw inputs even if Spring cannot bind a numeric/date parameter. */
    @ModelAttribute
    public void prepareRecovery(HttpServletRequest request) {
        if (!"POST".equals(request.getMethod())) return;
        String rawId = request.getParameter("tkKeiyakuId");
        if (rawId == null || !rawId.matches("[0-9]+")) return;
        int index = 0;
        try { index = Math.max(0, Integer.parseInt(request.getParameter("selectedKikanIndex"))); }
        catch (NumberFormatException ignored) { }
        request.setAttribute("mcmRecoveryUrl", backUrl(new BigDecimal(rawId), request.getParameter("torihikisakiNk"),
                request.getParameter("nonyusakiNk"), index).substring(9));
        request.setAttribute("mcmSubmittedValues", request.getParameterMap());
    }

    /**
     * 【移植】元VB: CPCoreUserControl.GetAuthorityDivision() 相当。
     *   「取引先契約関連」権限が「作成」でなければ登録・申請・破棄・機器選定起動を許可しない。
     */
    private boolean canUpdate() {
        return hasUpdateAuthority() && service.canUpdate(getLoginUserId());
    }

    // ===================================================================
    // 初期表示
    // 元VB: Mcm1005uScreen_Load
    //   遷移元からDeliveryで tkKeiyakuId / torihikisakiNk / nonyusakiNk を受取
    // ===================================================================

    @GetMapping
    public String index(@RequestParam(value = "tkKeiyakuId", required = false) BigDecimal tkKeiyakuId,
                        @RequestParam(value = "torihikisakiNk", required = false) String torihikisakiNk,
                        @RequestParam(value = "nonyusakiNk", required = false) String nonyusakiNk,
                        @RequestParam(value = "selectedKikanIndex", defaultValue = "-1") int selectedKikanIndex,
                        @RequestParam(value = "mcm1006uCompleted", defaultValue = "false") boolean mcm1006uCompleted,
                        @RequestParam(required = false) String returnTo, Model model, HttpSession session) {

        if (tkKeiyakuId == null && session.getAttribute("mcm1005u.keiyakuId") != null) {
            tkKeiyakuId = new BigDecimal(session.getAttribute("mcm1005u.keiyakuId").toString());
            returnTo = session.getAttribute("mcm1005u.returnTo") instanceof String target ? target : "/mcm1009u";
            Object mode = session.getAttribute("mcm1005u.seniMotoKbn");
            service.setEntryMode(session, tkKeiyakuId, getLoginUserId(), mode instanceof Number n ? n.intValue() : "SHONIN".equals(mode) ? 5 : "TENPO".equals(mode) ? 6 : 4);
            for (String name : java.util.Collections.list(session.getAttributeNames()))
                if (name.startsWith("mcm1005u.")) session.removeAttribute(name);
        }
        if (returnTo != null && !mcm1006uCompleted) session.removeAttribute(draftKey(tkKeiyakuId));
        if (session.getAttribute("MCM1005U_TENPO_ID") instanceof BigDecimal id && tkKeiyakuId != null && id.compareTo(tkKeiyakuId) == 0) {
            service.setEntryMode(session, tkKeiyakuId, getLoginUserId(), 6);
            session.removeAttribute(draftKey(tkKeiyakuId));
            session.removeAttribute("MCM1005U_TENPO_ID"); returnTo = "/mcm2001u";
        } else if ("/mcm1003u".equals(returnTo) || "/mcm3007u".equals(returnTo)) service.setEntryMode(session, tkKeiyakuId, getLoginUserId(), 4);
        if (mcm1006uCompleted) {
            Object value = session.getAttribute("MCM1006U_RESULT");
            if (value instanceof Mcm1006uForm result && canUpdate()
                    && ((tkKeiyakuId == null && result.getTkKeiyakuId() == null)
                        || (tkKeiyakuId != null && result.getTkKeiyakuId() != null && tkKeiyakuId.compareTo(result.getTkKeiyakuId()) == 0))) {
                var pending = service.previewSelection(result, getLoginUserId(), draft(session, tkKeiyakuId));
                session.setAttribute(draftKey(tkKeiyakuId), pending);
                session.removeAttribute("MCM1006U_RESULT");
                model.addAttribute("message", "機器選定を反映しました。登録ボタンで保存してください。");
            }
        }
        var pending = tkKeiyakuId == null && !mcm1006uCompleted ? null : draft(session, tkKeiyakuId);
        if (tkKeiyakuId == null && pending != null) tkKeiyakuId = BigDecimal.ZERO;
        Object activeContract = session.getAttribute("MCM1005U_ACTIVE_CONTRACT");
        if (returnTo != null || tkKeiyakuId == null || (activeContract instanceof BigDecimal id && id.compareTo(tkKeiyakuId) != 0))
            service.clearUnlock(session);
        session.setAttribute("MCM1005U_ACTIVE_CONTRACT", tkKeiyakuId);
        String returnKey = "MCM1005U_RETURN_" + tkKeiyakuId;
        if (java.util.Set.of("/mcm1003u", "/mcm1009u", "/mcm3007u", "/mcm2001u", "/mcm1010u", "/mcm1011u").contains(returnTo == null ? "" : returnTo))
            session.setAttribute(returnKey, returnTo);
        model.addAttribute("returnUrl", session.getAttribute(returnKey) == null ? "/menu" : session.getAttribute(returnKey));
        McmTkKeiyakuEntity keiyaku;
        List<McmTkKikanEntity> kikanList;

        if (tkKeiyakuId != null && tkKeiyakuId.signum() > 0) {
            // 更新モード: SearchUpdate()
            keiyaku = service.loadKeiyaku(tkKeiyakuId);
            if (keiyaku == null) throw new InputCheckException("対象の契約が見つかりません。検索画面から開き直してください。");
            kikanList = service.loadKikanList(tkKeiyakuId);
        } else {
            // 新規モード: SearchInsert()
            if (pending == null) throw new InputCheckException("検索画面から契約対象の機器を選定してください。");
            keiyaku = buildNewKeiyaku();
            keiyaku.setKeiyakuDt(pending.periods().get(0).getKaisiDt());
            keiyaku.setShokaiKeiyakuDt(pending.periods().get(0).getKaisiDt());
            var user = getLoginUserInfo(session);
            if (user != null) { keiyaku.setIraitantosya(user.getUserName()); keiyaku.setIraijigyosyoNk(user.getJigyosyoName()); }
            kikanList = pending.periods();
            tkKeiyakuId = keiyaku.getTkKeiyakuId();
        }

        if (pending != null) {
            kikanList = pending.periods();
            if (mcm1006uCompleted) for (int i=0; i<kikanList.size(); i++)
                if (kikanList.get(i).getTkKikanId().compareTo(pending.periodId()) == 0) selectedKikanIndex = i;
        }
        model.addAttribute("draftToken", pending == null ? "" : pending.token());
        model.addAttribute("selectionPending", pending != null);

        // 選択中の期間タブに紐づくサブタブデータ
        McmTkKikanEntity selectedKikan = null;
        List<McmTkTenkenEntity> tenkenList = Collections.emptyList();
        McmTkSiharaiEntity siharai = null;
        List<McmTkTenpuEntity> tenpuList = Collections.emptyList();

        if (selectedKikanIndex < 0) selectedKikanIndex = service.initialPeriodIndex(kikanList);
        selectedKikanIndex = Math.max(0, Math.min(selectedKikanIndex, kikanList.size() - 1));
        List<Map<String, Object>> periodData = new ArrayList<>();
        if (pending != null) for (var data : pending.data()) periodData.add(new java.util.HashMap<>(data));
        else for (McmTkKikanEntity period : kikanList) {
            Map<String, Object> data = new java.util.HashMap<>();
            data.put("tenken", service.loadTenkenList(period.getTkKikanId()));
            data.put("siharai", service.loadSiharai(period.getTkKikanId()));
            data.put("tenpu", service.loadTenpuList(period.getTkKikanId()));
            data.put("kiki", service.loadKikiList(period.getTkKikanId()));
            periodData.add(data);
        }
        if (!kikanList.isEmpty()) {
            selectedKikan = kikanList.get(selectedKikanIndex);
            tenkenList = (List<McmTkTenkenEntity>) periodData.get(selectedKikanIndex).get("tenken");
            siharai = (McmTkSiharaiEntity) periodData.get(selectedKikanIndex).get("siharai");
            tenpuList = (List<McmTkTenpuEntity>) periodData.get(selectedKikanIndex).get("tenpu");
        }
        Map<String, String> timeLabels = periodData.stream().anyMatch(data -> !((List<?>) data.get("kiki")).isEmpty())
                ? service.maintenanceTimeLabels() : Collections.emptyMap();
        for (int i=0; i<kikanList.size(); i++) {
            periodData.get(i).put("kiki", service.equipmentDisplay(
                    (List<Map<String, Object>>) periodData.get(i).get("kiki"), timeLabels));
            periodData.get(i).put("tabLabel", service.periodLabel(kikanList.get(i), (List<McmTkTenpuEntity>) periodData.get(i).get("tenpu")));
            periodData.get(i).put("changeStart", kikanList.get(i).getSyuryoDt() == null || kikanList.get(i).getSyuryoDt().getYear() >= 9999 ? "" : DATE_FMT.format(kikanList.get(i).getSyuryoDt().plusDays(1)));
            periodData.get(i).put("companies", service.maintenanceCompanies(kikanList.get(i).getTkKikanId(), pending));
            periodData.get(i).put("canEditCosts", canUpdate() && service.canEditPeriodCosts(keiyaku, kikanList.get(i)));
            periodData.get(i).put("canChangePeriod", canUpdate() && service.canChangePeriod(keiyaku, kikanList.get(i), kikanList));
            periodData.get(i).put("canEditPeriod", canUpdate() && service.canEditPeriodDetails(keiyaku, kikanList.get(i), kikanList));
        }
        List<McmTkTenkenEntity> auditRows = new ArrayList<>();
        for (var data : periodData) auditRows.addAll((List<McmTkTenkenEntity>) data.get("tenken"));
        model.addAttribute("audit", service.auditDisplay(keiyaku, auditRows));
        model.addAttribute("periodAudit", service.periodAuditDisplay(kikanList));
        model.addAttribute("periodData", periodData);
        if (!kikanList.isEmpty()) {
            if (torihikisakiNk == null) torihikisakiNk = kikanList.get(0).getTorihikisakiNk();
            if (nonyusakiNk == null) nonyusakiNk = kikanList.get(0).getNonyusakiNk();
        }
        model.addAttribute("keiyaku", keiyaku);
        model.addAttribute("kikanList", kikanList);
        model.addAttribute("selectedKikanIndex", selectedKikanIndex);
        model.addAttribute("selectedKikan", selectedKikan);
        model.addAttribute("tenkenList", tenkenList);
        model.addAttribute("siharai", siharai);
        model.addAttribute("tenpuList", tenpuList);
        model.addAttribute("torihikisakiNk", torihikisakiNk);
        model.addAttribute("nonyusakiNk", nonyusakiNk);
        model.addAttribute("shoninjotaiOptions", buildShoninjotaiOptions());
        // 【移植】元VB: UpdateButton/SinseiButton/KeiyakuHakoButton等.AuthorityIsThrough=False
        //   「取引先契約関連」権限が参照専用の場合、登録・申請・破棄・変更系ボタンを非活性にする。
        model.addAttribute("canUpdate", canUpdate() && service.canSave(keiyaku));
        model.addAttribute("canEditDetails", canUpdate() && service.canEditWithUnlock(keiyaku));
        model.addAttribute("canApply", canUpdate() && service.canEditDetails(keiyaku));
        model.addAttribute("canReleaseLock", canUpdate() && service.canReleaseLock(getLoginUserId()));
        model.addAttribute("lockReleased", service.isUnlocked(tkKeiyakuId));
        model.addAttribute("jotaiLabel", Map.of("0","依頼","1","見積","2","契約","3","破棄","4","解約","9","作成中").getOrDefault(keiyaku.getJotai(), ""));
        model.addAttribute("approvalLabel", Map.of("0","作成中","1","審査中","2","承認中","3","承認済","4","差戻中").getOrDefault(keiyaku.getShoninjotai(), ""));

        return "mcm1005u/index";
    }

    // ===================================================================
    // MCM1006U 起動（変更ボタン）
    // 元VB: Mcm1005uTabControl.vb HenkoButton_Click → RaiseEvent TabHenkoButtonClick
    // ===================================================================

    @PostMapping("/launchMcm1006u")
    public String launchMcm1006u(
            @RequestParam BigDecimal tkKeiyakuId,
            @RequestParam(required = false) BigDecimal tkKikanId,
            @RequestParam(required = false) String tekiyobi,
            @RequestParam(required = false) String torihikisakiNk,
            @RequestParam(required = false) String nonyusakiNk,
            @RequestParam(defaultValue = "0") int selectedKikanIndex,
            RedirectAttributes ra, HttpSession session) {

        String backUrl = backUrl(tkKeiyakuId, torihikisakiNk, nonyusakiNk, selectedKikanIndex);

        // 【移植】元VB: Mcm1005uTabControl.HenkoButton.AuthorityIsThrough=False（参照専用権限で非活性）。
        //   画面側のボタン非活性だけに依存せず、サーバー側でも実効権限を検証する。
        if (!canUpdate()) {
            ra.addFlashAttribute("error", "権限がないため実行できません。");
            return backUrl;
        }

        // 変更摘要日バリデーション
        if (tekiyobi == null || tekiyobi.isBlank()) {
            ra.addFlashAttribute("changeStartError", "変更適用開始日は必ず入力してください。");
            ra.addFlashAttribute("changeStartValue", tekiyobi == null ? "" : tekiyobi);
            return backUrl;
        }
        String kaisiDt;
        String syuryoDt;
        try {
            if (!tekiyobi.trim().matches("[0-9]{4}/[0-9]{2}/[0-9]{2}")) throw new java.time.DateTimeException("date format");
            LocalDate kaisi = LocalDate.parse(tekiyobi.trim(), DATE_FMT);
            if (kaisi.getYear() < 1) throw new java.time.DateTimeException("date year");
            kaisiDt = DATE_FMT.format(kaisi);
            syuryoDt = DATE_FMT.format(kaisi.plusYears(1).minusDays(1));
        } catch (java.time.DateTimeException e) {
            ra.addFlashAttribute("changeStartError", "変更適用開始日はyyyy/mm/ddの書式で入力してください。");
            ra.addFlashAttribute("changeStartValue", tekiyobi);
            return backUrl;
        }

        var pending = draft(session, tkKeiyakuId);
        if (pending != null) {
            if (tkKikanId == null || tkKikanId.compareTo(pending.periodId()) != 0) throw new InputCheckException("最新の期間から変更してください。");
            var delivery = org.springframework.util.SerializationUtils.clone(pending.selection());
            delivery.setRestorePendingSelection(true);
            for (var removed : delivery.getTankaDelRows()) {
                if (delivery.getTankaRows().stream().noneMatch(t -> java.util.Objects.equals(t.getTkTankaId(),removed.getTkTankaId()))) delivery.getTankaRows().add(removed);
            }
            delivery.setKaisiDt(kaisiDt); delivery.setSyuryoDt(syuryoDt);
            // 【VB準拠】閉じる時に「変更」を押したタブへ戻すため、戻り先を保持する
            session.setAttribute("MCM1006U_RETURN_ID", tkKeiyakuId);
            session.setAttribute("MCM1006U_RETURN_URL", backUrl.substring("redirect:".length()));
            session.setAttribute("MCM1006U_DELIVERY", delivery);
            return "redirect:/mcm1006u/step1";
        }
        // 対象期間のロード
        McmTkKikanEntity selectedKikan = null;
        if (tkKikanId != null) {
            List<McmTkKikanEntity> kikanList = service.loadKikanList(tkKeiyakuId);
            for (McmTkKikanEntity k : kikanList) {
                if (tkKikanId.compareTo(k.getTkKikanId()) == 0) {
                    selectedKikan = k;
                    break;
                }
            }
        }
        if (selectedKikan == null) {
            ra.addFlashAttribute("error", "対象期間が見つかりません。先に登録してください。");
            return backUrl;
        }

        McmTkKeiyakuEntity keiyaku = service.loadKeiyaku(tkKeiyakuId);
        if (!service.canChangePeriod(keiyaku, selectedKikan, service.loadKikanList(tkKeiyakuId))) throw new InputCheckException("現在の期間では変更できません。");
        service.validateChangeStart(selectedKikan, service.loadKikanList(tkKeiyakuId), LocalDate.parse(kaisiDt, DATE_FMT));
        // tkKikanId が存在する場合は TAB_HENKO(=2)、新規なら INSERT(=1)
        int seniMotoKbn = (tkKikanId != null) ? 2 : 1;

        Mcm1006uForm delivery = service.buildMcm1006uDelivery(
                selectedKikan, keiyaku, seniMotoKbn, kaisiDt, syuryoDt);
        // 【VB準拠】閉じる時に「変更」を押したタブへ戻すため、戻り先を保持する（VBは子画面を閉じるだけでタブ位置は維持）
        session.setAttribute("MCM1006U_RETURN_ID", tkKeiyakuId);
        session.setAttribute("MCM1006U_RETURN_URL", backUrl.substring("redirect:".length()));
        session.setAttribute("MCM1006U_DELIVERY", delivery);

        return "redirect:/mcm1006u/step1";
    }

    // ===================================================================
    // 期間タブ切替（AJAX）
    // 元VB: KikanTabControl_SelectedIndexChanged
    // ===================================================================

    @GetMapping("/loadKikan")
    @ResponseBody
    public Map<String, Object> loadKikan(@RequestParam BigDecimal tkKikanId) {
        List<McmTkTenkenEntity> tenkenList = service.loadTenkenList(tkKikanId);
        McmTkSiharaiEntity siharai = service.loadSiharai(tkKikanId);
        List<McmTkTenpuEntity> tenpuList = service.loadTenpuList(tkKikanId);
        return Map.of(
                "tenkenList", tenkenList,
                "siharai", siharai != null ? siharai : new McmTkSiharaiEntity(),
                "tenpuList", tenpuList
        );
    }

    // ===================================================================
    // 登録処理
    // 元VB: UpdateButton_Click → UpdButton()
    // ===================================================================

    @PostMapping("/save")
    public String save(
            // ── 契約ヘッダー ──
            @RequestParam(value = "tkKeiyakuId", required = false) BigDecimal tkKeiyakuId,
            @RequestParam(value = "keiyakuNo", required = false) String keiyakuNo,
            @RequestParam(value = "shokaiKeiyakuDt", required = false) String shokaiKeiyakuDt,
            @RequestParam(value = "keiyakuDt", required = false) String keiyakuDt,
            @RequestParam(value = "jidokosinFlg", defaultValue = "0") String jidokosinFlg,
            @RequestParam(value = "kaiyakuDt", required = false) String kaiyakuDt,
            @RequestParam(value = "keiyakumanryoDt", required = false) String keiyakumanryoDt,
            @RequestParam(value = "entyokeiyakumanryoDt", required = false) String entyokeiyakumanryoDt,
            @RequestParam(value = "packFlg", defaultValue = "0") String packFlg,
            @RequestParam(value = "packkeiyakunaiyo", required = false) String packkeiyakunaiyo,
            @RequestParam(value = "biko", required = false) String biko,
            @RequestParam(value = "kosinnaiyo", required = false) String kosinnaiyo,
            @RequestParam(value = "iraijigyosyoNk", required = false) String iraijigyosyoNk,
            @RequestParam(value = "iraitantosya", required = false) String iraitantosya,
            // ── 期間リスト ──
            @RequestParam(value = "kikanIds", required = false) List<BigDecimal> kikanIds,
            @RequestParam(value = "kaisiDts", required = false) List<String> kaisiDts,
            @RequestParam(value = "syuryoDts", required = false) List<String> syuryoDts,
            @RequestParam(value = "kikanBikos", required = false) List<String> kikanBikos,
            // ── 点検リスト ──
            @RequestParam(value = "tenkenIds", required = false) List<BigDecimal> tenkenIds,
            @RequestParam(value = "tenkenkaisus", required = false) List<BigDecimal> tenkenkaisus,
            @RequestParam(value = "tenkenBikos", required = false) List<String> tenkenBikos,
            // ── 支払情報 ──
            @RequestParam(value = "siharaiId", required = false) List<BigDecimal> siharaiId,
            @RequestParam(value = "siharaiKikanId", required = false) List<BigDecimal> siharaiKikanId,
            @RequestParam(value = "siharaiKaisu", required = false) List<BigDecimal> siharaiKaisu,
            @RequestParam(value = "siharaiTsuki", required = false) List<Integer> siharaiTsukiMonths,
            @RequestParam(value = "siharaiBiko", required = false) List<String> siharaiBiko,
            // ── 添付ファイル ──
            @RequestParam(value = "tenpuIds", required = false) List<BigDecimal> tenpuIds,
            @RequestParam(value = "tenpuKikanIds", required = false) List<BigDecimal> tenpuKikanIds,
            @RequestParam(value = "tenpuFileNks", required = false) List<String> tenpuFileNks,
            @RequestParam(value = "tenpuDirectories", required = false) List<String> tenpuDirectories,
            @RequestParam(value = "tenpuShoninjotais", required = false) List<String> tenpuShoninjotais,
            @RequestParam(value = "torihikisakiNk", required = false) String torihikisakiNk,
            @RequestParam(value = "nonyusakiNk", required = false) String nonyusakiNk,
            @RequestParam(value = "selectedKikanIndex", defaultValue = "0") int selectedKikanIndex,
            RedirectAttributes ra, HttpSession session, HttpServletRequest request) {

        request.setAttribute("mcmRecoveryUrl", backUrl(tkKeiyakuId, torihikisakiNk, nonyusakiNk, selectedKikanIndex).substring(9));
        request.setAttribute("mcmSubmittedValues", request.getParameterMap());
        // 【移植】元VB: UpdateButton.AuthorityIsThrough=False（参照専用権限で非活性）。
        //   画面側のボタン非活性だけに依存せず、サーバー側でも実効権限を検証する。
        if (!canUpdate()) {
            ra.addFlashAttribute("error", "権限がないため実行できません。");
            return backUrl(tkKeiyakuId, torihikisakiNk, nonyusakiNk, selectedKikanIndex);
        }

        var pending = draft(session, tkKeiyakuId);
        String token = request.getParameter("draftToken");
        if ((pending == null && token != null && !token.isBlank())
                || (pending != null && !pending.token().equals(token)))
            throw new com.daifuku.mcm.exception.ExclusiveControlException("他のユーザがデータを変更した可能性があります。処理をやり直してください。");
        String action = request.getParameter("operation");
        if ("upload".equals(action) || "deleteAttachment".equals(action)) {
            String period = request.getParameter("attachmentPeriodId");
            if (period == null || !period.matches("[0-9]+")) throw new InputCheckException("期間が選択されていません。");
            var periodId = new BigDecimal(period);
            if (pending != null) {
                var contract = service.loadKeiyaku(tkKeiyakuId);
                var selected = pending.periods().stream().filter(p -> p.getTkKikanId().compareTo(periodId)==0).findFirst().orElseThrow(() -> new InputCheckException("対象期間が見つかりません。"));
                if (!service.canEditPeriodDetails(contract, selected, pending.periods())) throw new InputCheckException("現在の期間では添付を変更できません。");
            }
            Runnable change = () -> {
                if ("upload".equals(action)) {
                    var multipart = org.springframework.web.util.WebUtils.getNativeRequest(request, org.springframework.web.multipart.MultipartHttpServletRequest.class);
                    attachments.add(tkKeiyakuId,periodId,multipart == null ? null : multipart.getFile("attachmentFile_"+period),getLoginUserId());
                } else {
                    BigDecimal attachmentId = new BigDecimal(request.getParameter("attachmentDeleteId"));
                    if (service.loadTenpuList(periodId).stream().noneMatch(t -> t.getTkTenpuId().compareTo(attachmentId)==0)) throw new InputCheckException("対象ファイルが見つかりません。");
                    attachments.delete(tkKeiyakuId,attachmentId,getLoginUserId());
                }
            };
            synchronized (session) {
                if (pending != null) {
                    if (draft(session,tkKeiyakuId) != pending) throw new InputCheckException("画面を開き直してください。");
                    session.setAttribute(draftKey(tkKeiyakuId),service.changeDraftAttachment(pending,getLoginUserId(),change));
                } else change.run();
            }
            ra.addFlashAttribute("submittedValues",request.getParameterMap());
            ra.addFlashAttribute("message","upload".equals(action) ? "ファイルを追加しました。" : "ファイルを削除しました。");
            return backUrl(tkKeiyakuId,torihikisakiNk,nonyusakiNk,selectedKikanIndex);
        }
        // ── 契約ヘッダー組み立て ──
        McmTkKeiyakuEntity keiyaku = new McmTkKeiyakuEntity();
        keiyaku.setTkKeiyakuId(tkKeiyakuId != null ? tkKeiyakuId : service.getNextKeiyakuId());
        keiyaku.setLastupdateDt(parseTimestamp(request.getParameter("keiyakuVersion")));
        keiyaku.setKeiyakuNo(keiyakuNo);
        keiyaku.setShokaiKeiyakuDt(parseDate(shokaiKeiyakuDt));
        keiyaku.setKeiyakuDt(parseDate(keiyakuDt));
        keiyaku.setJidokosinFlg(new BigDecimal(jidokosinFlg));
        keiyaku.setKaiyakuDt(parseDate(kaiyakuDt));
        keiyaku.setKeiyakumanryoDt(parseDate(keiyakumanryoDt));
        keiyaku.setEntyokeiyakumanryoDt(parseDate(entyokeiyakumanryoDt));
        keiyaku.setPackFlg(new BigDecimal(packFlg));
        keiyaku.setPackkeiyakunaiyo(packkeiyakunaiyo);
        keiyaku.setBiko(biko);
        keiyaku.setKosinnaiyo(kosinnaiyo);
        keiyaku.setIraijigyosyoNk(iraijigyosyoNk);
        keiyaku.setIraitantosya(iraitantosya);
        keiyaku.setJotai(McmConstants.JOTAI_KEIYAKU);
        keiyaku.setShoninjotai(McmConstants.SHONINJOTAI_SAKUSEICHU_CD);

        // ── 期間リスト組み立て ──
        List<McmTkKikanEntity> kikanList = new ArrayList<>();
        if (kikanIds != null) {
            for (int i = 0; i < kikanIds.size(); i++) {
                McmTkKikanEntity k = new McmTkKikanEntity();
                k.setTkKikanId(kikanIds.get(i));
                k.setTkKeiyakuId(keiyaku.getTkKeiyakuId());
                k.setLastupdateDt(parseTimestamp(request.getParameter("kikanVersion_" + k.getTkKikanId())));
                k.setKeiyakujikantai(getAt(request.getParameterValues("keiyakujikantai") == null ? null : java.util.Arrays.asList(request.getParameterValues("keiyakujikantai")), i));
                k.setKaisiDt(parseDate(getAt(kaisiDts, i)));
                k.setSyuryoDt(parseDate(getAt(syuryoDts, i)));
                k.setBiko(getAt(kikanBikos, i));
                String[] periodContractNos = request.getParameterValues("kikanKeiyakuNos");
                if (periodContractNos != null && i < periodContractNos.length) k.setKeiyakuNo(periodContractNos[i]);
                String[] methods = request.getParameterValues("hosyuhohos");
                if (methods != null && i < methods.length) k.setHosyuhoho(methods[i]);
                String[] standard = request.getParameterValues("hyojungokeiKins"), subtotal = request.getParameterValues("sikirisyokeiKins"), discounts = request.getParameterValues("syusseinebikiKins"), totals = request.getParameterValues("sikirigokeiKins");
                if (service.isUnlocked(tkKeiyakuId)) {
                    if (standard != null && i < standard.length) k.setHyojungokeiKin(parseAmount(standard[i]));
                    if (subtotal != null && i < subtotal.length) k.setSikirisyokeiKin(parseAmount(subtotal[i]));
                    if (discounts != null && i < discounts.length) k.setSyusseinebikiKin(parseAmount(discounts[i]));
                    if (totals != null && i < totals.length) k.setSikirigokeiKin(parseAmount(totals[i]));
                }
                k.setYukoFlg(new BigDecimal(McmConstants.YUKO_FLG_YUKO));
                kikanList.add(k);
            }
        }

        // ── 点検リスト組み立て ──
        List<McmTkTenkenEntity> tenkenList = new ArrayList<>();
        if (tenkenIds != null) {
            for (int i = 0; i < tenkenIds.size(); i++) {
                McmTkTenkenEntity t = new McmTkTenkenEntity();
                t.setTkTenkenId(tenkenIds.get(i));
                String company = request.getParameter("hoshugaisha_" + t.getTkTenkenId());
                String nightCompany = request.getParameter("yakanhoshugaisha_" + t.getTkTenkenId());
                t.setHoshugaishaId(company == null ? null : company.isBlank() ? BigDecimal.valueOf(-1) : new BigDecimal(company));
                t.setYakanhoshugaishaId(nightCompany == null ? null : nightCompany.isBlank() ? BigDecimal.valueOf(-1) : new BigDecimal(nightCompany));
                t.setTenkenkaisu(getAtBd(tenkenkaisus, i));
                t.setBiko(getAt(tenkenBikos, i));
                t.setLastupdateDt(parseTimestamp(request.getParameter("tenkenVersion_" + t.getTkTenkenId())));
                String[] weekdays = request.getParameterValues("tenkenkanoyobis");
                t.setTenkenkanoyobi(weekdays != null && i < weekdays.length ? weekdays[i] : null);
                t.setYakantaioumu(request.getParameter("yakantaioumu_" + t.getTkTenkenId()) != null ? "1" : "0");
                for (int month = 0; month < 12; month++)
                    t.setTsukiAt(month, request.getParameter("tenkenTsuki_" + t.getTkTenkenId() + "_" + month) != null ? "1" : "0");
                tenkenList.add(t);
            }
        }

        // ── 支払情報組み立て ──
        List<McmTkSiharaiEntity> siharaiList = new ArrayList<>();
        List<Boolean[]> siharaiTsukiList = new ArrayList<>();
        if (siharaiKikanId != null) {
            for (int i = 0; i < siharaiKikanId.size(); i++) {
                McmTkSiharaiEntity payment = new McmTkSiharaiEntity();
                payment.setTkSiharaiId(getAtBd(siharaiId, i));
                payment.setTkKikanId(siharaiKikanId.get(i));
                payment.setKaisu(getAtBd(siharaiKaisu, i));
                payment.setBiko(getAt(siharaiBiko, i));
                payment.setLastupdateDt(parseTimestamp(request.getParameter("siharaiVersion_" + payment.getTkSiharaiId())));
                Boolean[] months = new Boolean[12];
                java.util.Arrays.fill(months, false);
                String[] checked = request.getParameterValues("siharaiTsuki_" + payment.getTkSiharaiId());
                if (checked != null) for (String value : checked) {
                    int month = Integer.parseInt(value);
                    if (month < 1 || month > 12) throw new InputCheckException("支払月が不正です。");
                    months[month - 1] = true;
                }
                siharaiList.add(payment);
                siharaiTsukiList.add(months);
            }
        }

        // ── 添付ファイルリスト組み立て ──
        List<McmTkTenpuEntity> tenpuList = new ArrayList<>();
        if (tenpuIds != null) {
            for (int i = 0; i < tenpuIds.size(); i++) {
                McmTkTenpuEntity tp = new McmTkTenpuEntity();
                tp.setTkTenpuId(tenpuIds.get(i));
                tp.setTkKikanId(getAtBd(tenpuKikanIds, i));
                tp.setTenpufileNk(getAt(tenpuFileNks, i));
                tp.setDirectory(getAt(tenpuDirectories, i));
                tp.setShoninjotai(getAt(tenpuShoninjotais, i));
                tenpuList.add(tp);
            }
        }

        if (pending != null) {
            if (selectedKikanIndex < 0 || selectedKikanIndex >= kikanList.size()) throw new InputCheckException("期間が選択されていません。");
            Mcm1005uService.SavedSelection saved;
            synchronized (session) {
                if (draft(session, tkKeiyakuId) != pending) throw new InputCheckException("画面の情報が更新されています。開き直してください。");
                saved = service.saveSelection(pending, keiyaku, kikanList, tenkenList, siharaiList, siharaiTsukiList,
                        kikanList.get(selectedKikanIndex).getTkKikanId(), request.getParameter("operation"), getLoginUserId());
                session.removeAttribute(draftKey(tkKeiyakuId));
            }
            var actualPeriods = service.loadKikanList(saved.contractId());
            // 保存に成功した今回の再表示だけ、仮IDと採番後IDの対応を渡す。
            ra.addFlashAttribute("uiIdMapping", Map.of("contract", tkKeiyakuId.stripTrailingZeros().toPlainString(),
                    "periods", saved.periodIds(), "inspections", saved.inspectionIds()));
            for (int i=0; i<actualPeriods.size(); i++) if (actualPeriods.get(i).getTkKikanId().compareTo(saved.periodId()) == 0) selectedKikanIndex = i;
            Object caller = session.getAttribute("MCM1005U_RETURN_" + tkKeiyakuId);
            if (caller != null) session.setAttribute("MCM1005U_RETURN_" + saved.contractId(), caller);
            ra.addFlashAttribute("message", "sinsei".equals(request.getParameter("operation")) ? "申請が完了しました。" : "登録が完了しました。");
            if ("report".equals(request.getParameter("operation"))) ra.addFlashAttribute("reportPeriodId", saved.periodId().toPlainString());
            return backUrl(saved.contractId(), torihikisakiNk, nonyusakiNk, selectedKikanIndex);
        }
        boolean report = "report".equals(request.getParameter("operation"));
        if (report) {
            if (selectedKikanIndex < 0 || selectedKikanIndex >= kikanList.size()) throw new InputCheckException("期間が選択されていません。");
            service.checkReport(tkKeiyakuId, kikanList.get(selectedKikanIndex).getTkKikanId(), keiyaku.getKaiyakuDt() != null, getLoginUserId());
        }
        if ("sinsei".equals(request.getParameter("operation"))) {
            if (selectedKikanIndex < 0 || selectedKikanIndex >= kikanList.size()) throw new InputCheckException("期間が選択されていません。");
            service.saveAndSinsei(keiyaku, kikanList, tenkenList, siharaiList, siharaiTsukiList, tenpuList,
                    kikanList.get(selectedKikanIndex).getTkKikanId(), getLoginUserId());
            ra.addFlashAttribute("message", "申請が完了しました。");
        } else {
            service.save(keiyaku, kikanList, tenkenList, siharaiList, siharaiTsukiList, tenpuList, getLoginUserId());
            ra.addFlashAttribute("message", "登録が完了しました。");
        }

        if (report) ra.addFlashAttribute("reportPeriodId", kikanList.get(selectedKikanIndex).getTkKikanId().toPlainString());

        return backUrl(keiyaku.getTkKeiyakuId(), torihikisakiNk, nonyusakiNk, selectedKikanIndex);
    }

    // ===================================================================
    // 申請処理
    // 元VB: SinseiButton_Click
    // ===================================================================

    @PostMapping("/sinsei")
    public String sinsei(@RequestParam BigDecimal tkKeiyakuId,
                         @RequestParam BigDecimal tkKikanId,
                         @RequestParam(required = false) String torihikisakiNk,
                         @RequestParam(required = false) String nonyusakiNk,
                         @RequestParam(defaultValue = "0") int selectedKikanIndex,
                         RedirectAttributes ra, HttpSession session) {
        String backUrl = backUrl(tkKeiyakuId, torihikisakiNk, nonyusakiNk, selectedKikanIndex);
        // 【移植】元VB: SinseiButton.AuthorityIsThrough=False（参照専用権限で非活性）。
        if (!canUpdate()) {
            ra.addFlashAttribute("error", "権限がないため実行できません。");
            return backUrl;
        }
        service.sinsei(tkKeiyakuId, tkKikanId, getLoginUserId());
        ra.addFlashAttribute("message", "申請を完了しました。");
        return backUrl;
    }

    // ===================================================================
    // 添付ファイル削除
    // 元VB: FileDelButton_Click
    // ===================================================================

    @PostMapping("/deleteTenpu")
    public String deleteTenpu(@RequestParam BigDecimal tkTenpuId,
                               @RequestParam BigDecimal tkKeiyakuId,
                               @RequestParam(required = false) String torihikisakiNk,
                               @RequestParam(required = false) String nonyusakiNk,
                               @RequestParam(defaultValue = "0") int selectedKikanIndex,
                               RedirectAttributes ra, HttpSession session) {
        String backUrl = backUrl(tkKeiyakuId, torihikisakiNk, nonyusakiNk, selectedKikanIndex);
        // 【移植】元VB: FileDelButton.AuthorityIsThrough=False（参照専用権限で非活性）。
        if (!canUpdate()) {
            ra.addFlashAttribute("error", "権限がないため実行できません。");
            return backUrl;
        }
        if (draft(session, tkKeiyakuId) != null) throw new InputCheckException("機器選定の変更を登録してからファイルを削除してください。");
        attachments.delete(tkKeiyakuId, tkTenpuId, getLoginUserId());
        return backUrl;
    }

    // ===================================================================
    // ユーティリティ
    // ===================================================================

    private McmTkKeiyakuEntity buildNewKeiyaku() {
        McmTkKeiyakuEntity e = new McmTkKeiyakuEntity();
        e.setTkKeiyakuId(BigDecimal.ZERO);
        e.setJotai(McmConstants.JOTAI_KEIYAKU);
        e.setShoninjotai(McmConstants.SHONINJOTAI_SAKUSEICHU_CD);
        e.setPackFlg(BigDecimal.ZERO);
        e.setJidokosinFlg(BigDecimal.ONE);
        return e;
    }

    private List<Map<String, String>> buildShoninjotaiOptions() {
        return List.of(
                Map.of("key", McmConstants.SHONINJOTAI_SAKUSEICHU_CD,  "value", McmConstants.SHONINJOTAI_SAKUSEICHU_NK),
                Map.of("key", McmConstants.SHONINJOTAI_SHINSACHU_CD,   "value", McmConstants.SHONINJOTAI_SHINSACHU_NK),
                Map.of("key", McmConstants.SHONINJOTAI_SHONINCHU_CD,   "value", McmConstants.SHONINJOTAI_SHONINCHU_NK),
                Map.of("key", McmConstants.SHONINJOTAI_SHONINZUMI_CD,  "value", McmConstants.SHONINJOTAI_SHONINZUMI_NK),
                Map.of("key", McmConstants.SHONINJOTAI_SASHIMODOSHI_CD,"value", McmConstants.SHONINJOTAI_SASHIMODOSHI_NK)
        );
    }

    /** VBのC0表示（円記号・3桁区切り）と、編集中の数値の両方を受け付ける。 */
    private BigDecimal parseAmount(String value) {
        String number = value == null ? "" : value.trim().replaceFirst("^[¥￥]", "");
        if (!number.matches("-?(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+)(?:\\.[0-9]+)?"))
            throw new InputCheckException("金額は数値で入力してください。");
        return new BigDecimal(number.replace(",", ""));
    }

    private java.time.LocalDateTime parseDate(String s) {
        if (s == null || s.isBlank()) return null;
        try {
            return LocalDate.parse(s.trim().replace('-', '/'), DATE_FMT).atStartOfDay();
        } catch (java.time.format.DateTimeParseException e) {
            throw new InputCheckException("日付が正しくありません。実在する日付をyyyy/MM/dd形式で入力してください。");
        }
    }

    private java.time.LocalDateTime parseTimestamp(String value) {
        if (value == null || value.isBlank()) return null;
        try { return java.time.LocalDateTime.parse(value); }
        catch (java.time.format.DateTimeParseException ex) { throw new InputCheckException("更新情報が不正です。画面を開き直してください。"); }
    }

    private String backUrl(BigDecimal id, String supplier, String customer, int index) {
        return "redirect:" + UriComponentsBuilder.fromPath("/mcm1005u")
                .queryParam("tkKeiyakuId", id).queryParam("torihikisakiNk", "{supplier}")
                .queryParam("nonyusakiNk", "{customer}").queryParam("selectedKikanIndex", index)
                .encode().buildAndExpand(Map.of("supplier", supplier == null ? "" : supplier, "customer", customer == null ? "" : customer)).toUriString();
    }

    private String getAt(List<String> list, int i) {
        return (list != null && i < list.size()) ? list.get(i) : null;
    }

    private BigDecimal getAtBd(List<BigDecimal> list, int i) {
        return (list != null && i < list.size()) ? list.get(i) : null;
    }

    @Override
    protected String getScreenTitle() { return "取引先契約内容"; }

    @Override
    protected String getFunctionId() { return "MCM1005U"; }
}