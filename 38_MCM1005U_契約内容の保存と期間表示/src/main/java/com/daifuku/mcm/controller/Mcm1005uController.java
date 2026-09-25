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

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.entity.*;
import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.form.Mcm1006uForm;
import com.daifuku.mcm.service.Mcm1005uService;
import jakarta.servlet.http.HttpSession;
import jakarta.servlet.http.HttpServletRequest;
import org.springframework.web.util.UriComponentsBuilder;
import com.daifuku.mcm.exception.InputCheckException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

@Controller
@RequestMapping("/mcm1005u")
public class Mcm1005uController extends BaseController {
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
        if (!"POST".equals(request.getMethod()) || !request.getRequestURI().endsWith("/mcm1005u/save")) return;
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
                        @RequestParam(value = "selectedKikanIndex", defaultValue = "0") int selectedKikanIndex,
                        @RequestParam(value = "mcm1006uCompleted", defaultValue = "false") boolean mcm1006uCompleted,
                        @RequestParam(required = false) String returnTo, Model model, HttpSession session) {

        if (tkKeiyakuId == null && session.getAttribute("mcm1005u.keiyakuId") != null) {
            tkKeiyakuId = new BigDecimal(session.getAttribute("mcm1005u.keiyakuId").toString());
            returnTo = "/mcm1009u";
            for (String name : java.util.Collections.list(session.getAttributeNames()))
                if (name.startsWith("mcm1005u.")) session.removeAttribute(name);
        }
        String returnKey = "MCM1005U_RETURN_" + tkKeiyakuId;
        if (java.util.Set.of("/mcm1003u", "/mcm1009u", "/mcm3007u", "/mcm2001u", "/mcm1010u").contains(returnTo == null ? "" : returnTo))
            session.setAttribute(returnKey, returnTo);
        model.addAttribute("returnUrl", session.getAttribute(returnKey) == null ? "/menu" : session.getAttribute(returnKey));
        McmTkKeiyakuEntity keiyaku;
        List<McmTkKikanEntity> kikanList;

        if (tkKeiyakuId != null) {
            // 更新モード: SearchUpdate()
            keiyaku = service.loadKeiyaku(tkKeiyakuId);
            if (keiyaku == null) throw new InputCheckException("対象の契約が見つかりません。検索画面から開き直してください。");
            kikanList = service.loadKikanList(tkKeiyakuId);
        } else {
            // 新規モード: SearchInsert()
            keiyaku = buildNewKeiyaku();
            kikanList = Collections.emptyList();
            tkKeiyakuId = keiyaku.getTkKeiyakuId();
        }

        // MCM1006U 選定結果の適用（変換元: TabHenkoButtonClick → UpdateTKKikiData）
        if (mcm1006uCompleted) {
            Mcm1006uForm result = (Mcm1006uForm) session.getAttribute("MCM1006U_RESULT");
            if (result != null && canUpdate() && service.canEditDetails(keiyaku) && result.getTkKeiyakuId() != null && result.getTkKikanId() != null
                    && result.getTkKeiyakuId().compareTo(tkKeiyakuId) == 0
                    && kikanList.stream().anyMatch(k -> k.getTkKikanId().compareTo(result.getTkKikanId()) == 0)) {
                try {
                    service.applyMcm1006uResult(result, getLoginUserId());
                    for (int i = 0; i < kikanList.size(); i++) {
                        if (kikanList.get(i).getTkKikanId().compareTo(result.getTkKikanId()) == 0) selectedKikanIndex = i;
                    }
                    session.removeAttribute("MCM1006U_RESULT");
                    model.addAttribute("message", "機器情報を更新しました。");
                } catch (Exception ex) {
                    model.addAttribute("error", "機器情報の更新に失敗しました: " + ex.getMessage());
                }
            }
        }

        // 選択中の期間タブに紐づくサブタブデータ
        McmTkKikanEntity selectedKikan = null;
        List<McmTkTenkenEntity> tenkenList = Collections.emptyList();
        McmTkSiharaiEntity siharai = null;
        List<McmTkTenpuEntity> tenpuList = Collections.emptyList();

        selectedKikanIndex = Math.max(0, Math.min(selectedKikanIndex, kikanList.size() - 1));
        if (!kikanList.isEmpty()) {
            selectedKikan = kikanList.get(selectedKikanIndex);
            BigDecimal kikanId = selectedKikan.getTkKikanId();
            tenkenList = service.loadTenkenList(kikanId);
            siharai = service.loadSiharai(kikanId);
            tenpuList = service.loadTenpuList(kikanId);
        }

        List<Map<String, Object>> periodData = new ArrayList<>();
        for (McmTkKikanEntity period : kikanList) {
            Map<String, Object> data = new java.util.HashMap<>();
            data.put("tenken", service.loadTenkenList(period.getTkKikanId()));
            data.put("siharai", service.loadSiharai(period.getTkKikanId()));
            data.put("tenpu", service.loadTenpuList(period.getTkKikanId()));
            data.put("kiki", service.loadKikiList(period.getTkKikanId()));
            periodData.add(data);
        }
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
        model.addAttribute("canUpdate", canUpdate());
        model.addAttribute("canEditDetails", canUpdate() && service.canEditDetails(keiyaku));
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
            ra.addFlashAttribute("error", "変更摘要日を入力してください。");
            return backUrl;
        }
        String kaisiDt;
        String syuryoDt;
        try {
            LocalDate kaisi = LocalDate.parse(tekiyobi.trim(), DATE_FMT);
            kaisiDt = DATE_FMT.format(kaisi);
            syuryoDt = DATE_FMT.format(kaisi.plusYears(1).minusDays(1));
        } catch (Exception e) {
            ra.addFlashAttribute("error", "変更摘要日の形式が正しくありません（yyyy/MM/dd）。");
            return backUrl;
        }

        // 対象期間のロード
        McmTkKikanEntity selectedKikan = null;
        if (tkKikanId != null) {
            List<McmTkKikanEntity> kikanList = service.loadKikanList(tkKeiyakuId);
            for (McmTkKikanEntity k : kikanList) {
                if (tkKikanId.equals(k.getTkKikanId())) {
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
        if (!service.canEditDetails(keiyaku)) throw new InputCheckException("現在の状態では変更できません。");
        // tkKikanId が存在する場合は TAB_HENKO(=2)、新規なら INSERT(=1)
        int seniMotoKbn = (tkKikanId != null) ? 2 : 1;

        Mcm1006uForm delivery = service.buildMcm1006uDelivery(
                selectedKikan, keiyaku, seniMotoKbn, kaisiDt, syuryoDt);
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

        if ("upload".equals(request.getParameter("operation"))) {
            String period = request.getParameter("attachmentPeriodId");
            if (period == null || !period.matches("[0-9]+")) throw new InputCheckException("期間が選択されていません。");
            var multipart = org.springframework.web.util.WebUtils.getNativeRequest(request, org.springframework.web.multipart.MultipartHttpServletRequest.class);
            attachments.add(tkKeiyakuId, new BigDecimal(period), multipart == null ? null : multipart.getFile("attachmentFile_" + period), getLoginUserId());
            ra.addFlashAttribute("submittedValues", request.getParameterMap());
            ra.addFlashAttribute("message", "ファイルを追加しました。");
            return backUrl(tkKeiyakuId, torihikisakiNk, nonyusakiNk, selectedKikanIndex);
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
        try {
            service.sinsei(tkKeiyakuId, tkKikanId, getLoginUserId());
            ra.addFlashAttribute("message", "申請を完了しました。");
        } catch (McmBusinessException ex) {
            ra.addFlashAttribute("error", ex.getMessage());
        }
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
        service.deleteTenpu(tkTenpuId, tkKeiyakuId, getLoginUserId());
        return backUrl;
    }

    // ===================================================================
    // ユーティリティ
    // ===================================================================

    private McmTkKeiyakuEntity buildNewKeiyaku() {
        McmTkKeiyakuEntity e = new McmTkKeiyakuEntity();
        e.setTkKeiyakuId(service.getNextKeiyakuId());
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
