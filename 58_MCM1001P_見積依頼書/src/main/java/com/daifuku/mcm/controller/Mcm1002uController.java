package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.List;
import java.util.Map;

import jakarta.servlet.http.HttpSession;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

/**
 * 【変換元】Mcm1002u1Screen.vb + Mcm1002u2Screen.vb
 *   MCM1002U 需要家見積作成 Controller
 */

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.common.Mcm1002uConstants;
import com.daifuku.mcm.form.Mcm1002u1Form;
import com.daifuku.mcm.form.Mcm1002u2Form;
import com.daifuku.mcm.form.Mcm1002uDeliveryDto;
import com.daifuku.mcm.service.Mcm1002uService;

@Controller
@RequestMapping("/mcm1002u")
public class Mcm1002uController extends BaseController {

    private static final Logger log = LoggerFactory.getLogger(Mcm1002uController.class);

    private final Mcm1002uService service;

    public Mcm1002uController(Mcm1002uService service) {
        this.service = service;
    }

    /**
     * 【移植】元VB: CPCoreUserControl.GetAuthorityDivision() 相当。
     *   「取引先見積関連」権限が「作成」でなければ登録・依頼書発行を許可しない。
     */
    private boolean canUpdate() {
        return hasUpdateAuthority() && service.canUpdate(getLoginUserId());
    }

    // ========================================================
    // Screen1: 機器構成選定（初期表示）
    // 【変換元】Mcm1002u1Screen_Load
    //
    //   呼出し元パターン（3種）:
    //     1) MCM1001U（新規） → URLクエリパラメータで plantId 等を直接渡す
    //     2) MCM1003U（修正/複製・プラント1件） → session["mcm1002uDelivery"] にセット済み
    //     3) MCM1008U（複製・プラント複数選択後） → session["mcm1002uDelivery"] にセット済み
    // ========================================================
    @GetMapping
    public String index(
            @RequestParam(required = false) Long plantId,
            @RequestParam(required = false) Long tmKeiyakujikanId,
            @RequestParam(name = "mitsumori", required = false) Integer mitsumoriParam,
            @RequestParam(required = false) String supportId,
            @RequestParam(required = false) String plantNk,
            @RequestParam(required = false) String nonyusakiCd,
            @RequestParam(required = false) String nonyusakiNk,
            @RequestParam(required = false) Boolean copyFlg,
            @RequestParam(name = "nav", required = false) String nav,
            Model model, HttpSession session) {

        Mcm1002uDeliveryDto delivery = null;

        if (plantId == null) {
            Object flashNav = model.asMap().get("mcm1002u.nav");
            Object flashDelivery = model.asMap().get("mcm1002u.navDelivery");

            boolean regularNavigation =
                    nav != null
                    && nav.equals(flashNav)
                    && flashDelivery instanceof Mcm1002uDeliveryDto;

            if (!regularNavigation) {
                // URL直接アクセス。前回のセッション情報は表示に使用しない。
                // 別タブの作業を壊さないよう、ここではセッションを削除しない。
                Mcm1002u1Form initialForm = new Mcm1002u1Form();
                initialForm.setReturnUrl("/mcm1001u");
                model.addAttribute("form", initialForm);
                model.addAttribute("canUpdate", false);
                return "mcm1002u/index";
            }

            delivery = (Mcm1002uDeliveryDto) flashDelivery;
            normalizeFukuseiJotai(delivery);
            session.setAttribute("mcm1002uDelivery", delivery);

            Object flashForm = model.asMap().get("mcm1002u.navScreen1Form");
            if (flashForm instanceof Mcm1002u1Form savedForm) {
                // Screen2の「前へ」。選択済みのチェック状態を復元する。
                session.setAttribute("mcm1002u1Form", savedForm);
                model.addAttribute("form", savedForm);
                model.addAttribute("canUpdate", canUpdate());
                return "mcm1002u/index";
            }

            // MCM1003U/MCM1008Uからの新しい遷移。
            session.removeAttribute("mcm1002u1Form");
            session.removeAttribute("mcm1002u2Form");
        }

        Mcm1002u1Form form = new Mcm1002u1Form();

        if (plantId != null) {
            // MCM1001UからのplantId付き遷移は従来どおり扱う。
            delivery = new Mcm1002uDeliveryDto();
            delivery.setMitsumoriFlg(
                    mitsumoriParam != null
                            ? mitsumoriParam
                            : Mcm1002uConstants.SHINKI);
            delivery.setPlantId(java.math.BigDecimal.valueOf(plantId));
            delivery.setPlantNk(plantNk);
            delivery.setNonyusakiCd(nonyusakiCd);
            delivery.setNonyusakiNk(nonyusakiNk);
            delivery.setSupportId(supportId);
            delivery.setTmKeiyakujikanId(tmKeiyakujikanId);
            delivery.setCopyFlg(copyFlg != null && copyFlg);
            delivery.setReturnUrl("/mcm1001u");
            normalizeFukuseiJotai(delivery);

            session.setAttribute("mcm1002uDelivery", delivery);
            session.removeAttribute("mcm1002u1Form");
            session.removeAttribute("mcm1002u2Form");
        }

        if (delivery != null) {
            form.setPlantId(delivery.getPlantId());
            form.setTmKeiyakujikanId(delivery.getTmKeiyakujikanId());
            form.setNonyusakiCd(delivery.getNonyusakiCd());
            form.setNonyusakiNk(delivery.getNonyusakiNk());
            form.setSupportId(delivery.getSupportId());
            form.setPlantNk(delivery.getPlantNk());
            form.setJotai(delivery.getJotai());
            form.setMitsumoriFlg(delivery.getMitsumoriFlg());
            form.setReturnUrl(delivery.getReturnUrl() != null ? delivery.getReturnUrl() : "/mcm1003u");
        } else {
            form.setReturnUrl("/mcm1003u");
        }

        /*
         * 【不一致修正】元VB: Mcm1002u1Screen_Load() - MCM_TM_MITSUMORITableAdapter.Fill()
         *   plantIdをキーにMCM_MA_PLANT/MCM_MA_NONYUSAKIマスタを結合検索し、
         *   NONYUSAKI_ID・PLANT_ID・SUPPORT_ID等のヘッダ情報を取得する。
         *   この値は最終的にMCM_TM_MITSUMORI新規行の必須項目としてsaveAll()で保存される。
         *   従来はこの検索が行われておらず、nonyusakiId等がform上に存在しないまま
         *   登録され、MCM_TM_MITSUMORI.PLANT_ID等がNULLになる不具合の原因になっていた。
         */
        if (form.getPlantId() != null) {
            Map<String, Object> header = service.findMitsumoriHeader(form.getPlantId());
            if (header != null) {
                form.setNonyusakiId(toBigDecimal(header.get("NONYUSAKI_ID")));
                form.setNonyusakijusyo1Nk(toStr(header.get("JUSYO1_NK")));
                form.setNonyusakijusyo2Nk(toStr(header.get("JUSYO2_NK")));
                form.setNonyubusyoNk(toStr(header.get("NONYUBUSYO_NK")));
                form.setNonyutantosyaNk(toStr(header.get("NONYUTANTOSYA_NK")));
                form.setNonyutelNo(toStr(header.get("NONYUTEL_NO")));
                form.setNonyufaxNo(toStr(header.get("NONYUFAX_NO")));
                // 【変換元】NONYUSAKI_CD/NONYUSAKI_NK/SUPPORT_ID/PLANT_NKはDeliveryから
                //   受け取れなかった場合(セッション復元等)のフォールバックとしてもマスタ値で補完する。
                if (form.getNonyusakiCd() == null) form.setNonyusakiCd(toStr(header.get("NONYUSAKI_CD")));
                if (form.getNonyusakiNk() == null) form.setNonyusakiNk(toStr(header.get("NONYUSAKI_NK")));
                if (form.getSupportId() == null) form.setSupportId(toStr(header.get("SUPPORT_ID")));
                if (form.getPlantNk() == null) form.setPlantNk(toStr(header.get("PLANT_NK")));
            }
        }

        // 【変換元】Mcm1002u1Screen_Load - 修正時のReadOnly制御
        //   状態が「作成中」「依頼」以外の場合、契約有無チェック列を編集不可にする
        boolean readOnly = false;
        if (form.getMitsumoriFlg() == Mcm1002uConstants.SYUSEI) {
            String jotai = form.getJotai();
            readOnly = jotai != null
                    && !Mcm1002uConstants.JOTAI_SAKUSEICHU_NK.equals(jotai)
                    && !Mcm1002uConstants.JOTAI_IRAI_NK.equals(jotai);
        }
        form.setReadOnly(readOnly);

        // 機器構成パターン取得（チェックフラグ付き）
        Long resolvedPlantId = form.getPlantId() != null ? form.getPlantId().longValue() : null;
        if (resolvedPlantId != null) {
            Long keiyakuId = form.getTmKeiyakujikanId() != null ? form.getTmKeiyakujikanId() : 0L;
            form.setKikikoseiRows(service.loadKikikoseiWithCheckFlags(resolvedPlantId, keiyakuId));
            form.setKikimeisaiRows(service.loadKikimeisai(resolvedPlantId));
            form.setKotaimeisaiRows(service.loadKotaimeisai(resolvedPlantId, keiyakuId));

            // 【要望対応】初期表示時は全グリッドのチェックボックスをONにする
            form.getKikikoseiRows().forEach(r -> r.setChecked(true));
            form.getKikimeisaiRows().forEach(r -> r.setChecked(true));
            form.getKotaimeisaiRows().forEach(r -> r.setChecked(true));
        }

        model.addAttribute("form", form);
        // 【移植】元VB: NextButton.AuthorityIsThrough=True だが後続のUpdateButton等が
        //   AuthorityIsThrough=False のため、参照専用ユーザーは次画面で操作できない
        //   ことを明示する目的でここでも権限フラグを渡す。
        model.addAttribute("canUpdate", canUpdate());
        return "mcm1002u/index";
    }

    // ========================================================
    // Screen1 → Screen2 遷移（次へボタン）
    // 【変換元】Mcm1002u1Screen → Mcm1002u2Screen遷移
    // ========================================================
    @PostMapping("/next")
    public String next(@ModelAttribute Mcm1002u1Form form,
                       HttpSession session, RedirectAttributes ra) {

        if (form.getPlantId() == null) {
            ra.addFlashAttribute(
                    "errors",
                    List.of("プラントが選択されていません。見積作成検索画面から選択してください。"));
            return "redirect:/mcm1002u";
        }

        session.setAttribute("mcm1002u1Form", form);
        return redirectToStep2(ra);
    }

    /**
     * 【#471不一致修正】元VB: Mcm1002u1Screen_Load()
     *   If IsNull(jotai) Or Me.mitsumoriFlg = Mcm1002uConstant.FUKUSEI Then jotai = 作成中
     *   複製で作成する見積は新規見積のため、複製元の状態（依頼/見積/契約等）を引き継がず
     *   「作成中」とする。これにより、複製元の状態に関係なく契約有無チェック・登録が可能となり、
     *   依頼書発行時の状態遷移（作成中→依頼）も正しく行われる。
     *   ※ 既存フローへの影響を避けるため、本対応では複製時のみ適用する
     *     （VBのIsNull(jotai)分岐＝通常新規時の補完は未適用。報告事項参照）。
     */
    private static void normalizeFukuseiJotai(Mcm1002uDeliveryDto delivery) {
        if (delivery != null && delivery.getMitsumoriFlg() == Mcm1002uConstants.FUKUSEI) {
            delivery.setJotai(Mcm1002uConstants.JOTAI_SAKUSEICHU_NK);
        }
    }

    /** Screen2への正規遷移を示すFlash属性名（#374） */
    private static final String FLASH_STEP2_NAV = "mcm1002u.step2Nav";

    /**
     * 【#374不一致修正】Screen2へ正規遷移する。
     *   Screen1（index）と同様に、使い捨てのnavトークンをURLパラメータとFlash属性の
     *   両方に設定する。Flash属性は1回のリダイレクトでのみ有効なため、コピーした
     *   Screen2のURLを別タブに貼り付けた直接アクセスとは区別できる。
     */
    private String redirectToStep2(RedirectAttributes ra) {
        String nav = java.util.UUID.randomUUID().toString();
        ra.addAttribute("nav", nav);
        ra.addFlashAttribute(FLASH_STEP2_NAV, nav);
        return "redirect:/mcm1002u/step2";
    }

    // ========================================================
    // Screen2: 見積依頼機器選定（表示）
    // 【変換元】Mcm1002u2Screen_Load
    // ========================================================
    @GetMapping("/step2")
    public String step2(@ModelAttribute("form") Mcm1002u2Form flashForm,
                        @RequestParam(name = "nav", required = false) String nav,
                        Model model, HttpSession session) {

        /*
         * 【#374不一致修正】登録後にコピーしたScreen2のURLを別タブへ貼り付けると、
         *   セッションに残ったDelivery（登録後にtmKeiyakujikanIdが設定済み）から
         *   前回登録したデータが表示されていた。
         *   Screen1（index）と同様に、正規遷移（navトークン一致）かつセッションに
         *   Delivery情報がある場合のみセッション情報を表示に使用する。
         *   それ以外（URL直接アクセス）は空の画面を参照専用で表示する。
         *   別タブの作業を壊さないよう、ここではセッションを削除しない。
         */
        Object sessionDelivery = session.getAttribute("mcm1002uDelivery");
        boolean regularNavigation =
                nav != null
                && nav.equals(model.asMap().get(FLASH_STEP2_NAV))
                && sessionDelivery instanceof Mcm1002uDeliveryDto;

        if (!regularNavigation) {
            model.addAttribute("form", new Mcm1002u2Form());
            model.addAttribute("hosyuhohoOptions", service.getHosyuhohoOptions());
            model.addAttribute("tenkenkanoyobiOptions", service.getTenkenkanoyobiOptions());
            model.addAttribute("canUpdate", false);
            model.addAttribute("directAccess", true);
            model.addAttribute("jikantaiLocked", false);
            return "mcm1002u/step2";
        }

        Mcm1002uDeliveryDto delivery = (Mcm1002uDeliveryDto) sessionDelivery;

        /*
         * 【#373不一致修正】登録／依頼書発行の失敗時（SQLエラー等）に、DBから
         *   作り直したFormを表示すると、ユーザーが変更した入力内容（回答希望日・
         *   8H/24H・保守方法等）が失われてしまう。エラー時はController側で
         *   RedirectAttributes.addFlashAttribute("form", form) によりFlash経由で
         *   同じ属性名("form")の入力済みFormを引き継いでいるため、Spring MVCが
         *   自動的に@ModelAttribute("form")としてバインドしたFlash復元後のFormが
         *   ここに渡ってくる。その場合はDB再検索を行わず、そのまま画面へ戻す。
         */
        boolean restoredFromFlash = flashForm != null && !flashForm.getMitsumoriRows().isEmpty();
        Mcm1002u2Form form;
        if (restoredFromFlash) {
            form = flashForm;
        } else {
            Mcm1002u1Form screen1Form = (Mcm1002u1Form) session.getAttribute("mcm1002u1Form");
            form = service.buildScreen2Form(screen1Form, delivery);
        }

        model.addAttribute("form", form);
        model.addAttribute("delivery", delivery);
        // 保守方法・点検可能曜日のドロップダウン用
        model.addAttribute("hosyuhohoOptions", service.getHosyuhohoOptions());
        model.addAttribute("tenkenkanoyobiOptions", service.getTenkenkanoyobiOptions());
        // 【移植】元VB: UpdateButton/IraisyoHakkoButton.AuthorityIsThrough=False
        //   「取引先見積関連」権限が参照専用の場合、登録・依頼書発行ボタンを非活性にする。
        model.addAttribute("canUpdate", canUpdate());
        model.addAttribute("directAccess", false);
        // 【#471】元VB: Mcm1002u2Screen_Load() If Me.copyFlg Then 8H/24H列.ReadOnly = True
        //   見積複製＋「期間情報を複写する」の場合のみ8H/24Hを編集不可にする（登録後も継続）。
        model.addAttribute("jikantaiLocked", delivery.isCopyFlg());
        return "mcm1002u/step2";
    }

    // ========================================================
    // 登録ボタン
    // 【変換元】Mcm1002u2Screen.UpdateButton_Click → updateMethod()
    // ========================================================
    @PostMapping("/save")
    public String save(@ModelAttribute Mcm1002u2Form form,
                       HttpSession session, RedirectAttributes ra) {

        // 【移植】元VB: UpdateButton.AuthorityIsThrough=False（参照専用権限で非活性）。
        //   画面側のボタン非活性だけに依存せず、サーバー側でも実効権限を検証する。
        if (!canUpdate()) {
            ra.addFlashAttribute("errors", List.of("権限がないため実行できません。"));
            return redirectToStep2(ra);
        }

        Mcm1002uDeliveryDto delivery = (Mcm1002uDeliveryDto) session.getAttribute("mcm1002uDelivery");
        Mcm1002u1Form screen1Form = (Mcm1002u1Form) session.getAttribute("mcm1002u1Form");

        // 【#471】期間情報複写ありの場合、画面非活性だけに依存せず8H/24Hを表示時の値へ戻す
        service.restoreLockedJikantai(form, screen1Form, delivery);

        // バリデーション
        List<String> errors = service.validate8H24H(form);
        if (!errors.isEmpty()) {
            ra.addFlashAttribute("errors", errors);
            // 【#373不一致修正】バリデーションエラー時も入力内容を保持する。
            ra.addFlashAttribute("form", form);
            return redirectToStep2(ra);
        }

        try {
            String loginUser = resolveLoginUserName(session);
            Long keiyakujikanId = service.saveAll(form, screen1Form, delivery, loginUser, false);
            // 【変換元】Mcm1002u2Screen.vb - afterUpdate()
            //   登録後はmitsumoriFlgを「修正」に変更し、以後は登録済みの契約時間ID
            //   から画面を再表示する（Screen1のフォールバック情報には頼らない）。
            updateDeliveryAfterSave(session, delivery, keiyakujikanId);
            ra.addFlashAttribute("message", "登録を完了しました。");
        } catch (Exception e) {
            log.error("MCM1002U 登録処理に失敗しました", e);
            ra.addFlashAttribute("errors", List.of("登録に失敗しました: " + e.getMessage()));
            // 【#373不一致修正】登録失敗時（SQLエラー等）も入力内容を保持する。
            ra.addFlashAttribute("form", form);
        }

        return redirectToStep2(ra);
    }

    /**
     * 登録後、delivery情報を更新してセッションに保存し直す。
     * 【変換元】Mcm1002u2Screen.vb - afterUpdate()
     *   mitsumoriFlg = Mcm1002uConstant.SYUSEI （見積フラグを「修正」とする）
     */
    private void updateDeliveryAfterSave(HttpSession session, Mcm1002uDeliveryDto delivery, Long keiyakujikanId) {
        if (delivery == null || keiyakujikanId == null) {
            return;
        }
        delivery.setTmKeiyakujikanId(keiyakujikanId);
        delivery.setMitsumoriFlg(Mcm1002uConstants.SYUSEI);
        session.setAttribute("mcm1002uDelivery", delivery);
    }

    // ========================================================
    // 依頼書発行ボタン
    // 【変換元】Mcm1002u2Screen.IraisyoHakkoButton_Click
    // ========================================================
    @PostMapping("/iraisho")
    public String iraisho(@ModelAttribute Mcm1002u2Form form,
                          HttpSession session, RedirectAttributes ra) {

        Mcm1002uDeliveryDto delivery = (Mcm1002uDeliveryDto) session.getAttribute("mcm1002uDelivery");
        Mcm1002u1Form screen1Form = (Mcm1002u1Form) session.getAttribute("mcm1002u1Form");
        if (delivery == null || screen1Form == null) {
            ra.addFlashAttribute("error", "見積作成画面から再度操作してください。");
            return "redirect:/mcm1002u";
        }
        try {
            service.restoreReportDisplay(form, screen1Form, delivery);
            String state = delivery.getJotai();
            boolean editableState = state == null
                    || Mcm1002uConstants.JOTAI_SAKUSEICHU_NK.equals(state)
                    || Mcm1002uConstants.JOTAI_IRAI_NK.equals(state);
            boolean saveBeforePrint = canUpdate() && editableState;
            if (form.getMitsumoriRows().isEmpty()) {
                throw new IllegalArgumentException("依頼書の出力対象がありません。");
            }
            if (!saveBeforePrint) service.restoreSavedReportRows(form);
            // VBと同様、登録する場合だけ編集項目を検査する。
            // 参照／確定済みの再発行は、画面の未保存値ではなく登録済みデータを出力する。
            if (saveBeforePrint) {
                service.restoreLockedJikantai(form, screen1Form, delivery);
                List<String> errors = service.validate8H24H(form);
                if (!errors.isEmpty()) {
                    ra.addFlashAttribute("errors", errors);
                    ra.addFlashAttribute("form", form);
                    return redirectToStep2(ra);
                }
                if (state == null) delivery.setJotai(Mcm1002uConstants.JOTAI_SAKUSEICHU_NK);
                Long id = service.saveAll(form, screen1Form, delivery, resolveLoginUserName(session), true);
                updateDeliveryAfterSave(session, delivery, id);
            }
            List<Long> ids = form.getMitsumoriRows().stream()
                    .map(row -> row.getTmIraiId()).filter(id -> id != null && id > 0).distinct().toList();
            var deliveries = service.buildIraisyoDeliveryList(ids);
            String token = com.daifuku.mcm.dto.Mcm1001pExportState.remember(session, deliveries);
            ra.addFlashAttribute("excelDownloadUrl", "/mcm1001p/download?state=" + token);
            // 新規採番後も全取引先と入力内容を引き継ぐ。先頭1件へ作り直さない。
            ra.addFlashAttribute("form", form);
        } catch (Exception e) {
            log.error("MCM1002U 依頼書発行に失敗しました", e);
            ra.addFlashAttribute("errors", List.of("依頼書発行に失敗しました。出力対象と入力内容を確認してください。"));
            ra.addFlashAttribute("form", form);
        }
        return redirectToStep2(ra);
    }

    /**
     * ログインユーザーIDを登録者・更新者として使用する。
     * 【変換元】TODO: ログインユーザー名を取得 のプレースホルダを解消。
     */
    private String resolveLoginUserName(HttpSession session) {
        String userId = getLoginUserId();
        return userId != null ? userId : "system";
    }

    // ========================================================
    // 前へボタン
    // 【変換元】Mcm1002u2Screen.MaeButton_Click
    // ========================================================
    @GetMapping("/back")
    public String back(HttpSession session, RedirectAttributes ra) {
        Object delivery = session.getAttribute("mcm1002uDelivery");
        Object screen1Form = session.getAttribute("mcm1002u1Form");

        if (!(delivery instanceof Mcm1002uDeliveryDto)
                || !(screen1Form instanceof Mcm1002u1Form)) {
            return "redirect:/mcm1002u";
        }

        String nav = java.util.UUID.randomUUID().toString();
        ra.addAttribute("nav", nav);
        ra.addFlashAttribute("mcm1002u.nav", nav);
        ra.addFlashAttribute("mcm1002u.navDelivery", delivery);
        ra.addFlashAttribute("mcm1002u.navScreen1Form", screen1Form);

        return "redirect:/mcm1002u";
    }

    // ========================================================
    // 見積作成検索画面へ
    // 【変換元】TorihikisakiMitsumoriSakuseiKensakuGamenButton_Click
    //   元コード: '' 取引先見積作成検索画面へ遷移する
    //             ForwardScreen(McmScreenIdConstant.MCM1001U)
    // ========================================================
    @GetMapping("/search")
    public String toSearch() {
        return "redirect:/mcm1001u";
    }

    private static BigDecimal toBigDecimal(Object v) {
        if (v == null) return null;
        if (v instanceof BigDecimal) return (BigDecimal) v;
        return new BigDecimal(v.toString());
    }

    private static String toStr(Object v) {
        return v == null ? null : v.toString();
    }

    @Override
    protected String getScreenTitle() {
        return "需要家見積作成";
    }

    @Override
    protected String getFunctionId() {
        return "MCM1002U";
    }
}
