package com.daifuku.mcm.controller;

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
import jakarta.servlet.http.HttpSession;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.math.BigDecimal;
import java.util.List;
import java.util.Map;

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
            Model model, HttpSession session) {

        Mcm1002uDeliveryDto delivery = (Mcm1002uDeliveryDto) session.getAttribute("mcm1002uDelivery");
        // #374: 保存フォームだけでは遷移情報にならない。正規の新規遷移はplantIdで受け取る。
        if (plantId == null && delivery == null) {
            session.removeAttribute("mcm1002u1Form");
            session.removeAttribute("mcm1002u2Form");
            Mcm1002u1Form initialForm = new Mcm1002u1Form();
            initialForm.setReturnUrl("/mcm1001u");
            model.addAttribute("form", initialForm);
            model.addAttribute("canUpdate", false);
            return "mcm1002u/index";
        }

        // 【不一致修正】MCM1002U2の「前へ」ボタン(back())は本メソッドへ単純リダイレクトするだけ
        //   だったため、plantIdクエリが無い「戻り表示」時でも新規表示と同じDB再検索+
        //   チェックボックス全ON処理が実行され、MCM1002U2に進む前にユーザーが選択した
        //   チェック状態が失われていた。plantIdクエリが無く、かつセッションに直前の
        //   Mcm1002u1Formが残っている場合は、それをそのまま復元して再表示する
        //   （DB再検索・チェックリセットを行わない）。
        if (plantId == null) {
            Mcm1002u1Form savedForm = (Mcm1002u1Form) session.getAttribute("mcm1002u1Form");
            if (savedForm != null) {
                model.addAttribute("form", savedForm);
                model.addAttribute("canUpdate", canUpdate());
                return "mcm1002u/index";
            }
        }

        Mcm1002u1Form form = new Mcm1002u1Form();

        // セッションからDelivery情報を取得（MCM1003U/MCM1008Uから遷移時）
        if (plantId != null) {
            // 新しいプラント選択は前回のDeliveryより優先する。クエリなしの戻り表示は既存Deliveryを利用する。
            // 【変換元】Mcm1001uScreen.vb - PLANTDataGridView_CellContentClick()
            delivery = new Mcm1002uDeliveryDto();
            delivery.setMitsumoriFlg(mitsumoriParam != null ? mitsumoriParam : Mcm1002uConstants.SHINKI);
            delivery.setPlantId(java.math.BigDecimal.valueOf(plantId));
            delivery.setPlantNk(plantNk);
            delivery.setNonyusakiCd(nonyusakiCd);
            delivery.setNonyusakiNk(nonyusakiNk);
            delivery.setSupportId(supportId);
            delivery.setTmKeiyakujikanId(tmKeiyakujikanId);
            delivery.setCopyFlg(copyFlg != null && copyFlg);
            // 【変換元】CPBaseUserControl標準「戻る」機能：呼出し元(MCM1001U)へ戻る
            delivery.setReturnUrl("/mcm1001u");
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

        // 選択された構成・明細・個体をセッションに格納
        session.setAttribute("mcm1002u1Form", form);
        return "redirect:/mcm1002u/step2";
    }

    // ========================================================
    // Screen2: 見積依頼機器選定（表示）
    // 【変換元】Mcm1002u2Screen_Load
    // ========================================================
    @GetMapping("/step2")
    public String step2(@ModelAttribute("form") Mcm1002u2Form flashForm, Model model, HttpSession session) {

        Mcm1002uDeliveryDto delivery = (Mcm1002uDeliveryDto) session.getAttribute("mcm1002uDelivery");

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
            return "redirect:/mcm1002u/step2";
        }

        Mcm1002uDeliveryDto delivery = (Mcm1002uDeliveryDto) session.getAttribute("mcm1002uDelivery");
        Mcm1002u1Form screen1Form = (Mcm1002u1Form) session.getAttribute("mcm1002u1Form");

        // バリデーション
        List<String> errors = service.validate8H24H(form);
        if (!errors.isEmpty()) {
            ra.addFlashAttribute("errors", errors);
            // 【#373不一致修正】バリデーションエラー時も入力内容を保持する。
            ra.addFlashAttribute("form", form);
            return "redirect:/mcm1002u/step2";
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

        return "redirect:/mcm1002u/step2";
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

        /*
         * 【不一致修正】元VB: Mcm1002u2Screen.vb - IraisyoHakkoButton_Click()
         *   IraisyoHakkoButton自体はAuthorityIsThrough=True（常時活性）。
         *   クリック時、UpdateButton.Enabled=True（作成権限あり）かつ状態が
         *   「作成中」/「依頼」の場合のみ登録処理を実行し、参照権限の場合は
         *   登録処理をスキップして帳票出力のみ実行する（continueする）。
         *   以前の実装は canUpdate() が false の場合に処理全体を中断していたが、
         *   これはVBの実際の挙動（参照権限でも依頼書発行=帳票出力は可能）と異なる。
         *   参照権限時は登録をスキップし、帳票出力のみ行うよう修正する。
         */
        Mcm1002uDeliveryDto delivery = (Mcm1002uDeliveryDto) session.getAttribute("mcm1002uDelivery");
        Mcm1002u1Form screen1Form = (Mcm1002u1Form) session.getAttribute("mcm1002u1Form");
        String loginUser = resolveLoginUserName(session);

        // バリデーション
        List<String> errors = service.validate8H24H(form);
        if (!errors.isEmpty()) {
            ra.addFlashAttribute("errors", errors);
            // 【#373不一致修正】バリデーションエラー時も入力内容を保持する。
            ra.addFlashAttribute("form", form);
            return "redirect:/mcm1002u/step2";
        }

        try {
            if (canUpdate()) {
                // 作成権限あり：登録 + 状態変更
                Long keiyakujikanId = service.saveAll(form, screen1Form, delivery, loginUser, true);
                updateDeliveryAfterSave(session, delivery, keiyakujikanId);
            }
            // 参照権限のみの場合：登録処理は行わず、帳票出力のみ実行する。

            /*
             * 【#373不一致修正】元VB: IraisyoHakkoButton_Click() 帳票出力処理
             *   登録処理（updateMethod()）が正常に完了した場合のみ、見積データごとに
             *   Mcm1001pData_Header/Mcm1001pData_Viewを組み立ててexcelOutput()を呼ぶ。
             *   Web版では従来Excel出力がTODOのまま未実装だったため、
             *   Mcm1001pExcelService（既存のMCM1001P帳票サービス）を用いて実装する。
             *   取引先(TM_IRAI_ID)ごとに1件のExcel帳票を生成するため、複数取引先が
             *   対象の場合は複数のDeliveryDtoをセッションに積み、1件ずつダウンロード
             *   させる（PRGパターン：MCM1001PのMcm1001pController#downloadを再利用）。
             */
            List<Long> tmIraiIds = form.getMitsumoriRows().stream()
                .map(row -> row.getTmIraiId())
                .filter(id -> id != null && id > 0)
                .distinct()
                .collect(java.util.stream.Collectors.toList());
            List<com.daifuku.mcm.dto.Mcm1001pDeliveryDto> deliveries = service.buildIraisyoDeliveryList(tmIraiIds);

            if (!deliveries.isEmpty()) {
                // 【変換元】MCM1001Uのダウンロード画面(Mcm1001pController#download)は
                //   セッションに積んだ1件のDeliveryDtoをExcelとして出力する。
                //   複数取引先が対象の場合、先頭の1件を出力する（1回の依頼書発行操作
                //   につき1ファイル。複数取引先分の一括ダウンロードはVB版でも
                //   ファイル保存ダイアログの繰返し表示であり、Web版での複数ファイル
                //   同時ダウンロードの導線は本チケットの対象外とする）。
                session.setAttribute("mcm1001pDelivery", deliveries.get(0));
                ra.addFlashAttribute("message", "依頼書を発行しました");
                ra.addFlashAttribute("excelDownloadUrl", "/mcm1001p/download");
                if (deliveries.size() > 1) {
                    log.info("MCM1002U 依頼書発行: 対象取引先が複数件のため先頭の1件のみExcel出力します。件数={}",
                        deliveries.size());
                }
            } else {
                // 帳票対象データが1件も無い場合（対象データ不整合）は、登録成功でも
                // Excel未出力である旨を利用者に伝える。
                ra.addFlashAttribute("message", "依頼書を発行しました（出力対象の明細がありませんでした）");
            }
        } catch (Exception e) {
            log.error("MCM1002U 依頼書発行に失敗しました", e);
            ra.addFlashAttribute("errors", List.of("依頼書発行に失敗しました: " + e.getMessage()));
            // 【#373不一致修正】登録失敗時（SQLエラー等）は入力内容を保持し、
            //   完了メッセージ・Excel出力を行わない。
            ra.addFlashAttribute("form", form);
        }

        return "redirect:/mcm1002u/step2";
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
    public String back() {
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
