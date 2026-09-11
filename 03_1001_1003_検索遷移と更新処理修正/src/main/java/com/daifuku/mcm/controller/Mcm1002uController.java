package com.daifuku.mcm.controller;

/**
 * 【変換元】Mcm1002u1Screen.vb + Mcm1002u2Screen.vb
 *   MCM1002U 需要家見積作成 Controller
 */

import com.daifuku.mcm.common.Mcm1002uConstants;
import com.daifuku.mcm.form.Mcm1002u1Form;
import com.daifuku.mcm.form.Mcm1002u2Form;
import com.daifuku.mcm.form.Mcm1002uDeliveryDto;
import com.daifuku.mcm.service.Mcm1002uService;
import jakarta.servlet.http.HttpSession;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.util.List;

@Controller
@RequestMapping("/mcm1002u")
@RequiredArgsConstructor
public class Mcm1002uController {

    private final Mcm1002uService service;

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

        Mcm1002u1Form form = new Mcm1002u1Form();

        // セッションからDelivery情報を取得（MCM1003U/MCM1008Uから遷移時）
        Mcm1002uDeliveryDto delivery = (Mcm1002uDeliveryDto) session.getAttribute("mcm1002uDelivery");

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
        }

        model.addAttribute("form", form);
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
    public String step2(Model model, HttpSession session) {

        Mcm1002u1Form screen1Form = (Mcm1002u1Form) session.getAttribute("mcm1002u1Form");
        Mcm1002uDeliveryDto delivery = (Mcm1002uDeliveryDto) session.getAttribute("mcm1002uDelivery");

        Mcm1002u2Form form = service.buildScreen2Form(screen1Form, delivery);

        model.addAttribute("form", form);
        model.addAttribute("delivery", delivery);
        // 保守方法・点検可能曜日のドロップダウン用
        model.addAttribute("hosyuhohoOptions", service.getHosyuhohoOptions());
        model.addAttribute("tenkenkanoyobiOptions", service.getTenkenkanoyobiOptions());
        return "mcm1002u/step2";
    }

    // ========================================================
    // 登録ボタン
    // 【変換元】Mcm1002u2Screen.UpdateButton_Click → updateMethod()
    // ========================================================
    @PostMapping("/save")
    public String save(@ModelAttribute Mcm1002u2Form form,
                       HttpSession session, RedirectAttributes ra) {

        Mcm1002uDeliveryDto delivery = (Mcm1002uDeliveryDto) session.getAttribute("mcm1002uDelivery");

        // バリデーション
        List<String> errors = service.validate8H24H(form);
        if (!errors.isEmpty()) {
            ra.addFlashAttribute("errors", errors);
            return "redirect:/mcm1002u/step2";
        }

        try {
            // TODO: ログインユーザー名を取得
            String loginUser = "system";
            service.saveAll(form, delivery, loginUser, false);
            ra.addFlashAttribute("message", "登録を完了しました。");
        } catch (Exception e) {
            ra.addFlashAttribute("errors", List.of("登録に失敗しました: " + e.getMessage()));
        }

        return "redirect:/mcm1002u/step2";
    }

    // ========================================================
    // 依頼書発行ボタン
    // 【変換元】Mcm1002u2Screen.IraisyoHakkoButton_Click
    // ========================================================
    @PostMapping("/iraisho")
    public String iraisho(@ModelAttribute Mcm1002u2Form form,
                          HttpSession session, RedirectAttributes ra) {

        Mcm1002uDeliveryDto delivery = (Mcm1002uDeliveryDto) session.getAttribute("mcm1002uDelivery");
        String loginUser = "system";

        // バリデーション
        List<String> errors = service.validate8H24H(form);
        if (!errors.isEmpty()) {
            ra.addFlashAttribute("errors", errors);
            return "redirect:/mcm1002u/step2";
        }

        try {
            // 登録 + 状態変更
            service.saveAll(form, delivery, loginUser, true);

            // TODO: Excel出力（Apache POI）
            // service.generateExcelReport(delivery.getTmIraiId());

            ra.addFlashAttribute("message", "依頼書を発行しました");
        } catch (Exception e) {
            ra.addFlashAttribute("errors", List.of("依頼書発行に失敗しました: " + e.getMessage()));
        }

        return "redirect:/mcm1002u/step2";
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
}