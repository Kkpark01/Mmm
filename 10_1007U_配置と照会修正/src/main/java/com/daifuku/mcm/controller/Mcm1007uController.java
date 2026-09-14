package com.daifuku.mcm.controller;

import com.daifuku.mcm.service.Mcm1007uService;
import com.daifuku.mcm.service.Mcm1007uService.ScreenData;
import jakarta.servlet.http.HttpSession;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import java.math.BigDecimal;

/**
 * 【変換元】Mcm1007uScreen.vb - MCM1007U 単価精査 コントローラ
 * 【変換元行数】約70行（Screen.vb）
 * 【作成者】T.Kajimura (2009/02/10)
 *
 * 元コード: WinForms の Mcm1007uScreen_Load() イベント内で
 *           Delivery からパラメータ取得 → Fill() → 表示 を実行。
 * 変換後:   HttpSession からパラメータ取得 → Service 呼出 → Model に格納 → Thymeleaf 表示。
 *
 * セッションキー（呼出元画面がセット済み）:
 *   mcm1007u.atsukaikikiId   ← 取扱機器ID
 *   mcm1007u.torihikisakiId  ← 取引先ID
 *   mcm1007u.keiyakujikantai ← 契約時間帯
 */
@Controller
@RequestMapping("/mcm1007u")
public class Mcm1007uController {

    private static final Logger log = LoggerFactory.getLogger(Mcm1007uController.class);

    private final Mcm1007uService service;

    public Mcm1007uController(Mcm1007uService service) {
        this.service = service;
    }

    /**
     * 画面表示（データ読込）。
     *
     * 【変換元】Mcm1007uScreen.vb - Mcm1007uScreen_Load()
     *   元コード:
     *     Dim delivery As Mcm1007uDelivery = MyBase.GetDeliveryData()
     *     torihikisakiId = delivery.TorihikisakiID
     *     atsukaiId = delivery.AtsukaikikiID
     *     keiyakujikantai = delivery.Keiyakujikantai
     *     Me.MCM_MA_ATSUKAIKIKITableAdapter.Fill(...)
     *     Me.MCM_TM_MITSUMORITableAdapter.Fill(...)
     *
     * @param session HTTPセッション
     * @param model   モデル
     * @return テンプレートパス
     */
    @GetMapping
    public String index(HttpSession session, Model model) {
        model.addAttribute("header", null);
        model.addAttribute("keiyakujikantai", "");
        model.addAttribute("list", java.util.List.of());
        model.addAttribute("resultCount", 0);

        /*
         * 【変換元】Mcm1007uDelivery.vb
         *   元コード:
         *     Private _AtsukaikikiID As Integer
         *     Private _TorihikisakiID As Integer
         *     Private _Keiyakujikantai As Integer
         */
        BigDecimal atsukaikikiId = toBigDecimal(session.getAttribute("mcm1007u.atsukaikikiId"));
        BigDecimal torihikisakiId = toBigDecimal(session.getAttribute("mcm1007u.torihikisakiId"));
        BigDecimal keiyakujikantai = toBigDecimal(session.getAttribute("mcm1007u.keiyakujikantai"));

        if (atsukaikikiId == null || torihikisakiId == null || keiyakujikantai == null) {
            log.warn("MCM1007U: セッションにパラメータが不足しています");
            model.addAttribute("errors", java.util.List.of("パラメータが不足しています。前の画面からやり直してください。"));
            return "mcm1007u/index";
        }

        ScreenData data;
        try {
            data = service.loadScreenData(torihikisakiId, atsukaikikiId, keiyakujikantai);
        } catch (org.springframework.dao.DataAccessException e) {
            log.error("MCM1007U: 単価精査の読込失敗", e);
            model.addAttribute("errors", java.util.List.of("単価情報を読み込めませんでした。前の画面から照会し直してください。"));
            return "mcm1007u/index";
        }
        // VB clears all five header fields as well as the grid when no rows match.
        if (data.isEmpty()) {
            model.addAttribute("errors", java.util.List.of("該当するデータがありません。"));
            return "mcm1007u/index";
        }

        if (data.getHeader() != null) {
            model.addAttribute("header", data.getHeader());
        }
        model.addAttribute("keiyakujikantai", keiyakujikantai);
        model.addAttribute("list", data.getRows());
        model.addAttribute("resultCount", data.getRows().size());

        /*
         * 【変換元】Mcm1007uScreen.vb
         *   元コード: If MCM_TM_MITSUMORI.Count = 0 Then DisplayMessage(MSG_0002)
         */

        return "mcm1007u/index";
    }

    /**
     * セッション属性をBigDecimalに変換する。
     */
    private BigDecimal toBigDecimal(Object value) {
        if (value == null) return null;
        if (value instanceof BigDecimal) return (BigDecimal) value;
        try {
            return new BigDecimal(value.toString());
        } catch (NumberFormatException e) {
            return null;
        }
    }
}
