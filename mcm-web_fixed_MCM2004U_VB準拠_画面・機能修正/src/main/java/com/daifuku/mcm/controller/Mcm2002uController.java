package com.daifuku.mcm.controller;

import com.daifuku.mcm.constants.Mcm2002uConstants;
import com.daifuku.mcm.dto.Mcm2002uDeliveryDto;
import com.daifuku.mcm.form.Mcm2002uForm;
import com.daifuku.mcm.service.Mcm2002uService;
import com.daifuku.mcm.service.Mcm2002uService.BuildResult;
import com.daifuku.mcm.service.Mcm2002uService.SearchResult;
import jakarta.servlet.http.HttpSession;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.util.List;

/**
 * 【変換元】Mcm2002uScreen.vb（92,685 bytes / 22メソッド）
 * MCM2002U（カスタマー見積機器選定）Controller
 *
 * <p>画面遷移パターン:</p>
 * <ul>
 *   <li>MCM2001U → (session: mcm2002u.delivery) → MCM2002U</li>
 *   <li>MCM2002U → (session: mcm2003u.buildResult) → MCM2003U</li>
 *   <li>MCM2002U → (戻る) → MCM2001U</li>
 * </ul>
 *
 * <p>VB.NET のイベントハンドラとの対応:</p>
 * <ul>
 *   <li>Mcm2002uScreen_Load → GET /mcm2002u</li>
 *   <li>KIHONSETTEIButton_Click → POST /mcm2002u/kihon-settei</li>
 *   <li>KENSAKUButton_Click → GET /mcm2002u/back</li>
 * </ul>
 *
 * @author MCM Migration Tool
 */
@Controller
@RequestMapping("/mcm2002u")
public class Mcm2002uController {

    @org.springframework.web.bind.annotation.ExceptionHandler(org.springframework.dao.DataAccessException.class)
    public String dataAccessFailure(org.springframework.dao.DataAccessException e, RedirectAttributes ra) {
        log.error("MCM2002U data access failed",e);
        ra.addFlashAttribute("errors",List.of("画面の情報を取得できませんでした。再検索してから、もう一度操作してください。"));
        return "redirect:/mcm2004u";
    }

    private static final Logger log = LoggerFactory.getLogger(Mcm2002uController.class);

    /** セッションキー: MCM2001Uからの受渡しデータ */
    private static final String SESSION_DELIVERY = "mcm2002u.delivery";
    /** セッションキー: 検索結果（画面状態保持） */
    private static final String SESSION_SEARCH_RESULT = "mcm2002u.searchResult";
    /** セッションキー: MCM2003Uへの受渡しデータ */
    private static final String SESSION_MCM2003U_BUILD = "mcm2003u.buildResult";
    /** セッションキー: MCM2003Uへの受渡し（ヘッダー情報） */
    private static final String SESSION_MCM2003U_DELIVERY = "mcm2003u.delivery";

    private final Mcm2002uService service;

    public Mcm2002uController(Mcm2002uService service) {
        this.service = service;
    }

    // ================================================================
    // 1. 画面ロード (GET)
    // 【変換元】Mcm2002uScreen.vb - Mcm2002uScreen_Load()
    //   元コード:
    //     Dim delivery As Mcm2002uDelivery = MyBase.GetDeliveryData
    //     Me.plantId = delivery.PlantId
    //     ...
    //     If delivery.SeniMotoKbn = SENIMOTO_KBN_KENSAKU Then
    //       SearchInsert()
    //     Else
    //       SearchUpdate()
    //     End If
    // ================================================================
    @GetMapping
    public String index(HttpSession session, Model model) {

        // --- セッションからMCM2001Uの受渡しデータを取得 ---
        Mcm2002uDeliveryDto delivery =
            (Mcm2002uDeliveryDto) session.getAttribute(SESSION_DELIVERY);

        if (delivery == null) {
            log.warn("MCM2002U: セッションに受渡しデータがありません。MCM2001Uへリダイレクトします。");
            return "redirect:/mcm2001u";
        }

        // --- Service呼出し: 検索実行 ---
        SearchResult searchResult = service.loadScreen(delivery);

        // --- セッションに検索結果を保持 ---
        session.setAttribute(SESSION_SEARCH_RESULT, searchResult);

        // --- Formに初期値を設定 ---
        Mcm2002uForm form = new Mcm2002uForm();
        form.setNonyusakiCd(delivery.getNonyusakiCd());
        form.setNonyusakiNk(delivery.getNonyusakiNk());
        form.setSupportId(delivery.getSupportId());
        form.setPlantNk(delivery.getPlantNk());
        form.setSeniMotoKbn(searchResult.getSeniMotoKbn());

        // 既存開始日がある場合はフォームに設定
        if (searchResult.getOldKaisiDt() != null) {
            form.setKaisiDt(searchResult.getOldKaisiDt());
        }

        // グリッドデータをフォームに設定
        form.setKoseiRows(searchResult.getKoseiRows());
        form.setMeisaiRows(searchResult.getMeisaiRows());
        form.setKotaiRows(searchResult.getKotaiRows());
        form.setTankaRows(searchResult.getTankaRows());

        // --- Modelに設定 ---
        model.addAttribute("form", form);
        model.addAttribute("searchResult", searchResult);

        /**
         * 【変換元】Mcm2002uScreen_Load() - 表示制御
         *   元コード:
         *     Me.KAISI_DTTextBox.Visible = True
         *     Me.MCM_MA_KIKIKOSEIDataGridView.ReadOnly = True
         */
        model.addAttribute("showKaisiDt", searchResult.isShowKaisiDt());
        model.addAttribute("showKeiyakuJikantai", searchResult.isShowKeiyakuJikantai());
        model.addAttribute("readOnly", searchResult.isReadOnly());
        model.addAttribute("kaisiDtReadOnly", searchResult.isKaisiDtReadOnly());

        // 遷移元区分の表示名
        String seniMotoLabel = switch (searchResult.getSeniMotoKbn()) {
            case Mcm2002uConstants.SENIMOTO_KBN_KENSAKU -> "新規作成";
            case Mcm2002uConstants.SENIMOTO_KBN_SAISENTEI_LINK -> "見積修正";
            case Mcm2002uConstants.SENIMOTO_KBN_FUKUSEI_LINK -> "複製";
            default -> "";
        };
        model.addAttribute("seniMotoLabel", seniMotoLabel);

        log.info("MCM2002U: 画面ロード完了 [遷移元={}] [構成={}行] [明細={}行] [個体={}行] [単価={}行]",
            seniMotoLabel,
            searchResult.getKoseiRows().size(),
            searchResult.getMeisaiRows().size(),
            searchResult.getKotaiRows().size(),
            searchResult.getTankaRows().size());

        return "mcm2002u/index";
    }

    // ================================================================
    // 2. 見積基本設定ボタン (POST)
    // 【変換元】Mcm2002uScreen.vb - KIHONSETTEIButton_Click()（601行）
    //   元コード:
    //     (1) バリデーション（開始日・時間帯）
    //     (2) ConfirmMessage("見積基本設定を行います。…")
    //     (3) ブランドループ → 受渡しデータ構築
    //     (4) ForwardScreen(MCM2003U, delivery)
    //
    //   VB.NET の ConfirmMessage は JavaScript confirm() に変換済み。
    //   HTML側で confirm() → hidden formの submit で実現。
    // ================================================================
    @PostMapping("/kihon-settei")
    public String kihonSettei(
            @ModelAttribute("form") Mcm2002uForm form,
            HttpSession session,
            RedirectAttributes ra) {

        // --- セッションから検索結果を取得 ---
        SearchResult searchResult =
            (SearchResult) session.getAttribute(SESSION_SEARCH_RESULT);

        if (searchResult == null) {
            log.warn("MCM2002U: セッションに検索結果がありません。画面を再ロードします。");
            ra.addFlashAttribute("errors", List.of("セッションが切れました。再度操作してください。"));
            return "redirect:/mcm2002u";
        }

        // --- (1) バリデーション ---
        /**
         * 【変換元】KIHONSETTEIButton_Click() - バリデーション部分
         *   元コード:
         *     If IsNull(tmKeiyakujikanID) Then
         *       If IsNull(KAISI_DTTextBox.Text) Then → エラー
         *       ...
         *     End If
         */
        boolean hasTmKeiyakujikanIds = searchResult.getTmKeiyakujikanIds() != null
            && !searchResult.getTmKeiyakujikanIds().isEmpty();
        List<String> errors = service.validateKihonSettei(form, hasTmKeiyakujikanIds);

        if (!errors.isEmpty()) {
            ra.addFlashAttribute("errors", errors);
            return "redirect:/mcm2002u";
        }

        // --- (3) 受渡しデータ構築 ---
        /**
         * 【変換元】KIHONSETTEIButton_Click() - ブランドループ＋受渡しデータ構築
         *   元コード: For i = 0 To brandKoseiIdList.Length - 1
         *               ...
         *             Next
         *             Dim delivery3u As Mcm2003uDelivery = New Mcm2003uDelivery
         *             ...
         *             MyBase.ForwardScreen(MCM2003U, delivery3u)
         */
        BuildResult buildResult = service.buildDeliveryData(form, searchResult);

        if (!buildResult.getErrors().isEmpty()) {
            ra.addFlashAttribute("errors", buildResult.getErrors());
            return "redirect:/mcm2002u";
        }

        // --- (4) MCM2003U用セッションデータ設定 ---
        Mcm2002uDeliveryDto delivery =
            (Mcm2002uDeliveryDto) session.getAttribute(SESSION_DELIVERY);

        session.setAttribute(SESSION_MCM2003U_BUILD, buildResult);
        session.setAttribute(SESSION_MCM2003U_DELIVERY, delivery);

        log.info("MCM2002U: 基本設定完了 → MCM2003Uへ遷移 " +
            "[ブランド={}件] [構成={}件] [明細={}件] [個体={}件] [単価={}件] [期間={}件]",
            buildResult.getBrandDataList().size(),
            buildResult.getKoseiList().size(),
            buildResult.getMeisaiList().size(),
            buildResult.getKotaiList().size(),
            buildResult.getTankaList().size(),
            buildResult.getKikanList().size());

        return "redirect:/mcm2003u";
    }

    // ================================================================
    // 3. 戻るボタン (GET)
    // 【変換元】Mcm2002uScreen.vb - KENSAKUButton_Click()
    //   元コード:
    //     MyBase.ReDrawScreen()
    //   → MCM2001U（カスタマー見積検索）へ戻る
    // ================================================================
    @GetMapping("/back")
    public String back(HttpSession session) {

        /**
         * 【変換元】KENSAKUButton_Click()
         *   元コード: MyBase.ReDrawScreen()
         *   セッションクリア後、MCM2001Uへリダイレクト
         */
        session.removeAttribute(SESSION_DELIVERY);
        session.removeAttribute(SESSION_SEARCH_RESULT);

        log.info("MCM2002U: MCM2001Uへ戻る");
        return "redirect:/mcm2001u";
    }
}
