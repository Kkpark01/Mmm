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

    @Autowired
    private Mcm1005uService service;

    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");

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
                        Model model, HttpSession session) {

        McmTkKeiyakuEntity keiyaku;
        List<McmTkKikanEntity> kikanList;

        if (tkKeiyakuId != null) {
            // 更新モード: SearchUpdate()
            keiyaku = service.loadKeiyaku(tkKeiyakuId);
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
            if (result != null) {
                try {
                    service.applyMcm1006uResult(result, getLoginUserId());
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

        if (!kikanList.isEmpty() && selectedKikanIndex < kikanList.size()) {
            selectedKikan = kikanList.get(selectedKikanIndex);
            BigDecimal kikanId = selectedKikan.getTkKikanId();
            tenkenList = service.loadTenkenList(kikanId);
            siharai = service.loadSiharai(kikanId);
            tenpuList = service.loadTenpuList(kikanId);
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

        String backUrl = "redirect:/mcm1005u?tkKeiyakuId=" + tkKeiyakuId
                + (torihikisakiNk != null ? "&torihikisakiNk=" + torihikisakiNk : "")
                + (nonyusakiNk != null ? "&nonyusakiNk=" + nonyusakiNk : "")
                + "&selectedKikanIndex=" + selectedKikanIndex;

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
            @RequestParam(value = "siharaiId", required = false) BigDecimal siharaiId,
            @RequestParam(value = "siharaiKikanId", required = false) BigDecimal siharaiKikanId,
            @RequestParam(value = "siharaiKaisu", required = false) BigDecimal siharaiKaisu,
            @RequestParam(value = "siharaiTsuki", required = false) List<Integer> siharaiTsukiMonths,
            @RequestParam(value = "siharaiBiko", required = false) String siharaiBiko,
            // ── 添付ファイル ──
            @RequestParam(value = "tenpuIds", required = false) List<BigDecimal> tenpuIds,
            @RequestParam(value = "tenpuKikanIds", required = false) List<BigDecimal> tenpuKikanIds,
            @RequestParam(value = "tenpuFileNks", required = false) List<String> tenpuFileNks,
            @RequestParam(value = "tenpuDirectories", required = false) List<String> tenpuDirectories,
            @RequestParam(value = "tenpuShoninjotais", required = false) List<String> tenpuShoninjotais,
            @RequestParam(value = "torihikisakiNk", required = false) String torihikisakiNk,
            @RequestParam(value = "nonyusakiNk", required = false) String nonyusakiNk,
            @RequestParam(value = "selectedKikanIndex", defaultValue = "0") int selectedKikanIndex,
            RedirectAttributes ra, HttpSession session) {

        // ── 契約ヘッダー組み立て ──
        McmTkKeiyakuEntity keiyaku = new McmTkKeiyakuEntity();
        keiyaku.setTkKeiyakuId(tkKeiyakuId != null ? tkKeiyakuId : service.getNextKeiyakuId());
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
                tenkenList.add(t);
            }
        }

        // ── 支払情報組み立て ──
        List<McmTkSiharaiEntity> siharaiList = new ArrayList<>();
        List<Boolean[]> siharaiTsukiList = new ArrayList<>();
        if (siharaiKikanId != null) {
            McmTkSiharaiEntity s = new McmTkSiharaiEntity();
            s.setTkSiharaiId(siharaiId);
            s.setTkKikanId(siharaiKikanId);
            s.setKaisu(siharaiKaisu);
            s.setBiko(siharaiBiko);
            siharaiList.add(s);
            Boolean[] tsuki = new Boolean[12];
            for (int i = 0; i < 12; i++) tsuki[i] = Boolean.FALSE;
            if (siharaiTsukiMonths != null) {
                for (int m : siharaiTsukiMonths) {
                    if (m >= 1 && m <= 12) tsuki[m - 1] = Boolean.TRUE;
                }
            }
            siharaiTsukiList.add(tsuki);
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

        try {
            service.save(keiyaku, kikanList, tenkenList, siharaiList, siharaiTsukiList, tenpuList,
                    getLoginUserId());
            ra.addFlashAttribute("message", "登録を完了しました。");
        } catch (McmBusinessException ex) {
            ra.addFlashAttribute("error", ex.getMessage());
        }

        return "redirect:/mcm1005u?tkKeiyakuId=" + keiyaku.getTkKeiyakuId()
                + (torihikisakiNk != null ? "&torihikisakiNk=" + torihikisakiNk : "")
                + (nonyusakiNk != null ? "&nonyusakiNk=" + nonyusakiNk : "")
                + "&selectedKikanIndex=" + selectedKikanIndex;
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
        try {
            service.sinsei(tkKeiyakuId, tkKikanId, getLoginUserId());
            ra.addFlashAttribute("message", "申請を完了しました。");
        } catch (McmBusinessException ex) {
            ra.addFlashAttribute("error", ex.getMessage());
        }
        return "redirect:/mcm1005u?tkKeiyakuId=" + tkKeiyakuId
                + (torihikisakiNk != null ? "&torihikisakiNk=" + torihikisakiNk : "")
                + (nonyusakiNk != null ? "&nonyusakiNk=" + nonyusakiNk : "")
                + "&selectedKikanIndex=" + selectedKikanIndex;
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
        service.deleteTenpu(tkTenpuId, getLoginUserId());
        return "redirect:/mcm1005u?tkKeiyakuId=" + tkKeiyakuId
                + (torihikisakiNk != null ? "&torihikisakiNk=" + torihikisakiNk : "")
                + (nonyusakiNk != null ? "&nonyusakiNk=" + nonyusakiNk : "")
                + "&selectedKikanIndex=" + selectedKikanIndex;
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
            return LocalDate.parse(s.trim(), DATE_FMT).atStartOfDay();
        } catch (Exception e) {
            return null;
        }
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
