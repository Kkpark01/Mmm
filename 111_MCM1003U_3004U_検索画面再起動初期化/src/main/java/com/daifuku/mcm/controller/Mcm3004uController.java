package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;

import jakarta.servlet.http.HttpServletResponse;
import jakarta.servlet.http.HttpSession;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.WebDataBinder;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.InitBinder;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.config.McmUserDetailsService;
import com.daifuku.mcm.constants.Mcm3004uConstants;
import com.daifuku.mcm.dto.UserInfo;
import com.daifuku.mcm.form.Mcm3004uForm;
import com.daifuku.mcm.service.Mcm3004pExcelService;
import com.daifuku.mcm.service.Mcm3004uService;

/**
 * 【変換元】Mcm3004uScreen.vb
 * MCM3004U 作業予定検索コントローラ（9タブ）
 *
 * <p>元イベント対応</p>
 * <table>
 *   <tr><th>VB</th><th>Web</th></tr>
 *   <tr><td>Mcm3004uScreen_Load</td><td>{@code GET  /mcm3004u}</td></tr>
 *   <tr><td>SearchButton1〜9_Click</td><td>{@code POST /mcm3004u/search}（activeTabで分岐）</td></tr>
 *   <tr><td>ExcelOutputButton1〜9_Click</td><td>{@code POST /mcm3004u/excel}（activeTabで分岐）</td></tr>
 *   <tr><td>MCM_MO_SAGYOYOTEIDataGridView_CellContentClick</td>
 *       <td>{@code GET  /mcm3004u/link/mcm3005u?torokuId=…}</td></tr>
 *   <tr><td>tab_SelectedIndexChanged</td>
 *       <td>タブ切替はクライアント側で完結（各タブの検索結果はセッション保持）</td></tr>
 * </table>
 *
 * <p>VB版はタブを切り替えても各タブの DataGridView の内容が保持されるため、
 * Web版もタブごとの検索結果をすべてセッションに保持し、画面上は9つのグリッドを
 * 出力して JavaScript で表示切替する（タブ切替のためのサーバ通信を行わない）。</p>
 */
@Controller
@RequestMapping("/mcm3004u")
public class Mcm3004uController extends BaseController {

    @Autowired
    private Mcm3004uService service;

    @Autowired
    private Mcm3004pExcelService excelService;

    /** 利用権限区分3（取引先契約関連）の取得用 */
    @Autowired
    private McmUserDetailsService userDetailsService;

    private static final String SESSION_FORM = "MCM3004U_FORM";

    /** 画面から受け取るのは検索条件とアクティブタブのみに限定する。 */
    @InitBinder
    public void bindConditions(WebDataBinder binder) {
        binder.setAllowedFields(
                "activeTab",
                "kijunZuki1", "taishoPeriod1",
                "kijunZuki2", "taishoPeriod2", "jokenKbn2",
                "kijunZuki3", "taishoPeriod3",
                "kijunZuki4", "taishoPeriod4",
                "kijunZuki5", "taishoPeriod5", "fusokuKbn5",
                "kijunZuki6", "taishoPeriod6",
                "kijunZuki7", "taishoPeriod7", "kanryo7",
                "kijunZuki8", "taishoPeriod8",
                "kaisiDt9", "torihikisakiCd9", "torihikisakimei9");
    }

    // =====================================================================
    // 初期表示
    // 【変換元】Mcm3004uScreen_Load
    //   各タブの基準月＝システム日付(yyyy/MM)、対象期間＝4（タブ6のみ13）、
    //   タブ9の日付＝システム日付(yyyy/MM/dd)、取引先コード／取引先名は空。
    //   → Mcm3004uForm のフィールド初期値で表現している。
    //   利用権限区分3が0以下の場合はタブ9の金額4列を非表示にする。
    // =====================================================================

    @GetMapping
    public String index(Model model, HttpSession session) {
        Mcm3004uForm form = (Mcm3004uForm) session.getAttribute(SESSION_FORM);
        if (form == null) {
            form = new Mcm3004uForm();
            session.setAttribute(SESSION_FORM, form);
        }
        form.setTab9KingakuVisible(isKingakuVisible(session));
        setCommonAttributes(model, session);
        model.addAttribute("form", form);
        model.addAttribute("tabNames", Mcm3004uConstants.TAB_NAMES);
        return "mcm3004u/index";
    }

    /** メニューからの新規起動は全9タブを初期化。3005Uからの戻りでは通常URLを使う。 */
    @GetMapping("/open")
    public String open(HttpSession session) {
        session.removeAttribute(SESSION_FORM);
        return "redirect:/mcm3004u";
    }

    @GetMapping("/close")
    public String close(HttpSession session) {
        session.removeAttribute(SESSION_FORM);
        return "redirect:/menu";
    }

    /** メニューの副作用を実行せず、検索画面自身の状態だけを破棄する。 */
    @GetMapping("/reset")
    @org.springframework.web.bind.annotation.ResponseBody
    public org.springframework.http.ResponseEntity<Void> reset(HttpSession session) {
        session.removeAttribute(SESSION_FORM);
        return org.springframework.http.ResponseEntity.noContent().build();
    }

    // =====================================================================
    // 検索
    // 【変換元】SearchButton1_Click 〜 SearchButton9_Click
    //   0件時は MSG_0002 を表示し、該当タブの DataSet を Clear する。
    // =====================================================================

    @PostMapping("/search")
    public String search(@ModelAttribute Mcm3004uForm form,
                         HttpSession session, RedirectAttributes ra) {
        Mcm3004uForm merged = merge(form, session);
        try {
            service.search(merged);
            if (merged.getActiveRowCount() == 0) {
                // VB: MyBase.DisplayMessage(MSG_0002) → DataSet.Clear()
                merged.clearActiveRows();
                ra.addFlashAttribute("error", Mcm3004uConstants.MSG_NO_DATA);
            }
        } catch (IllegalArgumentException e) {
            // VB: HissoError / HissuError2 内の DisplayMessage 相当（入力エラー）
            merged.clearActiveRows();
            ra.addFlashAttribute("error", e.getMessage());
        } catch (Exception e) {
            logger.error("MCM3004U 検索に失敗しました（tab={}）", merged.getActiveTab(), e);
            merged.clearActiveRows();
            ra.addFlashAttribute("error", "検索処理でエラーが発生しました。");
        }
        session.setAttribute(SESSION_FORM, merged);
        return "redirect:/mcm3004u";
    }

    // =====================================================================
    // Excel出力
    // 【変換元】ExcelOutputButton1_Click 〜 ExcelOutputButton9_Click
    //   グリッドが0件なら MSG_0002 を表示して Return（帳票は起動しない）。
    //   VB版は「現在表示されているグリッドの内容」をそのまま MCM3004P へ渡すため、
    //   Web版も再検索せずセッション上の検索結果を出力する。
    // =====================================================================

    @PostMapping("/excel")
    public String excel(@ModelAttribute Mcm3004uForm form, HttpSession session,
                        RedirectAttributes ra, HttpServletResponse response) {
        Mcm3004uForm saved = (Mcm3004uForm) session.getAttribute(SESSION_FORM);
        if (saved == null) {
            ra.addFlashAttribute("error", Mcm3004uConstants.MSG_NO_DATA);
            return "redirect:/mcm3004u";
        }
        // 出力対象タブは画面のアクティブタブに合わせる（検索条件はセッションの値を維持）
        saved.setActiveTab(normalizeTab(form.getActiveTab()));
        saved.setTab9KingakuVisible(isKingakuVisible(session));
        session.setAttribute(SESSION_FORM, saved);

        if (saved.getActiveRowCount() == 0) {
            ra.addFlashAttribute("error", Mcm3004uConstants.MSG_NO_DATA);
            return "redirect:/mcm3004u";
        }

        try {
            byte[] data = excelService.generate(saved);
            String filename = excelService.fileName(saved.getActiveTab());
            response.setContentType(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader(HttpHeaders.CONTENT_DISPOSITION, ContentDisposition.attachment()
                    .filename(filename, StandardCharsets.UTF_8).build().toString());
            response.setHeader(HttpHeaders.CACHE_CONTROL, "no-store");
            response.setContentLength(data.length);
            response.getOutputStream().write(data);
            response.getOutputStream().flush();
            return null;
        } catch (Exception e) {
            logger.error("MCM3004U Excel出力に失敗しました（tab={}）", saved.getActiveTab(), e);
            if (response.isCommitted()) {
                return null;
            }
            response.reset();
            ra.addFlashAttribute("error", "Excelを出力できませんでした。");
            return "redirect:/mcm3004u";
        }
    }

    // =====================================================================
    // MCM3005U（作業予定登録）への遷移
    // 【変換元】MCM_MO_SAGYOYOTEIDataGridView_CellContentClick
    //   選択行の TOROKU_ID を Mcm3005uDelivery に格納し ForwardScreen(MCM3005U)
    // =====================================================================

    @GetMapping("/link/mcm3005u")
    public String linkMcm3005u(@RequestParam(required = false) BigDecimal torokuId,
                               HttpSession session) {
        Map<String, Object> delivery = new HashMap<>();
        delivery.put("torokuId", torokuId != null ? torokuId : BigDecimal.ZERO);
        session.setAttribute("MCM3005U_DELIVERY", delivery);
        return "redirect:/mcm3005u";
    }

    // =====================================================================
    // 内部処理
    // =====================================================================

    /**
     * 画面から送信された検索条件をセッション上のフォームへ反映する。
     * 検索結果（他タブ分を含む）はセッション側を正とし、POSTで上書きさせない。
     */
    private Mcm3004uForm merge(Mcm3004uForm posted, HttpSession session) {
        Mcm3004uForm saved = (Mcm3004uForm) session.getAttribute(SESSION_FORM);
        if (saved == null) {
            saved = new Mcm3004uForm();
        }
        saved.setActiveTab(normalizeTab(posted.getActiveTab()));

        saved.setKijunZuki1(posted.getKijunZuki1());
        saved.setTaishoPeriod1(posted.getTaishoPeriod1());

        saved.setKijunZuki2(posted.getKijunZuki2());
        saved.setTaishoPeriod2(posted.getTaishoPeriod2());
        saved.setJokenKbn2(posted.getJokenKbn2());

        saved.setKijunZuki3(posted.getKijunZuki3());
        saved.setTaishoPeriod3(posted.getTaishoPeriod3());

        saved.setKijunZuki4(posted.getKijunZuki4());
        saved.setTaishoPeriod4(posted.getTaishoPeriod4());

        saved.setKijunZuki5(posted.getKijunZuki5());
        saved.setTaishoPeriod5(posted.getTaishoPeriod5());
        saved.setFusokuKbn5(posted.getFusokuKbn5());

        saved.setKijunZuki6(posted.getKijunZuki6());
        saved.setTaishoPeriod6(posted.getTaishoPeriod6());

        saved.setKijunZuki7(posted.getKijunZuki7());
        saved.setTaishoPeriod7(posted.getTaishoPeriod7());
        saved.setKanryo7(posted.isKanryo7());

        saved.setKijunZuki8(posted.getKijunZuki8());
        saved.setTaishoPeriod8(posted.getTaishoPeriod8());

        saved.setKaisiDt9(posted.getKaisiDt9());
        saved.setTorihikisakiCd9(posted.getTorihikisakiCd9());
        saved.setTorihikisakimei9(posted.getTorihikisakimei9());

        saved.setTab9KingakuVisible(isKingakuVisible(session));
        return saved;
    }

    private int normalizeTab(int tabNo) {
        if (tabNo < Mcm3004uConstants.TAB_MIN || tabNo > Mcm3004uConstants.TAB_MAX) {
            return Mcm3004uConstants.TAB_MIN;
        }
        return tabNo;
    }

    /**
     * タブ9の金額4列を表示できるか判定する。
     * 【変換元】Mcm3004uScreen_Load
     *   If CPSettingInfo.GetUserInfo.RiyoKengen3 &lt;= 0 Then 金額4列 Visible = False
     */
    private boolean isKingakuVisible(HttpSession session) {
        UserInfo userInfo = getLoginUserInfo(session);
        if (userInfo == null) {
            logger.warn("[MCM3004U] ログインユーザー情報が取得できないため金額4列を非表示にします");
            return false;
        }
        // 【不具合修正】セッション上の UserInfo.riyoKengen3 はログイン時にしか設定されず、
        //   修正前にログインしたセッション（DevTools のセッション永続化で再起動後も残る）では
        //   null のままとなり、権限「作成」でも金額4列が画面・Excel出力から欠落していた。
        //   VB版 RiyoKengen3（McmLogin_UserInfo.xml: MCM_MO_TANTOKENGEN.KENGENBUNRUI_ID = 3）と
        //   同じ値を DB から都度取得して判定する。
        String kengen = userDetailsService.getRiyokengenKbn(userInfo.getUserId(), KENGENBUNRUI_ID_TORIHIKISAKI_KEIYAKU);
        userInfo.setRiyoKengen3(kengen);
        boolean visible;
        try {
            visible = kengen != null && Integer.parseInt(kengen.trim()) > 0;
        } catch (NumberFormatException e) {
            visible = false;
        }
        logger.debug("[MCM3004U] userId={}, RiyoKengen3={}, 金額4列表示={}", userInfo.getUserId(), kengen, visible);
        return visible;
    }

    /** 権限分類ID：取引先契約関連（VB: RiyoKengen3） */
    private static final int KENGENBUNRUI_ID_TORIHIKISAKI_KEIYAKU = 3;

    @Override
    protected String getScreenTitle() { return "作業予定検索"; }

    @Override
    protected String getFunctionId() { return "MCM3004U"; }
}
