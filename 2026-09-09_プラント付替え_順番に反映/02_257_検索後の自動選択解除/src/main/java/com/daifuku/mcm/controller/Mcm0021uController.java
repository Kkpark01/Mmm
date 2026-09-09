/**
 * 【変換元】Mcm0021uScreen.vb（724行）/ Mcm0021uScreen.Designer.vb（1,945行）
 *   MCM0021U プラント付替え コントローラ
 *   元コード: Mcm0021uScreen_Load / SearchButton_Click /
 *             TsukekaemotoButton_Click / TsukekaesakiButton_Click / ReplaceButton_Click
 *
 *   各ボタンイベントをHTTPエンドポイントにマッピング。
 *   付替え元/先のプラントIDはHttpSessionで保持。
 *   PRGパターン（Post-Redirect-Get）を採用。
 */
package com.daifuku.mcm.controller;

import java.math.BigDecimal;
import java.util.List;

import jakarta.servlet.http.HttpSession;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.constants.Mcm0021uConstants;
import com.daifuku.mcm.dto.Mcm0021uKikikoseiDto;
import com.daifuku.mcm.dto.Mcm0021uKikimeisaiSearchDto;
import com.daifuku.mcm.dto.Mcm0021uMotoRowDto;
import com.daifuku.mcm.dto.Mcm0021uNonyusakiDto;
import com.daifuku.mcm.dto.Mcm0021uSakiRowDto;
import com.daifuku.mcm.form.Mcm0021uForm;
import com.daifuku.mcm.form.Mcm0021uForm.MotoRowForm;
import com.daifuku.mcm.service.Mcm0021uService;

/**
 * MCM0021U（プラント付替え）画面のコントローラ
 * <p>画面URL: /mcm0021u</p>
 * <p>エンドポイント一覧:</p>
 * <ul>
 *   <li>GET  /mcm0021u           - 画面初期表示</li>
 *   <li>POST /mcm0021u/search    - 検索</li>
 *   <li>POST /mcm0021u/setMoto   - 付替え元に設定</li>
 *   <li>POST /mcm0021u/setSaki   - 付替え先に設定</li>
 *   <li>POST /mcm0021u/replace   - 付替え実行</li>
 *   <li>POST /mcm0021u/checkKeiyaku - 契約チェックAPI（Ajax）</li>
 * </ul>
 */
@Controller
@RequestMapping("/mcm0021u")
public class Mcm0021uController extends BaseController {

    private final Mcm0021uService service;
    private static final String SESSION_SEARCH = "mcm0021u_search";
    private static final String SESSION_COMPLETED = "mcm0021u_replaceCompleted";

    /** この画面で最後に完了した検索結果。検索条件と選択位置を元/先設定後も保持する。 */
    private static class SearchState {
        final String token = java.util.UUID.randomUUID().toString();
        final Mcm0021uForm criteria;
        final List<Mcm0021uNonyusakiDto> plants;
        final List<Mcm0021uKikikoseiDto> compositions;
        final List<Mcm0021uKikimeisaiSearchDto> details;
        BigDecimal selectedPlant, selectedComposition, selectedBrand;
        SearchState(Mcm0021uForm form, List<Mcm0021uNonyusakiDto> plants,
                    List<Mcm0021uKikikoseiDto> compositions, List<Mcm0021uKikimeisaiSearchDto> details) {
            criteria = new Mcm0021uForm();
            criteria.setNonyusakiCd(form.getNonyusakiCd()); criteria.setNonyusakiNk(form.getNonyusakiNk());
            criteria.setSupportId(form.getSupportId()); criteria.setPlantNk(form.getPlantNk());
            this.plants = List.copyOf(plants); this.compositions = List.copyOf(compositions); this.details = List.copyOf(details);
            // #257: 新しい検索は未選択で開始する。選択は利用者が元/先を設定した時に保存する。
        }
        void select(BigDecimal plant, BigDecimal composition, BigDecimal brand) {
            selectedPlant = plant;
            var candidates = compositions.stream().filter(k -> sameId(k.getPlantId(), plant)).toList();
            var row = candidates.stream().filter(k -> sameId(k.getKikikoseiId(), composition)
                    && (sameId(k.getBrandkoseiId(), brand) || (k.getBrandkoseiId() == null && brand == null)))
                    .findFirst().orElse(candidates.isEmpty() ? null : candidates.get(0));
            selectedComposition = row == null ? null : row.getKikikoseiId();
            selectedBrand = row == null ? null : row.getBrandkoseiId();
        }
    }
    private static boolean sameId(BigDecimal a, BigDecimal b) { return a != null && b != null && a.compareTo(b) == 0; }
    private static BigDecimal parseId(String value) {
        try { return value == null || value.isBlank() ? null : new BigDecimal(value.trim()); }
        catch (NumberFormatException ex) { return null; }
    }
    private static SearchState searchState(HttpSession session) { return (SearchState) session.getAttribute(SESSION_SEARCH); }
    private static Mcm0021uNonyusakiDto searchedPlant(HttpSession session, String token, BigDecimal id) {
        var state = searchState(session);
        if (state == null || !state.token.equals(token) || id == null || id.signum() <= 0) return null;
        return state.plants.stream().filter(p -> sameId(p.getPlantId(), id)).findFirst().orElse(null);
    }

    // セッションキー定数
    /** 【変換元】Private plantIdMoto As Integer */
    private static final String SESSION_PLANT_ID_MOTO = "mcm0021u_plantIdMoto";
    /** 【変換元】Private plantIdSaki As Integer */
    private static final String SESSION_PLANT_ID_SAKI = "mcm0021u_plantIdSaki";
    /** 【変換元】MCM_MA_KIKIKOSEI_MOTODataTable */
    private static final String SESSION_MOTO_ROWS = "mcm0021u_motoRows";
    /** 【変換元】MCM_MA_KIKIKOSEI_SAKIDataTable */
    private static final String SESSION_SAKI_ROWS = "mcm0021u_sakiRows";
    /** 【変換元】MCM_MA_BRAND_KOSEI_FOR_COMBODataTable */
    private static final String SESSION_BRAND_COMBO = "mcm0021u_brandCombo";
    /** 付替え先の納入先表示情報 */
    private static final String SESSION_SAKI_NONYUSAKI = "mcm0021u_sakiNonyusaki";

    public Mcm0021uController(Mcm0021uService service) {
        this.service = service;
    }

    // =================================================================
    // 画面表示（GET）
    // 【変換元】Mcm0021uScreen_Load()
    //   元コード: Dim delivery As Mcm0021uDelivery = MyBase.GetDeliveryData()
    //             If IsNotNull(delivery) Then plantId = delivery.PLANT_ID
    //             MCM_MA_KIKIKOSEI_MOTOTableAdapter.Fill(... plantId)
    // =================================================================
    /**
     * 画面初期表示
     * @param plantId 連携値プラントID（MCM0013U等からの遷移時）
     */
    @GetMapping
    public String index(@RequestParam(required = false) BigDecimal plantId,
                        @RequestParam(defaultValue = "false") boolean resume,
                        Model model, HttpSession session) {

        if (!resume || plantId != null) session.removeAttribute(SESSION_SEARCH);
        Mcm0021uForm form = new Mcm0021uForm();

        /*
         * 【変換元】連携値プラントIDを取得した場合、付替え元を検索する
         *   元コード: If IsNotNull(delivery) Then
         *               plantId = delivery.PLANT_ID
         *               MCM_MA_KIKIKOSEI_MOTOTableAdapter.Fill(... plantId)
         */
        if (plantId != null) {
            session.removeAttribute(SESSION_COMPLETED);
            List<Mcm0021uMotoRowDto> motoRows = service.findMotoByPlantId(plantId);
            session.setAttribute(SESSION_PLANT_ID_MOTO, plantId);
            session.setAttribute(SESSION_MOTO_ROWS, motoRows);
            form.setPlantIdMoto(plantId);
        }
        model.addAttribute("form", form);
        model.addAttribute("searched", false);
        model.addAttribute("nonyusakiList", List.of());
        model.addAttribute("kikikoseiList", List.of());
        model.addAttribute("kikimeisaiList", List.of());
        addSessionDataToModel(model, session);
        return "mcm0021u/index";

    }

    // =================================================================
    // 検索（POST）
    // 【変換元】SearchButton_Click()
    //   元コード: 全項目空白チェック → 3テーブル同時検索 → 0件チェック
    // =================================================================
    /**
     * 検索ボタン処理
     */
    @PostMapping("/search")
    public String search(@ModelAttribute Mcm0021uForm form,
                         Model model, HttpSession session, RedirectAttributes ra) {
        // 再検索が空・0件・失敗でも以前の検索対象を利用させない。
        session.removeAttribute(SESSION_SEARCH);

        /*
         * 【変換元】入力項目全てが空白かチェックする（#181対応）
         *   元コード: If IsNull(NONYUSAKI_CDTextBox.Text) And IsNull(SUPPORT_IDTextBox.Text)
         *             And IsNull(NONYUSAKI_NKTextBox.Text) And IsNull(PLANT_NKTextBox.Text) Then
         *               DisplayMessage(CPMessageConstant.MSG_0001)
         */
        if (isEmpty(form.getNonyusakiCd()) && isEmpty(form.getSupportId())
                && isEmpty(form.getNonyusakiNk()) && isEmpty(form.getPlantNk())) {
            ra.addFlashAttribute("errors",
                    List.of(Mcm0021uConstants.MSG_SEARCH_EMPTY));
            return "redirect:/mcm0021u";
        }

        /*
         * 【変換元】3テーブル同時検索
         *   元コード: Me.Fill(MCM_MA_NONYUSAKI, ...) → nonyuCount
         *             Me.Fill(MCM_MA_KIKIKOSEI, ...)
         *             Me.Fill(MCM_MA_KIKIMEISAI, ...)
         */
        List<Mcm0021uNonyusakiDto> nonyusakiList = service.searchNonyusaki(
                form.getNonyusakiCd(), form.getNonyusakiNk(),
                form.getSupportId(), form.getPlantNk());

        /*
         * 【変換元】検索結果0件チェック（#180対応）
         *   元コード: If nonyuCount = 0 Then DisplayMessage(CPMessageConstant.MSG_0002)
         *   検索処理は正常終了し、入力条件を保持して一覧を0件表示する。
         */
        if (nonyusakiList.isEmpty()) {
            session.setAttribute(SESSION_SEARCH, new SearchState(form, List.of(), List.of(), List.of()));
            model.addAttribute("form", form);
            model.addAttribute("errors",
                    List.of(Mcm0021uConstants.MSG_NO_DATA));
            model.addAttribute("searched", true);
            model.addAttribute("nonyusakiList", List.of());
            model.addAttribute("kikikoseiList", List.of());
            model.addAttribute("kikimeisaiList", List.of());
            addSessionDataToModel(model, session);
            return "mcm0021u/index";
        }

        List<Mcm0021uKikikoseiDto> kikikoseiList = service.searchKikikosei(
                form.getNonyusakiCd(), form.getNonyusakiNk(),
                form.getSupportId(), form.getPlantNk());

        List<Mcm0021uKikimeisaiSearchDto> kikimeisaiList = service.searchKikimeisai(
                form.getNonyusakiCd(), form.getNonyusakiNk(),
                form.getSupportId(), form.getPlantNk());

        model.addAttribute("form", form);
        model.addAttribute("searched", true);
        session.setAttribute(SESSION_SEARCH, new SearchState(form, nonyusakiList, kikikoseiList, kikimeisaiList));
        model.addAttribute("nonyusakiList", nonyusakiList);
        model.addAttribute("kikikoseiList", kikikoseiList);
        model.addAttribute("kikimeisaiList", kikimeisaiList);
        addSessionDataToModel(model, session);

        return "mcm0021u/index";
    }

    // =================================================================
    // 付替え元に設定（POST）
    // 【変換元】TsukekaemotoButton_Click()
    //   元コード: plantId = MCM_MA_NONYUSAKIDataGridView(COL_PLANT_ID_NONYUSAKI)
    //             MCM_MA_KIKIKOSEI_MOTOTableAdapter.Fill(... plantId)
    //             Me.plantIdMoto = plantId
    // =================================================================
    /**
     * 付替え元に設定ボタン処理
     */
    @PostMapping("/setMoto")
    public String setMoto(@RequestParam(name = "plantId", required = false) String rawPlantId,
                          @RequestParam(required = false) String searchToken,
                          @RequestParam(required = false) String selectedKikikoseiId,
                          @RequestParam(required = false) String selectedBrandkoseiId,
                          HttpSession session, RedirectAttributes ra) {
        BigDecimal plantId = parseId(rawPlantId);
        var searched = searchState(session);
        if (searched != null && searched.token.equals(searchToken) && !searched.plants.isEmpty()
                && (rawPlantId == null || rawPlantId.isBlank() || (plantId != null && plantId.signum() == 0))) {
            ra.addFlashAttribute("errors", List.of(Mcm0021uConstants.MSG_MOTO_NOT_CHECKED));
            return "redirect:/mcm0021u?resume=true";
        }
        if (searchedPlant(session, searchToken, plantId) == null) {
            ra.addFlashAttribute("errors",
                    List.of(Mcm0021uConstants.MSG_NOT_SEARCHED));
            return "redirect:/mcm0021u?resume=true";
        }

        var state = searchState(session);
        state.select(plantId, parseId(selectedKikikoseiId), parseId(selectedBrandkoseiId));
        if (state.selectedComposition == null || !service.hasKikiDetails(state.selectedComposition)) {
            ra.addFlashAttribute("errors", List.of(Mcm0021uConstants.MSG_NO_KIKI_DATA));
            return "redirect:/mcm0021u?resume=true";
        }
        List<Mcm0021uMotoRowDto> motoRows = service.findMotoByPlantId(plantId);

        /*
         * 【変換元】機器構成情報チェック
         *   元コード: If MCM_MA_KIKIKOSEI_MOTODataGridView.RowCount = 0 Then
         *               DisplayMessage(MSG_0045)
         */
        if (motoRows.isEmpty()) {
            ra.addFlashAttribute("errors",
                    List.of(Mcm0021uConstants.MSG_NO_KIKI_DATA));
            return "redirect:/mcm0021u?resume=true";
        }

        session.setAttribute(SESSION_PLANT_ID_MOTO, plantId);
        session.setAttribute(SESSION_MOTO_ROWS, motoRows);
        session.removeAttribute(SESSION_COMPLETED);
        ra.addFlashAttribute("infoMessage", "付替え元に設定しました。");

        return "redirect:/mcm0021u?resume=true";
    }

    // =================================================================
    // 付替え先に設定（POST）
    // 【変換元】TsukekaesakiButton_Click()
    //   元コード: plantId = MCM_MA_NONYUSAKIDataGridView(COL_PLANT_ID_NONYUSAKI)
    //             MCM_MA_KIKIKOSEI_SAKITableAdapter.Fill(... plantId)
    //             MCM_MA_BRAND_KOSEI_FOR_COMBOTableAdapter.Fill(... plantId)
    //             Me.plantIdSaki = plantId
    // =================================================================
    /**
     * 付替え先に設定ボタン処理
     */
    @PostMapping("/setSaki")
    public String setSaki(@RequestParam(name = "plantId", required = false) String rawPlantId,
                          @RequestParam(required = false) String searchToken,
                          @RequestParam(required = false) String selectedKikikoseiId,
                          @RequestParam(required = false) String selectedBrandkoseiId,
                          @RequestParam(required = false) String nonyusakiCd,
                          @RequestParam(required = false) String nonyusakiNk,
                          @RequestParam(required = false) String nonyusakikojoNk,
                          @RequestParam(required = false) String supportId,
                          @RequestParam(required = false) String plantNk,
                          HttpSession session, RedirectAttributes ra) {

        BigDecimal plantId = parseId(rawPlantId);
        var selected = searchedPlant(session, searchToken, plantId);
        if (selected == null) {
            ra.addFlashAttribute("errors",
                    List.of(Mcm0021uConstants.MSG_NOT_SEARCHED));
            return "redirect:/mcm0021u?resume=true";
        }

        searchState(session).select(plantId, parseId(selectedKikikoseiId), parseId(selectedBrandkoseiId));
        // 表示情報も検索結果から取得する（hiddenの書換えを表示へ反映しない）。
        nonyusakiCd = selected.getNonyusakiCd(); nonyusakiNk = selected.getNonyusakiNk();
        nonyusakikojoNk = selected.getNonyusakikojoNk(); supportId = selected.getSupportId(); plantNk = selected.getPlantNk();
        List<Mcm0021uSakiRowDto> sakiRows = service.findSakiByPlantId(plantId);
        List<Object[]> brandCombo = service.findBrandKoseiForCombo(plantId);

        session.setAttribute(SESSION_PLANT_ID_SAKI, plantId);
        session.setAttribute(SESSION_SAKI_ROWS, sakiRows);
        session.setAttribute(SESSION_BRAND_COMBO, brandCombo);

        // 付替え先の納入先情報をセッション保持（画面表示用）
        java.util.HashMap<String, String> sakiNonyusaki = new java.util.HashMap<>();
        sakiNonyusaki.put("nonyusakiCd", nonyusakiCd != null ? nonyusakiCd : "");
        sakiNonyusaki.put("nonyusakiNk", nonyusakiNk != null ? nonyusakiNk : "");
        sakiNonyusaki.put("nonyusakikojoNk", nonyusakikojoNk != null ? nonyusakikojoNk : "");
        sakiNonyusaki.put("supportId", supportId != null ? supportId : "");
        sakiNonyusaki.put("plantNk", plantNk != null ? plantNk : "");
        session.setAttribute(SESSION_SAKI_NONYUSAKI, sakiNonyusaki);

        ra.addFlashAttribute("infoMessage", "付替え先に設定しました。");
        return "redirect:/mcm0021u?resume=true";
    }

    // =================================================================
    // 付替え実行（POST）
    // 【変換元】ReplaceButton_Click()
    //   元コード: バリデーション → 契約チェック → 確認ダイアログ
    //             → 付替え実行 → UpdateAll → 元/先の設定保持（#265）
    // =================================================================
    /**
     * 付替え実行ボタン処理（#194対応: BindingResult追加でバインドエラー防止）
     */
    @PostMapping("/replace")
    public String replace(@ModelAttribute Mcm0021uForm form,
                          BindingResult bindingResult,
                          Model model, HttpSession session, RedirectAttributes ra) {
        synchronized (session) {
            return replaceLocked(form, bindingResult, model, session, ra);
        }
    }

    private String replaceLocked(Mcm0021uForm form, BindingResult bindingResult,
                                 Model model, HttpSession session, RedirectAttributes ra) {

        if (Boolean.TRUE.equals(session.getAttribute(SESSION_COMPLETED))) {
            ra.addFlashAttribute("errors", List.of(Mcm0021uConstants.MSG_TARGET_CHANGED));
            return "redirect:/mcm0021u?resume=true";
        }

        BigDecimal plantIdMoto = (BigDecimal) session.getAttribute(SESSION_PLANT_ID_MOTO);
        BigDecimal plantIdSaki = (BigDecimal) session.getAttribute(SESSION_PLANT_ID_SAKI);

        // バインドエラー検知時（型変換エラー等によるシステムエラーを防止し、メッセージを表示）
        // #194対応: redirectではなくforwardすることで、送信済みの選択内容（form）を保持したまま再表示する
        if (bindingResult.hasErrors()) {
            model.addAttribute("errors", List.of(bindingResult.getFieldErrors().stream()
                    .anyMatch(e -> e.getField().endsWith(".ikosaki"))
                    ? Mcm0021uConstants.MSG_IKOSAKI_NOT_SELECTED : Mcm0021uConstants.MSG_TARGET_CHANGED));
            model.addAttribute("form", form);
            model.addAttribute("searched", false);
            model.addAttribute("nonyusakiList", List.of());
            model.addAttribute("kikikoseiList", List.of());
            model.addAttribute("kikimeisaiList", List.of());
            addSessionDataToModel(model, session);
            return "mcm0021u/index";
        }

        // バリデーション
        // #194対応: redirectではなくforwardすることで、送信済みの選択内容（form）を保持したまま再表示する
        List<String> errors = service.validateReplace(
                plantIdMoto, plantIdSaki, form.getMotoRows());
        if (!errors.isEmpty()) {
            model.addAttribute("errors", errors);
            model.addAttribute("form", form);
            model.addAttribute("searched", false);
            model.addAttribute("nonyusakiList", List.of());
            model.addAttribute("kikikoseiList", List.of());
            model.addAttribute("kikimeisaiList", List.of());
            addSessionDataToModel(model, session);
            return "mcm0021u/index";
        }

        // 付替え実行
        String loginUser = getLoginUserId();
        String errorMsg;
        try {
            errorMsg = service.executeReplace(plantIdMoto, plantIdSaki, form.getMotoRows(), loginUser);
        } catch (IllegalStateException ex) {
            errorMsg = Mcm0021uConstants.MSG_TARGET_CHANGED;
        }

        if (errorMsg != null) {
            model.addAttribute("errors", List.of(errorMsg));
            model.addAttribute("form", form);
            model.addAttribute("searched", false);
            model.addAttribute("nonyusakiList", List.of());
            model.addAttribute("kikikoseiList", List.of());
            model.addAttribute("kikimeisaiList", List.of());
            addSessionDataToModel(model, session);
            return "mcm0021u/index";
        }

        // #265: 元・先と選択内容を保持する。元は実行時の設定、先は更新後のDB内容を表示。
        @SuppressWarnings("unchecked")
        var source = (List<Mcm0021uMotoRowDto>) session.getAttribute(SESSION_MOTO_ROWS);
        session.setAttribute(SESSION_MOTO_ROWS, withSelections(source, form));
        session.setAttribute(SESSION_COMPLETED, true);

        ra.addFlashAttribute("message", Mcm0021uConstants.MSG_REPLACE_COMPLETE);
        return "redirect:/mcm0021u?resume=true";
    }

    // =================================================================
    // 契約チェックAPI（Ajax用）
    // 【変換元】ReplaceButton_Click() 内の契約チェック部分
    //   元コード: getTkKeiyakuList() / getUkKeiyakuList() → IndexOfで交差判定
    //   JavaではAjaxリクエストで事前チェックし、結果をJavaScript側でconfirm表示
    // =================================================================
    /**
     * 契約不整合チェック（Ajax API）
     * @return true=不整合あり（警告確認が必要）
     */
    @PostMapping("/checkKeiyaku")
    @ResponseBody
    public boolean checkKeiyaku(@RequestBody List<MotoRowForm> motoRows) {
        return service.checkKeiyakuConflict(motoRows);
    }

    // =================================================================
    // ユーティリティ
    // =================================================================

    /**
     * セッション内の付替え元/先データをModelに追加する
     */
    @SuppressWarnings("unchecked")
    private void addSessionDataToModel(Model model, HttpSession session) {
        var form = (Mcm0021uForm) model.getAttribute("form");
        if (form != null) {
            form.setPlantIdMoto((BigDecimal) session.getAttribute(SESSION_PLANT_ID_MOTO));
            form.setPlantIdSaki((BigDecimal) session.getAttribute(SESSION_PLANT_ID_SAKI));
        }
        var state = searchState(session);
        if (state != null) {
            if (form != null) {
                form.setNonyusakiCd(state.criteria.getNonyusakiCd()); form.setNonyusakiNk(state.criteria.getNonyusakiNk());
                form.setSupportId(state.criteria.getSupportId()); form.setPlantNk(state.criteria.getPlantNk());
            }
            model.addAttribute("searched", true);
            model.addAttribute("nonyusakiList", state.plants);
            model.addAttribute("kikikoseiList", state.compositions);
            model.addAttribute("kikimeisaiList", state.details);
            model.addAttribute("searchToken", state.token);
            model.addAttribute("selectedPlantId", state.selectedPlant);
            model.addAttribute("selectedKikikoseiId", state.selectedComposition);
            model.addAttribute("selectedBrandkoseiId", state.selectedBrand);
        }
        model.addAttribute("plantIdMoto", session.getAttribute(SESSION_PLANT_ID_MOTO));
        model.addAttribute("plantIdSaki", session.getAttribute(SESSION_PLANT_ID_SAKI));
        model.addAttribute("motoRows", withSelections((List<Mcm0021uMotoRowDto>) session.getAttribute(SESSION_MOTO_ROWS), form));
        boolean completed = Boolean.TRUE.equals(session.getAttribute(SESSION_COMPLETED));
        if (completed) {
            // 元のブランド名を持つ古いDTOを再利用せず、更新後の関連IDから名称を取得する。
            session.setAttribute(SESSION_SAKI_ROWS, service.findSakiByPlantId(
                    (BigDecimal) session.getAttribute(SESSION_PLANT_ID_SAKI)));
        }
        model.addAttribute("replaceCompleted", completed);
        model.addAttribute("sakiRows", session.getAttribute(SESSION_SAKI_ROWS));
        model.addAttribute("brandCombo", session.getAttribute(SESSION_BRAND_COMBO));
        model.addAttribute("sakiNonyusaki", session.getAttribute(SESSION_SAKI_NONYUSAKI));
    }

    /**
     * 文字列が空またはnullかチェック（全角空白対応）
     */
    private static boolean isEmpty(String s) {
        return s == null || s.strip().isEmpty();
    }

    private static List<Mcm0021uMotoRowDto> withSelections(List<Mcm0021uMotoRowDto> source, Mcm0021uForm form) {
        if (source == null) return List.of();
        if (form == null || form.getMotoRows() == null || form.getMotoRows().isEmpty()) return source;
        return source.stream().map(row -> {
            var copy = new Mcm0021uMotoRowDto();
            org.springframework.beans.BeanUtils.copyProperties(row, copy);
            form.getMotoRows().stream().filter(r -> r != null
                    && sameId(r.getKikikoseiId(), row.getKikikoseiId())
                    && sameId(r.getBrandkoseiId(), row.getBrandkoseiId())).findFirst().ifPresent(r -> {
                copy.setIdoCheck(r.isIdoCheck() ? 1 : 0); copy.setIkosaki(r.getIkosaki());
            });
            return copy;
        }).toList();
    }

    @Override
    protected String getScreenTitle() {
        return Mcm0021uConstants.SCREEN_TITLE;
    }

    @Override
    protected String getFunctionId() {
        return Mcm0021uConstants.SCREEN_ID;
    }
}
