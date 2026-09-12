package com.daifuku.mcm.controller;

import com.daifuku.mcm.constants.Mcm2001uConstants;
import com.daifuku.mcm.dto.Mcm2001uDeliveryDto;
import com.daifuku.mcm.dto.Mcm2002uDeliveryDto;
import com.daifuku.mcm.dto.Mcm2001uRowDto;
import com.daifuku.mcm.form.Mcm2001uForm;
import com.daifuku.mcm.service.Mcm2001uService;
import com.daifuku.mcm.service.Mcm2001uService.SearchResult;
import com.daifuku.mcm.service.Mcm2001uService.SenteiResult;
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

/**
 * 【変換元】Mcm2001uScreen.vb
 *   MCM2001U カスタマー見積検索 コントローラ
 *   VB.NET → Java変換
 *
 *   画面タイトル: 保守見積作成検索
 *   3グリッド構成（納入先 / プラント / 取引先見積）
 *   ボタン2個（検索 / 選定）
 *   リンク2種（見積依頼NO→MCM1004U / 契約NO→MCM1005U）
 *
 *   変換パターン: C（検索参照） + F（画面間連携）
 */
@Controller
@RequestMapping("/mcm2001u")
public class Mcm2001uController extends com.daifuku.mcm.common.BaseController {
    @org.springframework.beans.factory.annotation.Autowired private com.daifuku.mcm.service.Mcm2004uService permissions;

    private static final Logger log = LoggerFactory.getLogger(Mcm2001uController.class);

    private final Mcm2001uService service;

    public Mcm2001uController(Mcm2001uService service) {
        this.service = service;
    }

    // ================================================================
    //  初期表示
    // ================================================================

    /**
     * 【変換元】Mcm2001uScreen.vb - Mcm2001uScreen_Load()
     *   初期表示
     *
     *   元コード:
     *     Dim delivery = MyBase.GetDeliveryData
     *     If IsNotNull(delivery) Then
     *       Me.search()
     *       検索条件をReadOnly化
     *     End If
     *
     * @param model Thymeleafモデル
     * @param session HTTPセッション
     * @return テンプレートパス
     */
    @GetMapping
    public String index(Model model, HttpSession session) {

        // delivery経由のアクセスかチェック
        Mcm2001uDeliveryDto delivery = (Mcm2001uDeliveryDto)
                session.getAttribute(Mcm2001uConstants.SESSION_KEY_DELIVERY);

        if (delivery != null) {
            /*
             * 【変換元】L18-28
             *   元コード: Me.search()
             *             Me.NONYUSAKI_CDTextBox.ReadOnly = True ...
             *             Me.SEARCHButton.Enabled = False
             */
            SearchResult result = service.searchByDelivery(delivery);
            com.daifuku.mcm.common.CustomerScreenSupport.rotate(session,"mcm2001u");
            Mcm2001uForm form = new Mcm2001uForm();
            form.setNonyusakiCd(delivery.getNonyusakiCd());form.setNonyusakiNk(delivery.getNonyusakiNk());
            form.setSupportId(delivery.getSupportId());form.setPlantNk(delivery.getPlantNk());
            session.setAttribute(Mcm2001uConstants.SESSION_KEY_SEARCH_RESULT, result);
            session.setAttribute(Mcm2001uConstants.SESSION_KEY_FORM, form);
            model.addAttribute("form", form);
            model.addAttribute("nonyusakiList", result.getNonyusakiList());
            model.addAttribute("plantList", result.getPlantList());
            model.addAttribute("mitsumoriList", result.getMitsumoriList());
            model.addAttribute("readOnly", true);
            model.addAttribute("hasDelivery", true);
        } else {
            // セッションに検索結果があれば復元（PRGパターン）
            @SuppressWarnings("unchecked")
            SearchResult savedResult = (SearchResult)
                    session.getAttribute(Mcm2001uConstants.SESSION_KEY_SEARCH_RESULT);
            Mcm2001uForm savedForm = (Mcm2001uForm)
                    session.getAttribute(Mcm2001uConstants.SESSION_KEY_FORM);

            model.addAttribute("form", savedForm != null ? savedForm : new Mcm2001uForm());
            if (savedResult != null) {
                model.addAttribute("nonyusakiList", savedResult.getNonyusakiList());
                model.addAttribute("plantList", savedResult.getPlantList());
                model.addAttribute("mitsumoriList", savedResult.getMitsumoriList());
            }
            model.addAttribute("readOnly", false);
            model.addAttribute("hasDelivery", false);
        }

        model.addAttribute("selectedNonyusakiId",session.getAttribute("mcm2001u.selectedNonyusakiId"));
        model.addAttribute("selectedPlantId",session.getAttribute("mcm2001u.selectedPlantId"));
        model.addAttribute("selectedTmIds",session.getAttribute("mcm2001u.selectedTmIds"));
        model.addAttribute("workflowToken",com.daifuku.mcm.common.CustomerScreenSupport.token(session,"mcm2001u"));
        return "mcm2001u/index";
    }

    // ================================================================
    //  検索
    // ================================================================

    /**
     * 【変換元】Mcm2001uScreen.vb - SEARCHButton_Click()
     *   検索ボタン押下
     *
     *   元コード:
     *     Me.Fill(MCM_MA_NONYUSAKI, params)
     *     Me.Fill(MCM_MA_PLANT, params)
     *     Me.Fill(MCM_TM_MITSUMORI, params)
     *
     * @param form 検索フォーム
     * @param ra リダイレクト属性
     * @param session HTTPセッション
     * @return リダイレクト先
     */
    @PostMapping("/search")
    public String search(@ModelAttribute Mcm2001uForm form,
                          RedirectAttributes ra,
                          HttpSession session) {

        SearchResult result = service.searchByForm(form);

        if (result.hasErrors()) {
            ra.addFlashAttribute("errors", result.getErrors());
        }

        com.daifuku.mcm.common.CustomerScreenSupport.rotate(session,"mcm2001u");
        // PRGパターン: セッションに検索結果を保持
        session.setAttribute(Mcm2001uConstants.SESSION_KEY_FORM, form);
        session.setAttribute(Mcm2001uConstants.SESSION_KEY_SEARCH_RESULT, result);

        return "redirect:/mcm2001u";
    }

    // ================================================================
    //  選定
    // ================================================================

    /**
     * 【変換元】Mcm2001uScreen.vb - SENTEIButton_Click()
     *   選定ボタン押下 → MCM2002Uへ遷移
     *
     *   元コード:
     *     チェック行収集 → 重複チェック → 納入先/プラント情報取得
     *     → ForwardScreen(MCM2002U, delivery)
     *
     * @param selectedNonyusakiId 選択中の納入先ID
     * @param selectedPlantId 選択中のプラントID
     * @param selectedIds チェックした見積の契約時間ID（検索結果との照合対象）
     * @param ra リダイレクト属性
     * @param session HTTPセッション
     * @return リダイレクト先
     */
    @PostMapping("/sentei")
    public String sentei(@RequestParam(required=false) String workflowToken,@RequestParam(required = false) BigDecimal selectedNonyusakiId,
                          @RequestParam(required = false) BigDecimal selectedPlantId,
                          @RequestParam(value="selectedTmKeiyakujikanId", required=false) List<BigDecimal> selectedIds,
                          RedirectAttributes ra,
                          HttpSession session) {

        if(!com.daifuku.mcm.common.CustomerScreenSupport.token(session,"mcm2001u").equals(workflowToken)){ra.addFlashAttribute("error","画面の情報が更新されています。再検索してください。");return "redirect:/mcm2001u";}
        // セッションからデータ復元
        @SuppressWarnings("unchecked")
        SearchResult searchResult = (SearchResult)
                session.getAttribute(Mcm2001uConstants.SESSION_KEY_SEARCH_RESULT);
        Mcm2001uDeliveryDto delivery = (Mcm2001uDeliveryDto)
                session.getAttribute(Mcm2001uConstants.SESSION_KEY_DELIVERY);

        List<Map<String, Object>> nonyusakiList =
                searchResult != null ? searchResult.getNonyusakiList() : null;
        List<Map<String, Object>> plantList =
                searchResult != null ? searchResult.getPlantList() : null;

        // 選択IDをサーバー側の検索結果と照合し、別の納入先・プラントを混在させない。
        if (searchResult == null || !containsId(nonyusakiList,"NONYUSAKI_ID",selectedNonyusakiId)
                || !containsPlant(plantList,selectedPlantId,selectedNonyusakiId)) {
            ra.addFlashAttribute("errors",List.of("納入先・プラントを選択し直してください。"));
            return "redirect:/mcm2001u";
        }
        List<Mcm2001uRowDto> rows = new java.util.ArrayList<>();
        for (BigDecimal id : selectedIds == null ? List.<BigDecimal>of() : selectedIds) {
            Mcm2001uRowDto source = searchResult.getMitsumoriList().stream()
                .filter(r -> same(id,r.getTmKeiyakujikanId()) && same(selectedPlantId,r.getPlantId())
                    && same(selectedNonyusakiId,r.getNonyusakiId())).findFirst().orElse(null);
            if (source == null) {
                ra.addFlashAttribute("errors",List.of("選択した見積を確認できません。再検索してください。"));
                return "redirect:/mcm2001u";
            }
            if (rows.stream().noneMatch(r -> same(id,r.getTmKeiyakujikanId()))) {
                Mcm2001uRowDto row = new Mcm2001uRowDto();org.springframework.beans.BeanUtils.copyProperties(source,row);
                row.setCheckbox(1);rows.add(row);
            }
        }

        // 選定処理実行
        if(!com.daifuku.mcm.common.CustomerScreenSupport.update(session,"MCM2003U",hasUpdateAuthority(),()->permissions.getAuthority(getLoginUserId(),"MCM2003U"))){ra.addFlashAttribute("error","権限がないため遷移できません。");return "redirect:/mcm2001u";}
        SenteiResult result = service.sentei(
                rows, selectedNonyusakiId, selectedPlantId,
                nonyusakiList, plantList, delivery);

        if (!result.isSuccess()) {
            ra.addFlashAttribute("errors", List.of(result.getErrorMessage()));
            return "redirect:/mcm2001u";
        }

        session.setAttribute("mcm2001u.selectedNonyusakiId",selectedNonyusakiId);
        session.setAttribute("mcm2001u.selectedPlantId",selectedPlantId);
        session.setAttribute("mcm2001u.selectedTmIds",selectedIds==null?List.of():selectedIds);
        com.daifuku.mcm.common.CustomerScreenSupport.rotate(session,"mcm2001u");
        Mcm2002uDeliveryDto next = new Mcm2002uDeliveryDto();
        next.setNonyusakiId(BigDecimal.valueOf(result.getNonyusakiId()));next.setNonyusakiCd(result.getNonyusakiCd());
        next.setNonyusakiNk(result.getNonyusakiNk());next.setPlantId(BigDecimal.valueOf(result.getPlantId()));
        next.setPlantNk(result.getPlantNk());next.setSupportId(result.getSupportId());
        next.setTmKeiyakujikanId(result.getTmKeiyakujikanId());next.setSeniMotoKbn(result.getSeniMotoKbn());
        next.setUmKihonMitsumoriId(result.getUmKihonMitsumoriId()>0?BigDecimal.valueOf(result.getUmKihonMitsumoriId()):null);
        for(String key:java.util.Collections.list(session.getAttributeNames()))if(key.startsWith("mcm2002u."))session.removeAttribute(key);
        session.setAttribute("mcm2002u.delivery",next);
        for(String key:java.util.Collections.list(session.getAttributeNames()))if(key.startsWith("mcm2003u."))session.removeAttribute(key);
        session.removeAttribute("MCM2003U_FORM");

        log.info("MCM2001U → MCM2002U 遷移: 納入先ID={}, プラントID={}",
                result.getNonyusakiId(), result.getPlantId());

        return "redirect:/mcm2002u";
    }

    // ================================================================
    //  リンク遷移
    // ================================================================

    /**
     * 【変換元】Mcm2001uScreen.vb - MCM_TM_MITSUMORIDataGridView_CellContentClick()
     *   見積依頼NOリンク → MCM1004Uへ遷移
     *
     *   元コード:
     *     If dgv.Columns(e.ColumnIndex).Name = TM_IRAI_NO_MITUMORI Then
     *       newDelMcm1004u.TmKeiyakujikanID = TM_KEIYAKUJIKAN_ID
     *       newDelMcm1004u.SeniMotoKbn = SENIMOTO_TENPO
     *       ForwardScreen(MCM1004U, newDelMcm1004u)
     *
     * @param tmKeiyakujikanId 取引先見積契約時間ID
     * @param session HTTPセッション
     * @return リダイレクト先
     */
    @GetMapping("/link/mcm1004u")
    public String linkToMcm1004u(@RequestParam BigDecimal tmKeiyakujikanId,HttpSession session) {
        SearchResult result=(SearchResult)session.getAttribute(Mcm2001uConstants.SESSION_KEY_SEARCH_RESULT);
        if(result==null||result.getMitsumoriList().stream().noneMatch(r->same(r.getTmKeiyakujikanId(),tmKeiyakujikanId)))return "redirect:/mcm2001u";
        return "redirect:/mcm1004u?tmKeiyakujikanId="+tmKeiyakujikanId.toPlainString()+"&seniMotoKbn=2";
    }

    /**
     * 【変換元】Mcm2001uScreen.vb - MCM_TM_MITSUMORIDataGridView_CellContentClick()
     *   契約NOリンク → MCM1005Uへ遷移
     *
     *   元コード:
     *     ElseIf dgv.Columns(e.ColumnIndex).Name = TKA_KEIYAKU_NO_MITUMORI Then
     *       newDelMcm1005u.KeiyakuId = TKA_TK_KEIYAKU_ID
     *       newDelMcm1005u.SeniMotoKbn = SENIMOTO_TENPO
     *       ForwardScreen(MCM1005U, newDelMcm1005u)
     *
     * @param keiyakuId 取引先契約ID
     * @param session HTTPセッション
     * @return リダイレクト先
     */
    @GetMapping("/link/mcm1005u")
    public String linkToMcm1005u(@RequestParam BigDecimal keiyakuId,HttpSession session) {
        SearchResult result=(SearchResult)session.getAttribute(Mcm2001uConstants.SESSION_KEY_SEARCH_RESULT);
        if(result==null||result.getMitsumoriList().stream().noneMatch(r->same(r.getTkaTkKeiyakuId(),keiyakuId)))return "redirect:/mcm2001u";
        return "redirect:/mcm1005u?tkKeiyakuId="+keiyakuId.toPlainString();
    }
    private static boolean same(BigDecimal a,BigDecimal b){return a!=null&&b!=null&&a.compareTo(b)==0;}
    @ExceptionHandler(org.springframework.dao.DataAccessException.class)
    public String dataAccessFailure(org.springframework.dao.DataAccessException e, RedirectAttributes ra) {
        log.error("MCM2001U data access failed",e);
        ra.addFlashAttribute("errors",List.of("画面の情報を取得できませんでした。再検索してから、もう一度操作してください。"));
        return "redirect:/mcm2004u";
    }
    private static boolean containsId(List<Map<String,Object>> rows,String key,BigDecimal id) {
        return id!=null&&rows!=null&&rows.stream().anyMatch(r->r.get(key)!=null&&same(id,new BigDecimal(r.get(key).toString())));
    }
    private static boolean containsPlant(List<Map<String,Object>> rows,BigDecimal plant,BigDecimal nonyu) {
        return rows!=null&&rows.stream().anyMatch(r->containsId(List.of(r),"PLANT_ID",plant)&&containsId(List.of(r),"NONYUSAKI_ID",nonyu));
    }
    @Override protected String getScreenTitle(){return "保守見積作成検索";}
    @Override protected String getFunctionId(){return "MCM2001U";}
}
