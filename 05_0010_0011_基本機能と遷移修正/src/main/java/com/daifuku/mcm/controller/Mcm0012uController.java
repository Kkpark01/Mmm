/**
 * 【変換元】Mcm0012uScreen.vb（611行）
 *   MCM0012U プラントマスタ画面コントローラ
 *   ボタンマッピング:
 *     UpdateButton_Click()          → POST /mcm0012u/update
 *     NONYUSAKITSUIKAButton_Click() → POST /mcm0012u/addEquipment
 *     RowDeleteButton_Click()       → POST /mcm0012u/deleteRow
 */
package com.daifuku.mcm.controller;

import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import jakarta.servlet.http.HttpSession;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import org.springframework.web.util.UriUtils;

import com.daifuku.mcm.common.AppConstants;
import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.common.Mcm0012uConstants;
import com.daifuku.mcm.dto.UserInfo;
import com.daifuku.mcm.form.Mcm0012uForm;
import com.daifuku.mcm.service.ComboBoxDataService;
import com.daifuku.mcm.service.Mcm0012uService;

@Controller
@RequestMapping("/mcm0012u")
public class Mcm0012uController extends BaseController {

    private static final Logger log = LoggerFactory.getLogger(Mcm0012uController.class);

    @Autowired
    private Mcm0012uService service;

	@Autowired
	private ComboBoxDataService comboBoxDataService;

    // TODO: ComboBoxDataServiceが実装済みの場合はコメント解除してください
    // @Autowired
    // private ComboBoxDataService comboBoxDataService;

    @Override
    protected String getScreenTitle() { return "プラントマスタ"; }

    @Override
    protected String getFunctionId() { return "MCM0012U"; }

    /** 0011U等からの戻り先を、入力エラー・DBエラーによる再表示でも保持する。 */
    @ModelAttribute
    public void preserveReturnUrl(@RequestParam(required = false) String returnUrl, Model model) {
        model.addAttribute("returnUrl", returnUrl != null ? returnUrl : "/mcm0010u");
    }

    /**
     * 【変換元】Mcm0012uScreen_Load() + Search()
     *   元コード: Me.DaoContainer.Fill(Me.Mcm0012uDataSet, "MCM_MA_PLANT")
     *            Me.DaoContainer.Fill(Me.Mcm0012uDataSet, "MCM_MA_BRAND_KOSEI")
     */
    @GetMapping
    public String index(
            @RequestParam(required = false) Long plantId,
            @RequestParam(required = false) Long nonyusakiId,
            @RequestParam(required = false) String nonyusakiCd,
            @RequestParam(required = false) String nonyusakiNk,
            @RequestParam(required = false) String returnUrl,
            @RequestParam(required = false, defaultValue = "1") int updateMode,
            Model model, HttpSession session) {

        log.info("MCM0012U 画面表示: plantId={}, nonyusakiId={}, updateMode={}", plantId, nonyusakiId, updateMode);
        // ★修正：session.getAttribute("userInfo") は誤ったセッションキーで、
        //   実際の格納キー（AppConstants.SESSION_USER_INFO="loginUserInfo"）と不一致のため
        //   常にnullとなっていた（MCM0016U/0018U/0024Uと同一パターンの不具合）。
        //   BaseController.getLoginUserInfo(session) を使用する。
        UserInfo userInfo = getLoginUserInfo(session);
        String authority = userInfo != null ? userInfo.getAuthorityDivision() : "";

        Mcm0012uForm form;
        if (plantId != null && updateMode == Mcm0012uConstants.UPDATE_MODE_UPDATE) {
            form = service.loadPlantData(plantId, nonyusakiCd, nonyusakiNk, nonyusakiId);
         // 修正後
        } else {
            form = new Mcm0012uForm();
            form.setUpdateMode(Mcm0012uConstants.UPDATE_MODE_INSERT);
            form.setNonyusakiCd(nonyusakiCd);
            form.setNonyusakiNk(nonyusakiNk);
            form.setNonyusakiId(nonyusakiId);

            // ★修正：デフォルト空行の手動追加を廃止。
            //   画面側(HTML)の末尾空行テンプレート(brandKoseiRows[bi], bi=size())が
            //   1件目の入力欄として機能するため、ここで別途1件追加すると
            //   brandKoseiRows中身が空のダミー行と画面の空行テンプレートが
            //   二重に生成され、JS初期化処理で[0]がDOMから消える。
            //   結果、POST時に[0]がSpringバインドで空のまま自動生成され、
            //   validate()の必須チェックに引っかかり「入力しているのに登録できない」原因となっていた。
            form.setBrandKoseiRows(new ArrayList<>());
        }

        model.addAttribute("form", form);
        setComboBoxData(model);
        model.addAttribute("isInspection", AppConstants.AUTHORITY_DIVISION_INSPECTION.equals(authority));
        model.addAttribute("screenTitle", getScreenTitle());
        model.addAttribute("returnUrl", returnUrl != null ? returnUrl : "/mcm0010u");
        // ★修正：nounyukubunListの設定はsetComboBoxData()に集約したため、ここでの個別追加は削除

        return "mcm0012u/index";
    }

    /**
     * 【変換元】UpdateButton_Click()
     *   元コード: CPValidateUtility.CheckValidate → DaoContainer.Update
     */
    @PostMapping("/update")
    public String update(@ModelAttribute("form") Mcm0012uForm form,
                         HttpSession session, RedirectAttributes ra, Model model,@RequestParam(required = false) String returnUrl) {
        log.info("MCM0012U 登録/更新: plantId={}, updateMode={}, plantSakujoFlg={}",
                form.getPlantId(), form.getUpdateMode(), form.isPlantSakujoFlg());
        // ★修正：誤ったセッションキー("userInfo")によりログインユーザーが取得できず、
        //   常に"system"にフォールバックしていた（MCM0016U/0018U/0024Uと同一パターン）。
        UserInfo userInfo = getLoginUserInfo(session);
        String loginUser = userInfo != null ? userInfo.getUserId() : "system";

        if (!form.isPlantSakujoFlg()) {
            List<String> errors = service.validateForSave(form);
            if (!errors.isEmpty()) {
                model.addAttribute("errors", errors);
                model.addAttribute("form", form);
                setComboBoxData(model);
                model.addAttribute("screenTitle", getScreenTitle());
                return "mcm0012u/index";
            }
        } else {
            // ★修正：ブランド構成情報が存在するプラントの削除は、確認メッセージ表示前ではなく
            //   登録ボタン押下時に「入力エラー」として扱う。
            //   （削除チェックボックス押下時の確認メッセージは表示したままにするため、
            //     こちらはサーバー側の登録処理でのみ判定する）
            List<String> brandExistsErrors = service.checkBrandKoseiExistsForPlantDelete(form);
            if (!brandExistsErrors.isEmpty()) {
                model.addAttribute("errors", brandExistsErrors);
                model.addAttribute("form", form);
                setComboBoxData(model);
                model.addAttribute("screenTitle", getScreenTitle());
                return "mcm0012u/index";
            }
        }

        try {
            if (form.isPlantSakujoFlg()) {
                List<String> deleteErrors = service.checkPlantDeletable(form.getPlantId());
                if (!deleteErrors.isEmpty()) {
                    ra.addFlashAttribute("errors", deleteErrors);
                    return buildRedirectUrl(form,returnUrl);
                }
                // ★追加：グリッドで削除マークされたブランド構成行を先にDBから削除する
                service.deleteMarkedBrandKoseiRows(form, loginUser);
                service.deletePlant(form.getPlantId(), loginUser);
                ra.addFlashAttribute("message", "プラントを削除しました");
                return buildRedirectUrlAfterPlantDelete(form, returnUrl);
            }

            Long savedPlantId = service.saveAll(form, loginUser);
            form.setPlantId(savedPlantId);
            form.setUpdateMode(Mcm0012uConstants.UPDATE_MODE_UPDATE);
            ra.addFlashAttribute("message", "登録を完了しました。");
        } catch (org.springframework.dao.DataIntegrityViolationException e) {
            // ★追加（#254）：DB制約違反（桁あふれ等）はユーザーが理解できる文言に変換
            log.error("MCM0012U 登録/更新エラー（データ整合性違反）", e);
            model.addAttribute("errors", List.of(
                "入力された数値が登録可能な桁数の上限を超えています。金額・台数などの項目をご確認のうえ、再度ご入力ください。"));
            model.addAttribute("form", form);
            setComboBoxData(model);
            model.addAttribute("screenTitle", getScreenTitle());
            return "mcm0012u/index";
        } catch (Exception e) {
            log.error("MCM0012U 登録/更新エラー", e);
            model.addAttribute("errors", List.of("登録処理でエラーが発生しました: " + e.getMessage()));
            model.addAttribute("form", form);
            setComboBoxData(model);
            model.addAttribute("screenTitle", getScreenTitle());
            return "mcm0012u/index";
        }
        return buildRedirectUrl(form, returnUrl);
    }

    /**
     * 行削除ボタン用：指定ブランド構成が削除可能かを判定する。
     *
     * ★追加：画面の「行削除」ボタン押下時点で使用中チェックを行うために追加。
     *   従来は「登録」ボタン押下時（update）にしかチェックしておらず、
     *   行削除ボタン押下時点ではエラーが表示されなかったため、MCM0015Uと
     *   同一方式（GET /deletable）に合わせて追加した。
     *   なお登録時のチェック（validateForSave内）は二重防御として残す。
     *
     * @param brandkoseiId 判定対象のブランド構成ID（新規行は空でも可）
     * @return deletable=false のとき message に理由を格納して返す
     */
    @GetMapping("/deletable")
    @ResponseBody
    public java.util.Map<String, Object> deletable(
            @RequestParam(name = "brandkoseiId", required = false) Long brandkoseiId) {
        java.util.Map<String, Object> body = new java.util.LinkedHashMap<>();
        if (brandkoseiId == null) {
            // 新規行（未保存）は削除可として扱う。
            body.put("deletable", true);
            return body;
        }
        List<String> errors = service.checkBrandKoseiDeletable(brandkoseiId);
        boolean ok = errors.isEmpty();
        body.put("deletable", ok);
        body.put("message", ok ? "" : errors.get(0));
        return body;
    }

    /**
     * 【変換元】RowDeleteButton_Click()
     *   元コード: If ConfirmMessage("MCM_MSG_DELETE") <> Yes Then Return
     */
    @PostMapping("/deleteRow")
    public String deleteRow(@RequestParam Long brandkoseiId,
                            @ModelAttribute("form") Mcm0012uForm form,
                            HttpSession session, RedirectAttributes ra,Model model,@RequestParam(required = false) String returnUrl) {
        log.info("MCM0012U 行削除: brandkoseiId={}", brandkoseiId);
        // ★修正：誤ったセッションキー("userInfo")によりログインユーザーが取得できず、
        //   常に"system"にフォールバックしていた（MCM0016U/0018U/0024Uと同一パターン）。
        UserInfo userInfo = getLoginUserInfo(session);
        String loginUser = userInfo != null ? userInfo.getUserId() : "system";

        try {
            // 【変換元】BeforeDeleteBrandKoseiCheck()
            List<String> errors = service.checkBrandKoseiDeletable(brandkoseiId);
            if (!errors.isEmpty()) {
                model.addAttribute("errors", errors);
                model.addAttribute("form", form);
                setComboBoxData(model);
                model.addAttribute("screenTitle", getScreenTitle());
                return "mcm0012u/index";
            }
            service.deleteBrandKosei(brandkoseiId, loginUser);
            ra.addFlashAttribute("message", "行を削除しました");
        } catch (Exception e) {
            log.error("MCM0012U 行削除エラー", e);
            model.addAttribute("errors", List.of("削除処理でエラーが発生しました: " + e.getMessage()));
            model.addAttribute("form", form);
            setComboBoxData(model);
            model.addAttribute("screenTitle", getScreenTitle());
            return "mcm0012u/index";
        }
        return buildRedirectUrl(form,returnUrl);
    }

    /**
     * 【変換元】NONYUSAKITSUIKAButton_Click()
     *   元コード: 保存後 → Mcm0013uDelivery で MCM0013U へ遷移
     */
    @PostMapping("/addEquipment")
    public String addEquipment(@ModelAttribute("form") Mcm0012uForm form,
                               HttpSession session, RedirectAttributes ra, Model model,
                               @RequestParam(required = false) String returnUrl) {
        log.info("MCM0012U 納入機器追加: plantId={}", form.getPlantId());
        // ★修正：誤ったセッションキー("userInfo")によりログインユーザーが取得できず、
        //   常に"system"にフォールバックしていた（MCM0016U/0018U/0024Uと同一パターン）。
        UserInfo userInfo = getLoginUserInfo(session);
        String loginUser = userInfo != null ? userInfo.getUserId() : "system";

        List<String> errors = service.validateForAddEquipment(form);
        if (!errors.isEmpty()) {
            model.addAttribute("errors", errors);
            model.addAttribute("form", form);
            model.addAttribute("returnUrl", returnUrl);
            setComboBoxData(model);
            model.addAttribute("screenTitle", getScreenTitle());
            return "mcm0012u/index";
        }

        try {
            Long plantId = service.saveAll(form, loginUser);

            // ★ 0013Uの「戻る」で0012Uに戻れるURL（redirect:は不要）
            String backTo0012 = "/mcm0012u?plantId=" + plantId
                    + "&nonyusakiId=" + form.getNonyusakiId()
                    + "&nonyusakiCd=" + encodeParam(form.getNonyusakiCd())
                    + "&nonyusakiNk=" + encodeParam(form.getNonyusakiNk())
                    + (returnUrl != null && !returnUrl.isBlank()
                        ? "&returnUrl=" + encodeParam(returnUrl)
                        : "");

            return "redirect:/mcm0013u?plantId=" + plantId
                    + "&nonyusakiId=" + form.getNonyusakiId()
                    + "&nonyusakiCd=" + encodeParam(form.getNonyusakiCd())
                    + "&nonyusakiNk=" + encodeParam(form.getNonyusakiNk())
                    + "&returnTo0012Url=" + encodeParam(backTo0012);

        } catch (Exception e) {
            log.error("MCM0012U 保存→遷移エラー", e);
            model.addAttribute("errors", List.of("保存処理でエラーが発生しました: " + e.getMessage()));
            model.addAttribute("form", form);
            setComboBoxData(model);
            model.addAttribute("screenTitle", getScreenTitle());
            return "mcm0012u/index";
        }
    }

    // ========================================
    // Private Methods
    // ========================================

    /**
     * コンボボックスデータをModelに設定
     * 【変換元】Mcm0012uScreen_Load() - ComboBox.DataSource / Items 設定
     *
     * TODO: ComboBoxDataServiceが実装されたら、そちらから取得するように変更してください。
     *       現在は仮のデータ（空リスト）を設定しています。
     */
    private void setComboBoxData(Model model) {
        // ★修正：仮実装（ブランド1/ブランド2の2件固定）を廃止し、
        //   ComboBoxDataService.getBrandComboList()でブランドマスタ（MCM_MA_BRAND）から取得する。
        //   無効フラグ（YUKOU_FLG）がONのブランドは表示名の先頭に「X 」を付与し、
        //   有効なブランドの後（リスト下部）に表示される。
        List<Map<String, String>> brandList = comboBoxDataService.getBrandComboList();

        // ★修正：仮実装（8/24の2件固定）をComboBoxDataServiceの正式メソッドに置き換え。
        //   ComboBoxDataServiceの既存getHosyukeiyakuJikantaiList()（保守契約時間帯：8/24）は
        //   別項目用のため流用不可。稼働時間専用のgetKadoJikanList()（1～24）を使用する。
        List<Map<String, String>> kadojikanList = comboBoxDataService.getKadoJikanList();

        // ★修正：HTML側の参照名「kadoNissuList」に合わせて属性名を是正
        //         あわせて期待値（1～7）に基づきデータを実装
        List<Map<String, String>> kadoNissuList = new ArrayList<>();
        for (int i = 1; i <= 7; i++) {
            kadoNissuList.add(Map.of("key", String.valueOf(i), "value", String.valueOf(i)));
        }

        // VB McmSystemkadoyoubiDataTable_Item.xml: 1=日、2=月〜7=土。0010の表示とも一致させる。
        List<Map<String, String>> kadoYoubiList = new ArrayList<>();
        kadoYoubiList.add(Map.of("key", "2", "value", "月"));
        kadoYoubiList.add(Map.of("key", "3", "value", "火"));
        kadoYoubiList.add(Map.of("key", "4", "value", "水"));
        kadoYoubiList.add(Map.of("key", "5", "value", "木"));
        kadoYoubiList.add(Map.of("key", "6", "value", "金"));
        kadoYoubiList.add(Map.of("key", "7", "value", "土"));
        kadoYoubiList.add(Map.of("key", "1", "value", "日"));

        model.addAttribute("brandList", brandList);
        model.addAttribute("kadojikanList", kadojikanList);
        model.addAttribute("kadoNissuList", kadoNissuList);
        model.addAttribute("kadoYoubiList", kadoYoubiList);

        // ★修正：index()にのみ設定されていた納入区分リストをこちらに集約。
        //   これによりupdate()/deleteRow()/addEquipment()のバリデーションエラー時の
        //   再表示でも選択肢が消えなくなる（リグレッション防止）。
        model.addAttribute("nounyukubunList", comboBoxDataService.getNounyukubunList());

        // ※不要になった旧「nounyuKbnList」（HTML未参照のダミーデータ）は削除
    }

    /** PRGパターン用リダイレクトURL構築 */
    private String buildRedirectUrl(Mcm0012uForm form,String returnUrl) {
        StringBuilder sb = new StringBuilder("redirect:/mcm0012u?");
        if (form.getPlantId() != null) {
            sb.append("plantId=").append(form.getPlantId()).append("&");
        }
        if (form.getNonyusakiId() != null) {
            sb.append("nonyusakiId=").append(form.getNonyusakiId()).append("&");
        }
        sb.append("nonyusakiCd=").append(encodeParam(form.getNonyusakiCd())).append("&");
        sb.append("nonyusakiNk=").append(encodeParam(form.getNonyusakiNk())).append("&");
        sb.append("updateMode=").append(form.getUpdateMode());

        if (returnUrl != null && !returnUrl.isBlank()) {
            sb.append("&returnUrl=").append(encodeParam(returnUrl));
        }

        return sb.toString();
    }
    
    private String buildRedirectUrlAfterPlantDelete(Mcm0012uForm form, String returnUrl) {
        StringBuilder sb = new StringBuilder("redirect:/mcm0012u?");
        if (form.getNonyusakiId() != null) {
            sb.append("nonyusakiId=").append(form.getNonyusakiId()).append("&");
        }
        sb.append("nonyusakiCd=").append(encodeParam(form.getNonyusakiCd())).append("&");
        sb.append("nonyusakiNk=").append(encodeParam(form.getNonyusakiNk())).append("&");
        sb.append("updateMode=").append(Mcm0012uConstants.UPDATE_MODE_INSERT);
        if (returnUrl != null && !returnUrl.isBlank()) {
            sb.append("&returnUrl=").append(encodeParam(returnUrl));
        }
        return sb.toString();
    }

    /** URLパラメータのエンコード */
    private String encodeParam(String value) {
    	 return UriUtils.encode(value == null ? "" : value, StandardCharsets.UTF_8);
    }
}
