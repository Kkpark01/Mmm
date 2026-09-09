package com.daifuku.mcm.service;

import com.daifuku.mcm.dto.Mcm2001uDeliveryDto;
import com.daifuku.mcm.dto.Mcm2001uRowDto;
import com.daifuku.mcm.form.Mcm2001uForm;
import com.daifuku.mcm.repository.Mcm2001uJdbcRepository;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * 【変換元】Mcm2001uScreen.vb
 *   MCM2001U カスタマー見積検索 サービス
 *   VB.NET → Java変換
 *
 *   画面ロジック:
 *     Screen_Load → searchByDelivery()
 *     SEARCHButton_Click → searchByForm()
 *     SENTEIButton_Click → sentei()
 */
@Service
public class Mcm2001uService {

    private static final Logger log = LoggerFactory.getLogger(Mcm2001uService.class);

    private final Mcm2001uJdbcRepository jdbcRepository;

    public Mcm2001uService(Mcm2001uJdbcRepository jdbcRepository) {
        this.jdbcRepository = jdbcRepository;
    }

    // ================================================================
    //  検索結果格納用の内部クラス
    // ================================================================

    /**
     * 3グリッド分の検索結果を一括で返すための結果クラス
     *
     * 【変換元】VB.NETでは3つのDataTableをDataSet内に保持していたが、
     *   Java版では1つの結果オブジェクトに集約
     */
    public static class SearchResult {
        private List<Map<String, Object>> nonyusakiList = Collections.emptyList();
        private List<Map<String, Object>> plantList = Collections.emptyList();
        private List<Mcm2001uRowDto> mitsumoriList = Collections.emptyList();
        private List<String> errors = new ArrayList<>();

        public List<Map<String, Object>> getNonyusakiList() { return nonyusakiList; }
        public void setNonyusakiList(List<Map<String, Object>> v) { this.nonyusakiList = v; }
        public List<Map<String, Object>> getPlantList() { return plantList; }
        public void setPlantList(List<Map<String, Object>> v) { this.plantList = v; }
        public List<Mcm2001uRowDto> getMitsumoriList() { return mitsumoriList; }
        public void setMitsumoriList(List<Mcm2001uRowDto> v) { this.mitsumoriList = v; }
        public List<String> getErrors() { return errors; }
        public void addError(String msg) { this.errors.add(msg); }
        public boolean hasErrors() { return !errors.isEmpty(); }
    }

    // ================================================================
    //  Delivery経由の検索（Screen_Load時）
    // ================================================================

    /**
     * 【変換元】Mcm2001uScreen.vb - search()
     *   画面間遷移（MCM2003U等）からのdelivery経由検索
     *
     *   元コード:
     *     Dim tmKeiyakujikanID = Me.getUmKeiyakujikanId(delivery.UmKihonMitsumoriID)
     *     Me.MCM_MA_NONYUSAKITableAdapter.Fill(... delivery.NonyusakiID)
     *     Me.MCM_MA_PLANTTableAdapter.Fill(... delivery.PlantId)
     *     Me.MCM_TM_MITSUMORITableAdapter.Fill(... delivery.PlantId)
     *     If IsNotNull(tmKeiyakujikanID) Then
     *       For i = 0 To tmKeiyakujikanID.Length - 1
     *         mitsumoriRow.CHECKBOX = CHECKBOX_FLAG_ON
     *
     * @param delivery 呼び出し元画面からの受渡しデータ
     * @return 3グリッド分の検索結果（チェック済み）
     */
    public SearchResult searchByDelivery(Mcm2001uDeliveryDto delivery) {
        SearchResult result = new SearchResult();

        // 店舗基本見積IDから取引先見積契約時間IDを取得
        List<String> tmKeiyakujikanIds = Collections.emptyList();
        if (delivery.getUmKihonMitsumoriId() > 0) {
            tmKeiyakujikanIds = jdbcRepository.findUmKeiyakujikanIds(
                    delivery.getUmKihonMitsumoriId());
        }

        // プラントID指定で見積検索
        List<Mcm2001uRowDto> mitsumoriRows = jdbcRepository.searchMitsumoriByPlantId(
                BigDecimal.valueOf(delivery.getPlantId()));

        /*
         * 【変換元】Mcm2001uScreen.vb L56-65
         *   元コード: For i = 0 To tmKeiyakujikanID.Length - 1
         *             Dim mitsumoriRowList = Me.Mcm2001uDataSet.MCM_TM_MITSUMORI.Select(
         *               "TM_KEIYAKUJIKAN_ID=" & tmKeiyakujikanID(i))
         *             For Each mitsumoriRow ... mitsumoriRow.CHECKBOX = CHECKBOX_FLAG_ON
         */
        // チェックボックスON処理
        if (!tmKeiyakujikanIds.isEmpty() && mitsumoriRows != null) {
            for (Mcm2001uRowDto row : mitsumoriRows) {
                if (row.getTmKeiyakujikanId() != null
                        && tmKeiyakujikanIds.contains(
                                row.getTmKeiyakujikanId().toPlainString())) {
                    row.setCheckbox(1); // CHECKBOX_FLAG_ON
                }
            }
        }

        result.setMitsumoriList(mitsumoriRows != null ? mitsumoriRows : Collections.emptyList());

        result.setNonyusakiList(jdbcRepository.findNonyusakiById(BigDecimal.valueOf(delivery.getNonyusakiId())));
        result.setPlantList(jdbcRepository.findPlantById(BigDecimal.valueOf(delivery.getPlantId())));

        return result;
    }

    // ================================================================
    //  フォーム検索（検索ボタン押下時）
    // ================================================================

    /**
     * 【変換元】Mcm2001uScreen.vb - SEARCHButton_Click()
     *   検索ボタン押下時の処理
     *
     *   元コード:
     *     If IsNull(NONYUSAKI_CDTextBox.Text) And ... Then
     *       DisplayMessage(MSG_0001)
     *       Return
     *     End If
     *     Me.Fill(MCM_MA_NONYUSAKI, params)
     *     If nonyuCount = 0 Then DisplayMessage(MSG_0002)
     *     Me.Fill(MCM_MA_PLANT, params)
     *     Me.Fill(MCM_TM_MITSUMORI, params)
     *     リレーション設定
     *
     * @param form 検索フォーム（5条件）
     * @return 3グリッド分の検索結果
     */
    public SearchResult searchByForm(Mcm2001uForm form) {
        SearchResult result = new SearchResult();

        /*
         * 【変換元】L90-94
         *   元コード: If IsNull(NONYUSAKI_CDTextBox.Text) And ... Then
         *             MyBase.DisplayMessage(CPMessageConstant.MSG_0001)
         */
        // 全条件空白チェック
        if (form.isEmpty()) {
            result.addError("検索条件を入力してください。");
            return result;
        }

        // 納入先マスタ検索
        List<Map<String, Object>> nonyusakiList = jdbcRepository.searchNonyusaki(form);

        /*
         * 【変換元】L102-105
         *   元コード: If nonyuCount = 0 Then
         *             MyBase.DisplayMessage(CPMessageConstant.MSG_0002)
         */
        if (nonyusakiList.isEmpty()) {
            result.addError("該当するデータが見つかりません。");
            return result;
        }
        result.setNonyusakiList(nonyusakiList);

        // プラントマスタ検索
        List<Map<String, Object>> plantList = jdbcRepository.searchPlant(form);
        result.setPlantList(plantList);

        // 取引先見積検索
        List<Mcm2001uRowDto> mitsumoriList = jdbcRepository.searchMitsumori(form);
        result.setMitsumoriList(mitsumoriList);

        /*
         * 【変換元】L113-120 リレーション設定
         *   VB.NETではDataSetのリレーションで親子孫を連動表示していたが、
         *   Java(Web)版ではHTML上でJavaScript/Thymeleafで親子孫連動を実現
         *   → サーバ側のリレーション設定は不要
         */

        log.info("MCM2001U 検索完了: 納入先={}件, プラント={}件, 見積={}件",
                nonyusakiList.size(), plantList.size(), mitsumoriList.size());

        return result;
    }

    // ================================================================
    //  選定処理（選定ボタン押下時）
    // ================================================================

    /**
     * 選定処理の結果を格納するクラス
     */
    public static class SenteiResult {
        private boolean success;
        private String errorMessage;
        // MCM2002Uへの遷移パラメータ
        private int nonyusakiId;
        private String nonyusakiCd;
        private String nonyusakiNk;
        private int plantId;
        private String plantNk;
        private String supportId;
        private String[] tmKeiyakujikanId;
        private int umKihonMitsumoriId;
        private int seniMotoKbn;

        public boolean isSuccess() { return success; }
        public void setSuccess(boolean v) { this.success = v; }
        public String getErrorMessage() { return errorMessage; }
        public void setErrorMessage(String v) { this.errorMessage = v; }
        public int getNonyusakiId() { return nonyusakiId; }
        public void setNonyusakiId(int v) { this.nonyusakiId = v; }
        public String getNonyusakiCd() { return nonyusakiCd; }
        public void setNonyusakiCd(String v) { this.nonyusakiCd = v; }
        public String getNonyusakiNk() { return nonyusakiNk; }
        public void setNonyusakiNk(String v) { this.nonyusakiNk = v; }
        public int getPlantId() { return plantId; }
        public void setPlantId(int v) { this.plantId = v; }
        public String getPlantNk() { return plantNk; }
        public void setPlantNk(String v) { this.plantNk = v; }
        public String getSupportId() { return supportId; }
        public void setSupportId(String v) { this.supportId = v; }
        public String[] getTmKeiyakujikanId() { return tmKeiyakujikanId; }
        public void setTmKeiyakujikanId(String[] v) { this.tmKeiyakujikanId = v; }
        public int getUmKihonMitsumoriId() { return umKihonMitsumoriId; }
        public void setUmKihonMitsumoriId(int v) { this.umKihonMitsumoriId = v; }
        public int getSeniMotoKbn() { return seniMotoKbn; }
        public void setSeniMotoKbn(int v) { this.seniMotoKbn = v; }
    }

    /**
     * 【変換元】Mcm2001uScreen.vb - SENTEIButton_Click()
     *   選定ボタン押下時の処理
     *
     *   元コード:
     *     1. 納入先テーブルにデータが存在しない場合 → MSG_0067
     *     2. チェック行を収集 → keiyakujikanId配列
     *     3. 重複チェック → MSG_0056
     *     4. 納入先/プラント情報 → MCM2002Uへのdelivery構築
     *     5. ForwardScreen(Mcm2002uScreen, delivery)
     *
     * @param rows           見積グリッドの全行（チェック状態含む）
     * @param selectedNonyusakiId  選択中の納入先ID
     * @param selectedPlantId      選択中のプラントID
     * @param nonyusakiList  納入先検索結果（グリッド1）
     * @param plantList      プラント検索結果（グリッド2）
     * @param delivery       元のdelivery（遷移元区分等・nullの場合あり）
     * @return 選定結果（成功時はMCM2002U遷移パラメータを含む）
     */
    public SenteiResult sentei(List<Mcm2001uRowDto> rows,
                                BigDecimal selectedNonyusakiId,
                                BigDecimal selectedPlantId,
                                List<Map<String, Object>> nonyusakiList,
                                List<Map<String, Object>> plantList,
                                Mcm2001uDeliveryDto delivery) {

        SenteiResult result = new SenteiResult();

        /*
         * 【変換元】L136-139
         *   元コード: If Me.MCM_MA_NONYUSAKIDataGridView.RowCount = 0 Then
         *             MyBase.DisplayMessage(CPMessageConstant.MSG_0067)
         */
        if (nonyusakiList == null || nonyusakiList.isEmpty()) {
            result.setErrorMessage("納入先データが存在しません。先に検索を実行してください。");
            return result;
        }

        /*
         * 【変換元】L144-157
         *   元コード: For i = 0 To MCM_TM_MITSUMORIDataGridView.RowCount - 1
         *             If Cells(SENTAKU_MITSUMORI_CheckBox).Value = CHECKBOX_FLAG_ON Then
         *               ReDim Preserve keiyakujikanId(count)
         *               keiyakujikanId(count) = TM_KEIYAKUJIKAN_ID
         */
        // チェック行を収集
        List<String> checkedIds = new ArrayList<>();
        if (rows != null) {
            for (Mcm2001uRowDto row : rows) {
                if (row.getCheckbox() == 1 && row.getTmKeiyakujikanId() != null) {
                    checkedIds.add(row.getTmKeiyakujikanId().toPlainString());
                }
            }
        }

        /*
         * 【変換元】L160-172
         *   元コード: If checkFlg Then
         *             sql = "SELECT TMD.KIKIMEISAI_ID, TME.KOTAIKANRI_ID, COUNT(*) ..."
         *             If listCount > 1 Then
         *               MyBase.DisplayMessage(CPMessageConstant.MSG_0056)
         *               Return
         */
        // チェック行が存在する場合、重複チェック
        if (!checkedIds.isEmpty()) {
            boolean hasDuplicate = jdbcRepository.hasDuplicateKikimeisai(checkedIds);
            if (hasDuplicate) {
                result.setErrorMessage("選択された見積に機器明細の重複があります。");
                return result;
            }
        }

        /*
         * 【変換元】L180-190
         *   元コード: Dim nonyusakiId = CInt(MCM_MA_NONYUSAKIDataGridView.CurrentRow.Cells(...))
         *             Dim plantId = CInt(MCM_MA_PLANTDataGridView.CurrentRow.Cells(...))
         */
        // プラント選択チェック
        if (selectedPlantId == null) {
            result.setErrorMessage("プラントを選択してください。");
            return result;
        }

        /*
         * 【変換元】L192-230
         *   元コード: 納入先情報取得 → プラント情報取得 → delivery構築
         *             newDelMcm2002u.NonyusakiID = nonyusakiId
         *             newDelMcm2002u.NonyusakiCD = nonyusakiCd ... etc
         *             ForwardScreen(Mcm2002uScreen, newDelMcm2002u)
         */
        // 納入先情報を検索結果から取得
        String nonyusakiCd = "";
        String nonyusakiNk = "";
        if (selectedNonyusakiId != null) {
            for (Map<String, Object> nRow : nonyusakiList) {
                BigDecimal nId = decimalId(nRow.get("NONYUSAKI_ID"));
                if (nId != null && nId.compareTo(selectedNonyusakiId) == 0) {
                    nonyusakiCd = (String) nRow.getOrDefault("NONYUSAKI_CD", "");
                    nonyusakiNk = (String) nRow.getOrDefault("NONYUSAKI_NK", "");
                    break;
                }
            }
        }

        // プラント情報を検索結果から取得
        String plantNk = "";
        String supportId = "";
        if (plantList != null) {
            for (Map<String, Object> pRow : plantList) {
                BigDecimal pId = decimalId(pRow.get("PLANT_ID"));
                if (pId != null && pId.compareTo(selectedPlantId) == 0) {
                    supportId = (String) pRow.getOrDefault("SUPPORT_ID", "");
                    Object pNk = pRow.get("PLANT_NK");
                    plantNk = pNk != null ? pNk.toString() : "";
                    break;
                }
            }
        }

        // 結果構築（MCM2002Uへの遷移パラメータ）
        result.setSuccess(true);
        result.setNonyusakiId(selectedNonyusakiId != null ? selectedNonyusakiId.intValue() : 0);
        result.setNonyusakiCd(nonyusakiCd);
        result.setNonyusakiNk(nonyusakiNk);
        result.setPlantId(selectedPlantId.intValue());
        result.setPlantNk(plantNk);
        result.setSupportId(supportId);
        result.setTmKeiyakujikanId(
                checkedIds.isEmpty() ? null : checkedIds.toArray(new String[0]));

        /*
         * 【変換元】L232-240
         *   元コード: If IsNotNull(delivery) Then
         *             newDelMcm2002u.UmKihonMitsumoriID = delivery.UmKihonMitsumoriID
         *             newDelMcm2002u.SeniMotoKbn = delivery.SeniMotoKbn
         *           Else
         *             newDelMcm2002u.SeniMotoKbn = Mcm2002uConstant.SENIMOTO_KBN_KENSAKU
         */
        if (delivery != null) {
            result.setUmKihonMitsumoriId(delivery.getUmKihonMitsumoriId());
            result.setSeniMotoKbn(delivery.getSeniMotoKbn());
        } else {
            // 検索画面から直接遷移の場合
            // TODO: Mcm2002uConstant.SENIMOTO_KBN_KENSAKU の値を設定
            result.setSeniMotoKbn(0);
        }

        log.info("MCM2001U 選定完了: 納入先ID={}, プラントID={}, チェック行数={}",
                result.getNonyusakiId(), result.getPlantId(), checkedIds.size());

        return result;
    }
    private static BigDecimal decimalId(Object value) { return value == null ? null : new BigDecimal(value.toString()); }
}
