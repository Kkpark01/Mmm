package com.daifuku.mcm.service;

import com.daifuku.mcm.constants.Mcm2002uConstants;
import com.daifuku.mcm.dto.*;
import com.daifuku.mcm.form.Mcm2002uForm;
import com.daifuku.mcm.repository.Mcm2002uJdbcRepository;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 【変換元】Mcm2002uScreen.vb（92,685 bytes / 22メソッド）
 * MCM2002U（カスタマー見積機器選定）Service
 *
 * <p>画面ロード時の検索分岐（新規/複製/修正）、チェックボックス連動、
 * 基本設定ボタン押下時の受渡しデータ構築を担当する。</p>
 *
 * @author MCM Migration Tool
 */
@Service
public class Mcm2002uService {

    private final Mcm2002uJdbcRepository repo;

    private static final DateTimeFormatter FMT_YYYYMMDD =
        DateTimeFormatter.ofPattern("yyyy/MM/dd");

    public Mcm2002uService(Mcm2002uJdbcRepository repo) {
        this.repo = repo;
    }

    // ================================================================
    // 内部クラス: 検索結果
    // ================================================================

    /**
     * 画面ロード時の検索結果を保持する。
     * Controller → Session → HTML で利用。
     */
    public static class SearchResult {
        private List<Mcm2002uKoseiRowDto> koseiRows = new ArrayList<>();
        private List<Mcm2002uMeisaiRowDto> meisaiRows = new ArrayList<>();
        private List<Mcm2002uKotaiRowDto> kotaiRows = new ArrayList<>();
        private List<Mcm2002uTankaRowDto> tankaRows = new ArrayList<>();

        /** 契約開始日テキスト表示フラグ */
        private boolean showKaisiDt;
        /** 契約時間帯テキスト表示フラグ */
        private boolean showKeiyakuJikantai;
        /** グリッド読み取り専用 */
        private boolean readOnly;
        /** 開始日テキスト読み取り専用 */
        private boolean kaisiDtReadOnly;
        /** 状態 */
        private String jotai;
        /** 承認状態 */
        private String syoninJotai;
        /** 既存開始日（見積修正時） */
        private String oldKaisiDt;
        /** 取引先見積契約時間IDリスト */
        private List<BigDecimal> tmKeiyakujikanIds;
        /** 遷移元区分 */
        private int seniMotoKbn;
        /** カスタマー基本見積ID */
        private BigDecimal umKihonMitsumoriId;
        /** プラントID */
        private BigDecimal plantId;

        // --- Getters / Setters ---
        public List<Mcm2002uKoseiRowDto> getKoseiRows() { return koseiRows; }
        public void setKoseiRows(List<Mcm2002uKoseiRowDto> v) { this.koseiRows = v; }
        public List<Mcm2002uMeisaiRowDto> getMeisaiRows() { return meisaiRows; }
        public void setMeisaiRows(List<Mcm2002uMeisaiRowDto> v) { this.meisaiRows = v; }
        public List<Mcm2002uKotaiRowDto> getKotaiRows() { return kotaiRows; }
        public void setKotaiRows(List<Mcm2002uKotaiRowDto> v) { this.kotaiRows = v; }
        public List<Mcm2002uTankaRowDto> getTankaRows() { return tankaRows; }
        public void setTankaRows(List<Mcm2002uTankaRowDto> v) { this.tankaRows = v; }
        public boolean isShowKaisiDt() { return showKaisiDt; }
        public void setShowKaisiDt(boolean v) { this.showKaisiDt = v; }
        public boolean isShowKeiyakuJikantai() { return showKeiyakuJikantai; }
        public void setShowKeiyakuJikantai(boolean v) { this.showKeiyakuJikantai = v; }
        public boolean isReadOnly() { return readOnly; }
        public void setReadOnly(boolean v) { this.readOnly = v; }
        public boolean isKaisiDtReadOnly() { return kaisiDtReadOnly; }
        public void setKaisiDtReadOnly(boolean v) { this.kaisiDtReadOnly = v; }
        public String getJotai() { return jotai; }
        public void setJotai(String v) { this.jotai = v; }
        public String getSyoninJotai() { return syoninJotai; }
        public void setSyoninJotai(String v) { this.syoninJotai = v; }
        public String getOldKaisiDt() { return oldKaisiDt; }
        public void setOldKaisiDt(String v) { this.oldKaisiDt = v; }
        public List<BigDecimal> getTmKeiyakujikanIds() { return tmKeiyakujikanIds; }
        public void setTmKeiyakujikanIds(List<BigDecimal> v) { this.tmKeiyakujikanIds = v; }
        public int getSeniMotoKbn() { return seniMotoKbn; }
        public void setSeniMotoKbn(int v) { this.seniMotoKbn = v; }
        public BigDecimal getUmKihonMitsumoriId() { return umKihonMitsumoriId; }
        public void setUmKihonMitsumoriId(BigDecimal v) { this.umKihonMitsumoriId = v; }
        public BigDecimal getPlantId() { return plantId; }
        public void setPlantId(BigDecimal v) { this.plantId = v; }
    }

    // ================================================================
    // 1. 画面ロード（メインエントリ）
    // 【変換元】Mcm2002uScreen_Load()
    //   元コード: If delivery.SeniMotoKbn = SENIMOTO_KBN_KENSAKU Then
    //               SearchInsert()
    //             Else
    //               SearchUpdate()
    //             End If
    // ================================================================
    @Transactional(readOnly = true)
    public SearchResult loadScreen(Mcm2002uDeliveryDto delivery) {
        SearchResult result = new SearchResult();
        result.setSeniMotoKbn(delivery.getSeniMotoKbn());
        result.setPlantId(delivery.getPlantId());
        result.setUmKihonMitsumoriId(delivery.getUmKihonMitsumoriId());

        BigDecimal plantId = delivery.getPlantId();
        List<BigDecimal> tmKeiyakujikanIds = toBigDecimalList(delivery.getTmKeiyakujikanId());
        BigDecimal umKihonMitsumoriId = delivery.getUmKihonMitsumoriId();

        if (delivery.getSeniMotoKbn() == Mcm2002uConstants.SENIMOTO_KBN_KENSAKU) {
            // ----- 新規契約作成 -----
            searchInsert(result, tmKeiyakujikanIds, plantId);
        } else {
            // ----- 契約内容編集（修正 or 複製） -----
            searchUpdate(result, delivery);
        }

        result.setTmKeiyakujikanIds(tmKeiyakujikanIds);
        return result;
    }

    // ================================================================
    // 2. 新規検索
    // 【変換元】Mcm2002uScreen.vb - SearchInsert()（52行）
    //   元コード:
    //     If IsNull(tmKeiyakujikanID) Then
    //       Me.MCM_MA_KIKIKOSEITableAdapter.Fill(..., plantId)
    //     Else
    //       Me.Fill(..., SQLXML_KIKIKOSEI_INSERT)
    //     End If
    // ================================================================
    private void searchInsert(SearchResult result,
                              List<BigDecimal> tmKeiyakujikanIds,
                              BigDecimal plantId) {

        if (tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty()) {
            // --- 見積選定なし：マスタ検索（プラントIDのみ） ---
            /**
             * 【変換元】SearchInsert() - 見積選定なし分岐
             *   元コード: Me.MCM_MA_KIKIKOSEITableAdapter.Fill(DataSet, plantId)
             */
            result.setKoseiRows(repo.searchKikokoseiInsert(
                Collections.emptyList(), plantId));
            result.setMeisaiRows(repo.searchKikimeisaiInsert(
                Collections.emptyList(), plantId));
            result.setKotaiRows(repo.searchKotaimeisaiInsert(
                Collections.emptyList(), plantId));
            // 単価なし
            result.setShowKaisiDt(true);
            result.setShowKeiyakuJikantai(true);

        } else {
            // --- 見積選定あり：_Ins.xml 検索 ---
            /**
             * 【変換元】SearchInsert() - 見積選定あり分岐
             *   元コード: Me.Fill(DataSet, params, SQLXML_KIKIKOSEI_INSERT)
             */
            result.setKoseiRows(repo.searchKikokoseiInsert(
                tmKeiyakujikanIds, plantId));
            result.setMeisaiRows(repo.searchKikimeisaiInsert(
                tmKeiyakujikanIds, plantId));
            result.setKotaiRows(repo.searchKotaimeisaiInsert(
                tmKeiyakujikanIds, plantId));
            result.setTankaRows(repo.searchTankaInsert(
                tmKeiyakujikanIds, plantId));

            // 機器構成に依頼NOを設定
            setTmIraiNoOnKosei(result.getKoseiRows(), tmKeiyakujikanIds);
            // 機器明細に依頼NOを設定
            setTmIraiNoOnMeisai(result.getMeisaiRows(), tmKeiyakujikanIds);
        }
    }

    /**
     * 【変換元】SearchInsert() 内のループ
     *   元コード: koseiRow.TMV_TM_IRAI_NO = Me.getTmIraiNoKosei(...)
     */
    private void setTmIraiNoOnKosei(List<Mcm2002uKoseiRowDto> rows,
                                     List<BigDecimal> tmKeiyakujikanIds) {
        for (Mcm2002uKoseiRowDto row : rows) {
            List<List<Object>> iraiData = repo.findTmIraiNoByKosei(
                tmKeiyakujikanIds, row.getKikikoseiId());
            if (!iraiData.isEmpty()) {
                // 複数依頼NOをカンマ区切りで結合
                String iraiNos = iraiData.stream()
                    .map(r -> String.valueOf(r.get(0)))
                    .distinct()
                    .collect(Collectors.joining(","));
                row.setTmvTmIraiNo(iraiNos);
                // 最初のTM_KIKIKOSEI_IDを設定
                row.setTmvTmKikikoseiId((BigDecimal) iraiData.get(0).get(1));
            }
        }
    }

    /**
     * 【変換元】SearchInsert() 内のループ
     *   元コード: meisaiRow.TMV_TM_IRAI_NO = Me.getTmIraiNoMeisai(...)
     */
    private void setTmIraiNoOnMeisai(List<Mcm2002uMeisaiRowDto> rows,
                                      List<BigDecimal> tmKeiyakujikanIds) {
        for (Mcm2002uMeisaiRowDto row : rows) {
            List<List<Object>> iraiData = repo.findTmIraiNoByMeisai(
                tmKeiyakujikanIds, row.getKikimeisaiId());
            if (!iraiData.isEmpty()) {
                String iraiNos = iraiData.stream()
                    .map(r -> String.valueOf(r.get(0)))
                    .distinct()
                    .collect(Collectors.joining(","));
                row.setTmvTmIraiNo(iraiNos);
                row.setTmvTmKikimeisaiId((BigDecimal) iraiData.get(0).get(1));
            }
        }
    }

    // ================================================================
    // 3. 修正/複製 検索
    // 【変換元】Mcm2002uScreen.vb - SearchUpdate()（174行）
    //   遷移元区分に応じて検索パターンを切替え:
    //     SAISENTEI_LINK: getUmKeiyakujikanId → searchXxxUpdate
    //     FUKUSEI_LINK: searchXxxByUm（複製）
    //   見積修正リンク時は状態チェックによりReadOnly判定
    // ================================================================
    private void searchUpdate(SearchResult result, Mcm2002uDeliveryDto delivery) {
        BigDecimal plantId = delivery.getPlantId();
        BigDecimal umKihonMitsumoriId = delivery.getUmKihonMitsumoriId();
        List<BigDecimal> tmKeiyakujikanIds = toBigDecimalList(delivery.getTmKeiyakujikanId());

        // 見積修正リンクの場合: 取引先見積契約時間IDをDBから取得
        if (delivery.getSeniMotoKbn() == Mcm2002uConstants.SENIMOTO_KBN_SAISENTEI_LINK) {
            /**
             * 【変換元】SearchUpdate()
             *   元コード: Me.tmKeiyakujikanID = Me.getUmKeiyakujikanId(Me.umKihonMitsumoriId)
             */
            List<String> ids = repo.findUmKeiyakujikanIds(umKihonMitsumoriId);
            tmKeiyakujikanIds = ids.stream()
                .map(BigDecimal::new)
                .collect(Collectors.toList());
            result.setTmKeiyakujikanIds(tmKeiyakujikanIds);
        }

        List<BigDecimal> umIds = Collections.singletonList(umKihonMitsumoriId);

        if (tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty()) {
            // --- 取引先見積が選定されていない場合（ソフト・コントローラのみ） ---
            /**
             * 【変換元】SearchUpdate() - tmKeiyakujikanID無し分岐
             *   元コード: Me.Fill(DataSet, {umKihonMitsumoriId, plantId}, SQLXML_KIKIKOSEI)
             */
            result.setKoseiRows(repo.searchKikokoseiByUm(umIds, plantId));
            result.setMeisaiRows(repo.searchKikimeisaiByUm(umIds, plantId));
            result.setKotaiRows(repo.searchKotaimeisaiByUm(umIds, plantId));

            if (delivery.getSeniMotoKbn() == Mcm2002uConstants.SENIMOTO_KBN_SAISENTEI_LINK) {
                // 見積修正: 単価テーブルも検索
                result.setTankaRows(repo.searchTankaUpdate(umIds, tmKeiyakujikanIds, plantId));
                // 開始日を取得
                LocalDate kaisiDt = repo.findKaisiDt(umKihonMitsumoriId);
                if (kaisiDt != null) {
                    result.setOldKaisiDt(kaisiDt.format(FMT_YYYYMMDD));
                }
                result.setShowKaisiDt(true);
                result.setShowKeiyakuJikantai(true);
            } else {
                // 複製: 開始日・契約時間帯を表示
                result.setShowKaisiDt(true);
                result.setShowKeiyakuJikantai(true);
            }

        } else {
            // --- 取引先見積が選定されている場合 ---
            if (delivery.getSeniMotoKbn() == Mcm2002uConstants.SENIMOTO_KBN_SAISENTEI_LINK) {
                /**
                 * 【変換元】SearchUpdate() - 見積修正＋見積選定あり
                 *   元コード: Me.Fill(DataSet, {umKihonMitsumoriId, tmKeiyakujikanID, plantId}, SQLXML_KIKIKOSEI_UPDATE)
                 */
                result.setKoseiRows(repo.searchKikokoseiUpdate(umIds, tmKeiyakujikanIds, plantId));
                result.setMeisaiRows(repo.searchKikimeisaiUpdate(umIds, tmKeiyakujikanIds, plantId));
                result.setKotaiRows(repo.searchKotaimeisaiUpdate(umIds, tmKeiyakujikanIds, plantId));
                result.setTankaRows(repo.searchTankaUpdate(umIds, tmKeiyakujikanIds, plantId));
            } else {
                /**
                 * 【変換元】SearchUpdate() - 複製＋見積選定あり
                 *   元コード: Me.Fill(DataSet, {umKihonMitsumoriId, tmKeiyakujikanID, plantId}, SQLXML_XXX_UPDATE)
                 */
                result.setKoseiRows(repo.searchKikokoseiUpdate(umIds, tmKeiyakujikanIds, plantId));
                result.setMeisaiRows(repo.searchKikimeisaiUpdate(umIds, tmKeiyakujikanIds, plantId));
                result.setKotaiRows(repo.searchKotaimeisaiUpdate(umIds, tmKeiyakujikanIds, plantId));
                result.setTankaRows(repo.searchTankaInsert(tmKeiyakujikanIds, plantId));
            }

            // 機器構成・明細に依頼NOを設定
            setTmIraiNoOnKosei(result.getKoseiRows(), tmKeiyakujikanIds);
            setTmIraiNoOnMeisai(result.getMeisaiRows(), tmKeiyakujikanIds);
        }

        // --- 見積修正リンク時の状態チェック ---
        if (delivery.getSeniMotoKbn() == Mcm2002uConstants.SENIMOTO_KBN_SAISENTEI_LINK) {
            /**
             * 【変換元】Mcm2002uScreen_Load()
             *   元コード: Dim jotaiData = Me.getJotai(Me.umKihonMitsumoriId)
             *   状態が「見積」以外 or 承認状態が「審査中」「承認中」「承認済み」→ ReadOnly
             */
            List<Object> jotaiData = repo.findJotai(umKihonMitsumoriId);
            String jotai = String.valueOf(jotaiData.get(0));
            String syoninJotai = String.valueOf(jotaiData.get(1));
            result.setJotai(jotai);
            result.setSyoninJotai(syoninJotai);

            // TODO: McmConstant の定数値に合わせて定義すること
            boolean isReadOnly = !isJotaiMitsumori(jotai) ||
                (isJotaiMitsumori(jotai) && isShonshinInProgress(syoninJotai));
            result.setReadOnly(isReadOnly);

            if (isReadOnly && (tmKeiyakujikanIds == null || tmKeiyakujikanIds.isEmpty())) {
                result.setKaisiDtReadOnly(true);
            }
        }

        // 個体明細のチェック連動ロジック（searchUpdate 後処理）
        applyKotaiCheckLogic(result);
    }

    /**
     * 【変換元】SearchUpdate() 内の個体明細チェック連動
     *   元コード: 個体テーブルの各行 → 構成・明細のチェック状態に基づいて設定
     */
    private void applyKotaiCheckLogic(SearchResult result) {
        // 構成のチェック済みIDセット
        Set<BigDecimal> checkedKoseiIds = result.getKoseiRows().stream()
            .filter(r -> isChecked(r.getCheckFlg()))
            .map(Mcm2002uKoseiRowDto::getKikikoseiId)
            .collect(Collectors.toSet());

        // 明細のチェック済みIDセット
        Set<BigDecimal> checkedMeisaiIds = result.getMeisaiRows().stream()
            .filter(r -> isChecked(r.getCheckFlg()))
            .map(Mcm2002uMeisaiRowDto::getKikimeisaiId)
            .collect(Collectors.toSet());

        // 個体行について: 親構成・親明細が両方チェックされている場合のみチェック可能
        for (Mcm2002uKotaiRowDto kotai : result.getKotaiRows()) {
            if (!checkedKoseiIds.contains(kotai.getMaeKikikoseiId()) ||
                !checkedMeisaiIds.contains(kotai.getMafKikimeisaiId())) {
                kotai.setCheckFlg(BigDecimal.ZERO);
            }
        }
    }

    // ================================================================
    // 4. バリデーション
    // 【変換元】Mcm2002uScreen.vb - KIHONSETTEIButton_Click() 先頭部分
    //   元コード:
    //     If IsNull(Me.tmKeiyakujikanID) Then
    //       '' 契約開始日チェック
    //       If IsNull(KAISI_DTTextBox.Text) Then → エラー
    //       If Not IsDate(KAISI_DTTextBox.Text) Then → エラー
    //       '' 契約時間帯チェック
    //       If IsNull(KEIYAKUJIKANTAITextBox.Text) Then → エラー
    //       If Not IsNumeric(...) Then → エラー
    //       If val < 0 Or val > 24 Then → エラー
    //     End If
    // ================================================================
    public List<String> validateKihonSettei(Mcm2002uForm form,
                                            boolean hasTmKeiyakujikanIds) {
        List<String> errors = new ArrayList<>();

        if (!hasTmKeiyakujikanIds) {
            // --- 契約開始日チェック ---
            if (form.getKaisiDt() == null || form.getKaisiDt().isBlank()) {
                errors.add("契約開始日を入力してください。");
            } else {
                try {
                    LocalDate.parse(form.getKaisiDt(), FMT_YYYYMMDD);
                } catch (DateTimeParseException e) {
                    errors.add("開始日付の書式を指定して下さい。(YYYY/MM/DD)");
                }
            }

            // --- 契約時間帯チェック ---
            if (form.getKeiyakuJikantai() == null || form.getKeiyakuJikantai().isBlank()) {
                errors.add("契約時間帯を入力してください。");
            } else {
                try {
                    int val = Integer.parseInt(form.getKeiyakuJikantai().trim());
                    if (val < 0 || val > 24) {
                        errors.add("契約時間帯は0～24の範囲で入力してください。");
                    }
                } catch (NumberFormatException e) {
                    errors.add("契約時間帯は数値で入力してください。");
                }
            }
        }

        return errors;
    }

    // ================================================================
    // 5. 基本設定ボタン → MCM2003U用受渡しデータ構築
    // 【変換元】Mcm2002uScreen.vb - KIHONSETTEIButton_Click()（601行）
    //
    //   元コード概要:
    //     (1) バリデーション → 前述validateKihonSettei()に分離
    //     (2) ブランド構成IDリスト取得 → getBrandKoseiId(plantId)
    //     (3) ブランドごとのループ:
    //         getBrandKosei(brandKoseiId) → brand情報取得
    //         構成→明細→個体→単価 の4階層を走査
    //         チェック有りの行を受渡しデータに設定
    //     (4) 開始期間リスト構築 → getKaisiKikan()
    //     (5) Mcm2003uDelivery にセットして画面遷移
    //
    //   ※本メソッドは601行の大規模ロジックを変換したもの。
    //     DataTable.Compute() 等は Java の Stream処理に変換。
    // ================================================================
    @Transactional(readOnly = true)
    public BuildResult buildDeliveryData(
            Mcm2002uForm form,
            SearchResult searchResult) {

        BuildResult buildResult = new BuildResult();
        List<String> errors = new ArrayList<>();

        BigDecimal plantId = searchResult.getPlantId();
        BigDecimal umKihonMitsumoriId = searchResult.getUmKihonMitsumoriId();
        List<BigDecimal> tmKeiyakujikanIds = searchResult.getTmKeiyakujikanIds();
        int seniMotoKbn = searchResult.getSeniMotoKbn();
        boolean isFukusei = (seniMotoKbn == Mcm2002uConstants.SENIMOTO_KBN_FUKUSEI_LINK);

        // チェック済み行を取得（form からの送信データを反映）
        applyFormCheckState(form, searchResult);

        // ========================================================
        // (2) ブランド構成IDリスト取得
        // 【変換元】KIHONSETTEIButton_Click()
        //   元コード: Dim brandKoseiIdList = Me.getBrandKoseiId(Me.plantId)
        // ========================================================
        List<String> brandKoseiIdList = repo.findBrandKoseiIds(plantId);

        // 受渡しデータリスト
        List<BrandData> brandDataList = new ArrayList<>();
        List<KoseiData> allKoseiList = new ArrayList<>();
        List<MeisaiData> allMeisaiList = new ArrayList<>();
        List<KotaiData> allKotaiList = new ArrayList<>();
        List<TankaData> allTankaList = new ArrayList<>();
        List<KikanData> allKikanList = new ArrayList<>();

        // ========================================================
        // (3) ブランドごとのループ
        // 【変換元】KIHONSETTEIButton_Click()
        //   元コード: For i = 0 To brandKoseiIdList.Length - 1
        //               Dim brandKosei = Me.getBrandKosei(brandKoseiIdList(i))
        //               For Each koseiRow In DataSet.MCM_MA_KIKIKOSEI.Select(...)
        //                 ...
        //               Next
        //             Next
        // ========================================================
        for (String brandKoseiIdStr : brandKoseiIdList) {
            BigDecimal brandKoseiId = new BigDecimal(brandKoseiIdStr);

            // ブランド情報取得
            List<Object> brandKosei = repo.findBrandKosei(brandKoseiId);
            if (brandKosei.isEmpty()) continue;

            BigDecimal brandId = (BigDecimal) brandKosei.get(0);
            String brandNk = (String) brandKosei.get(1);
            BigDecimal bkId = (BigDecimal) brandKosei.get(2);
            String brandsyosaiNk = (String) brandKosei.get(3);

            BrandData brandData = new BrandData();
            brandData.setBrandId(brandId);
            brandData.setBrandNk(brandNk);
            brandData.setBrandkoseiId(bkId);
            brandData.setBrandsyosaiNk(brandsyosaiNk);

            boolean hasCheckedKosei = false;

            // --- 構成ループ ---
            for (Mcm2002uKoseiRowDto koseiRow : searchResult.getKoseiRows()) {
                if (!isChecked(koseiRow.getCheckFlg())) continue;

                /**
                 * 【変換元】KIHONSETTEIButton_Click() - 構成ループ
                 *   元コード: If koseiRow.CHECK_FLG = CHECKBOX_FLAG_ON Then
                 *               koseiData.kikikoseiId = koseiRow.KIKIKOSEI_ID
                 *               ...
                 */
                hasCheckedKosei = true;
                KoseiData koseiData = new KoseiData();
                koseiData.setKikikoseiId(koseiRow.getKikikoseiId());
                koseiData.setCheckFlg(BigDecimal.ONE);
                koseiData.setTmKikikoseiId(koseiRow.getTmvTmKikikoseiId());
                koseiData.setUmKikikoseiId(koseiRow.getUmvUmKikikoseiId());
                koseiData.setBrandkoseiId(brandKoseiId);
                allKoseiList.add(koseiData);

                // 機器構成IDに属する明細行を検索
                BigDecimal kikikoseiId = koseiRow.getKikikoseiId();

                // --- 明細ループ ---
                for (Mcm2002uMeisaiRowDto meisaiRow : searchResult.getMeisaiRows()) {
                    if (!meisaiRow.getKikikoseiId().equals(kikikoseiId)) continue;

                    /**
                     * 【変換元】KIHONSETTEIButton_Click() - 明細ループ
                     *   元コード: For Each kikimeisaiRow In relatedMeisaiRows
                     *               If kikimeisaiRow.CHECK_FLG = CHECKBOX_FLAG_ON Then ...
                     */
                    MeisaiData meisaiData = new MeisaiData();
                    meisaiData.setKikimeisaiId(meisaiRow.getKikimeisaiId());
                    meisaiData.setKikikoseiId(kikikoseiId);
                    meisaiData.setCheckFlg(meisaiRow.getCheckFlg());
                    meisaiData.setTmKikimeisaiId(meisaiRow.getTmvTmKikimeisaiId());
                    meisaiData.setUmKikimeisaiId(meisaiRow.getUmvUmKikimeisaiId());
                    meisaiData.setSuryoNm(meisaiRow.getSuryoNm());
                    meisaiData.setControllerFlg(meisaiRow.getControllerFlg());
                    meisaiData.setBrandkoseiId(brandKoseiId);
                    allMeisaiList.add(meisaiData);

                    if (!isChecked(meisaiRow.getCheckFlg())) continue;

                    BigDecimal kikimeisaiId = meisaiRow.getKikimeisaiId();

                    // --- 個体ループ ---
                    /**
                     * 【変換元】KIHONSETTEIButton_Click() - 個体ループ
                     *   元コード: For Each kotaiRow In relatedKotaiRows
                     *               kotaiData.checkFlg = kotaiRow.CHECK_FLG
                     *               ...
                     */
                    int kotaiTankaCount = 0;
                    for (Mcm2002uKotaiRowDto kotaiRow : searchResult.getKotaiRows()) {
                        if (!kotaiRow.getMafKikimeisaiId().equals(kikimeisaiId)) continue;
                        if (!kotaiRow.getMaeKikikoseiId().equals(kikikoseiId)) continue;

                        KotaiData kotaiData = new KotaiData();
                        kotaiData.setKotaikanriId(kotaiRow.getMagKotaikanriId());
                        kotaiData.setCheckFlg(kotaiRow.getCheckFlg());
                        kotaiData.setTmKotaimeisaiId(kotaiRow.getTmvTmKotaimeisaiId());
                        kotaiData.setUmKotaimeisaiId(kotaiRow.getUmvUmKotaimeisaiId());
                        kotaiData.setKikimeisaiId(kikimeisaiId);
                        kotaiData.setKikikoseiId(kikikoseiId);
                        kotaiData.setBrandkoseiId(brandKoseiId);
                        allKotaiList.add(kotaiData);

                        if (isChecked(kotaiRow.getCheckFlg())) {
                            kotaiTankaCount++;
                        }
                    }

                    // --- 単価ループ ---
                    /**
                     * 【変換元】KIHONSETTEIButton_Click() - 単価ループ
                     *   元コード: For Each tankaRow In relatedTankaRows
                     *               tankaRow.CHECK_FLG = koseiRow.CHECK_FLG
                     *               tankaRow.UMV_UM_KIKIMEISAI_ID = umKikimeisaiId
                     *               tankaRow.SURYO_NM = kotaiTankaCount
                     */
                    BigDecimal umKikimeisaiId = meisaiRow.getUmvUmKikimeisaiId();
                    BigDecimal suryoNm = meisaiRow.getSuryoNm();

                    for (Mcm2002uTankaRowDto tankaRow : searchResult.getTankaRows()) {
                        if (!tankaRow.getKikimeisaiId().equals(kikimeisaiId)) continue;

                        TankaData tankaData = new TankaData();
                        tankaData.setCheckFlg(koseiRow.getCheckFlg());
                        tankaData.setUmKikimeisaiId(umKikimeisaiId);
                        // 2009/11/27 修正: suryoNm → kotaiTankaCount
                        tankaData.setSuryoNm(BigDecimal.valueOf(
                            kotaiTankaCount > 0 ? kotaiTankaCount : safeInt(suryoNm)));
                        tankaData.setUmTankaId(tankaRow.getUmvUmTankaId());
                        tankaData.setTmTankaId(tankaRow.getTmvTmTankaId());
                        tankaData.setKikimeisaiId(kikimeisaiId);
                        tankaData.setUmMitsumoriId(tankaRow.getUmvUmMitsumoriId());
                        tankaData.setHyojunKin(tankaRow.getHyojunKin());
                        tankaData.setSikiriKin(tankaRow.getSikiriKin());
                        tankaData.setKaisiDt(tankaRow.getKaisiDt());
                        tankaData.setControllerFlg(meisaiRow.getControllerFlg());
                        tankaData.setBrandkoseiId(brandKoseiId);
                        allTankaList.add(tankaData);
                    }

                    // --- コントローラの場合: 期間に単価がなければ追加 ---
                    /**
                     * 【変換元】KIHONSETTEIButton_Click() - コントローラ単価追加
                     *   元コード: If CHECK_FLG=ON And CONTROLLER_FLG=ON Then
                     *               For j = 0 To kaisiKikanList.Count - 1
                     *                 Dim tankaCount = DataSet.Compute("count(...)")
                     *                 If tankaCount = 0 Then → 新規追加
                     */
                    if (isChecked(meisaiRow.getCheckFlg()) &&
                        isChecked(meisaiRow.getControllerFlg())) {

                        List<List<Object>> kaisiKikanList = repo.findKaisiKikan(
                            tmKeiyakujikanIds, umKihonMitsumoriId, isFukusei);

                        for (List<Object> kikanRow : kaisiKikanList) {
                            LocalDate kaisiDate = (LocalDate) kikanRow.get(0);
                            if (kaisiDate != null && kaisiDate.getYear() == 9999 &&
                                kaisiDate.getMonthValue() == 12 && kaisiDate.getDayOfMonth() == 31) {
                                break; // 9999/12/31 で終了
                            }

                            // 既存の単価に該当期間が含まれるかチェック
                            boolean exists = allTankaList.stream().anyMatch(t ->
                                t.getKikimeisaiId().equals(kikimeisaiId) &&
                                t.getKaisiDt() != null &&
                                !t.getKaisiDt().isAfter(kaisiDate));
                            // TODO: DataSet.Compute相当の正確な期間包含チェックに置換

                            if (!exists) {
                                BigDecimal controllerKin = repo.findControllerKin(kikimeisaiId);

                                TankaData tankaData = new TankaData();
                                tankaData.setCheckFlg(BigDecimal.ONE);
                                tankaData.setUmKikimeisaiId(umKikimeisaiId);
                                tankaData.setSuryoNm(suryoNm);
                                tankaData.setUmTankaId(BigDecimal.ZERO);
                                tankaData.setTmTankaId(BigDecimal.ZERO);
                                tankaData.setKikimeisaiId(kikimeisaiId);
                                tankaData.setUmMitsumoriId(BigDecimal.ZERO);
                                tankaData.setHyojunKin(controllerKin);
                                tankaData.setSikiriKin(BigDecimal.ZERO);
                                tankaData.setKaisiDt(kaisiDate);
                                tankaData.setControllerFlg(meisaiRow.getControllerFlg());
                                tankaData.setBrandkoseiId(brandKoseiId);
                                allTankaList.add(tankaData);
                            }
                        }
                    }
                } // end meisai loop
            } // end kosei loop

            if (hasCheckedKosei) {
                brandDataList.add(brandData);
            }
        } // end brand loop

        // ========================================================
        // (4) 開始期間リスト構築
        // 【変換元】KIHONSETTEIButton_Click()
        //   元コード: Dim kaisiKikanList = Me.getKaisiKikan(tmKeiyakujikanIDList, umKihonMitsumoriId)
        // ========================================================
        List<List<Object>> kaisiKikanList = repo.findKaisiKikan(
            tmKeiyakujikanIds, umKihonMitsumoriId, isFukusei);

        for (List<Object> row : kaisiKikanList) {
            KikanData kikanData = new KikanData();
            kikanData.setKaisiDt((LocalDate) row.get(0));
            kikanData.setUmMitsumoriId((BigDecimal) row.get(1));
            kikanData.setUmKihonMitsumoriId((BigDecimal) row.get(2));
            kikanData.setUmuFlg((BigDecimal) row.get(3));
            allKikanList.add(kikanData);
        }

        buildResult.setBrandDataList(brandDataList);
        buildResult.setKoseiList(allKoseiList);
        buildResult.setMeisaiList(allMeisaiList);
        buildResult.setKotaiList(allKotaiList);
        buildResult.setTankaList(allTankaList);
        buildResult.setKikanList(allKikanList);
        buildResult.setErrors(errors);
        return buildResult;
    }

    /**
     * フォームのチェック状態を検索結果に適用する。
     * HTMLフォームから送信されたcheckFlg値を反映。
     */
    private void applyFormCheckState(Mcm2002uForm form, SearchResult result) {
        if (form.getKoseiRows() != null) {
            for (int i = 0; i < form.getKoseiRows().size() && i < result.getKoseiRows().size(); i++) {
                result.getKoseiRows().get(i).setCheckFlg(form.getKoseiRows().get(i).getCheckFlg());
            }
        }
        if (form.getMeisaiRows() != null) {
            for (int i = 0; i < form.getMeisaiRows().size() && i < result.getMeisaiRows().size(); i++) {
                result.getMeisaiRows().get(i).setCheckFlg(form.getMeisaiRows().get(i).getCheckFlg());
            }
        }
        if (form.getKotaiRows() != null) {
            for (int i = 0; i < form.getKotaiRows().size() && i < result.getKotaiRows().size(); i++) {
                result.getKotaiRows().get(i).setCheckFlg(form.getKotaiRows().get(i).getCheckFlg());
            }
        }
    }

    // ================================================================
    // 内部クラス: 受渡しデータ構築結果
    // ================================================================

    public static class BuildResult {
        private List<BrandData> brandDataList = new ArrayList<>();
        private List<KoseiData> koseiList = new ArrayList<>();
        private List<MeisaiData> meisaiList = new ArrayList<>();
        private List<KotaiData> kotaiList = new ArrayList<>();
        private List<TankaData> tankaList = new ArrayList<>();
        private List<KikanData> kikanList = new ArrayList<>();
        private List<String> errors = new ArrayList<>();

        public List<BrandData> getBrandDataList() { return brandDataList; }
        public void setBrandDataList(List<BrandData> v) { this.brandDataList = v; }
        public List<KoseiData> getKoseiList() { return koseiList; }
        public void setKoseiList(List<KoseiData> v) { this.koseiList = v; }
        public List<MeisaiData> getMeisaiList() { return meisaiList; }
        public void setMeisaiList(List<MeisaiData> v) { this.meisaiList = v; }
        public List<KotaiData> getKotaiList() { return kotaiList; }
        public void setKotaiList(List<KotaiData> v) { this.kotaiList = v; }
        public List<TankaData> getTankaList() { return tankaList; }
        public void setTankaList(List<TankaData> v) { this.tankaList = v; }
        public List<KikanData> getKikanList() { return kikanList; }
        public void setKikanList(List<KikanData> v) { this.kikanList = v; }
        public List<String> getErrors() { return errors; }
        public void setErrors(List<String> v) { this.errors = v; }
    }

    // ================================================================
    // 受渡しデータ DTO（MCM2003U連携用）
    // 【変換元】Mcm2003uData_Brand / Mcm2003uData_Kosei /
    //          Mcm2003uData_Meisai / Mcm2003uData_Kotai /
    //          Mcm2003uData_Tanka / Mcm2003uData_Kikan
    // ================================================================

    public static class BrandData {
        private BigDecimal brandId;
        private String brandNk;
        private BigDecimal brandkoseiId;
        private String brandsyosaiNk;

        public BigDecimal getBrandId() { return brandId; }
        public void setBrandId(BigDecimal v) { this.brandId = v; }
        public String getBrandNk() { return brandNk; }
        public void setBrandNk(String v) { this.brandNk = v; }
        public BigDecimal getBrandkoseiId() { return brandkoseiId; }
        public void setBrandkoseiId(BigDecimal v) { this.brandkoseiId = v; }
        public String getBrandsyosaiNk() { return brandsyosaiNk; }
        public void setBrandsyosaiNk(String v) { this.brandsyosaiNk = v; }
    }

    public static class KoseiData {
        private BigDecimal kikikoseiId;
        private BigDecimal checkFlg;
        private BigDecimal tmKikikoseiId;
        private BigDecimal umKikikoseiId;
        private BigDecimal brandkoseiId;

        public BigDecimal getKikikoseiId() { return kikikoseiId; }
        public void setKikikoseiId(BigDecimal v) { this.kikikoseiId = v; }
        public BigDecimal getCheckFlg() { return checkFlg; }
        public void setCheckFlg(BigDecimal v) { this.checkFlg = v; }
        public BigDecimal getTmKikikoseiId() { return tmKikikoseiId; }
        public void setTmKikikoseiId(BigDecimal v) { this.tmKikikoseiId = v; }
        public BigDecimal getUmKikikoseiId() { return umKikikoseiId; }
        public void setUmKikikoseiId(BigDecimal v) { this.umKikikoseiId = v; }
        public BigDecimal getBrandkoseiId() { return brandkoseiId; }
        public void setBrandkoseiId(BigDecimal v) { this.brandkoseiId = v; }
    }

    public static class MeisaiData {
        private BigDecimal kikimeisaiId;
        private BigDecimal kikikoseiId;
        private BigDecimal checkFlg;
        private BigDecimal tmKikimeisaiId;
        private BigDecimal umKikimeisaiId;
        private BigDecimal suryoNm;
        private BigDecimal controllerFlg;
        private BigDecimal brandkoseiId;

        public BigDecimal getKikimeisaiId() { return kikimeisaiId; }
        public void setKikimeisaiId(BigDecimal v) { this.kikimeisaiId = v; }
        public BigDecimal getKikikoseiId() { return kikikoseiId; }
        public void setKikikoseiId(BigDecimal v) { this.kikikoseiId = v; }
        public BigDecimal getCheckFlg() { return checkFlg; }
        public void setCheckFlg(BigDecimal v) { this.checkFlg = v; }
        public BigDecimal getTmKikimeisaiId() { return tmKikimeisaiId; }
        public void setTmKikimeisaiId(BigDecimal v) { this.tmKikimeisaiId = v; }
        public BigDecimal getUmKikimeisaiId() { return umKikimeisaiId; }
        public void setUmKikimeisaiId(BigDecimal v) { this.umKikimeisaiId = v; }
        public BigDecimal getSuryoNm() { return suryoNm; }
        public void setSuryoNm(BigDecimal v) { this.suryoNm = v; }
        public BigDecimal getControllerFlg() { return controllerFlg; }
        public void setControllerFlg(BigDecimal v) { this.controllerFlg = v; }
        public BigDecimal getBrandkoseiId() { return brandkoseiId; }
        public void setBrandkoseiId(BigDecimal v) { this.brandkoseiId = v; }
    }

    public static class KotaiData {
        private BigDecimal kotaikanriId;
        private BigDecimal checkFlg;
        private BigDecimal tmKotaimeisaiId;
        private BigDecimal umKotaimeisaiId;
        private BigDecimal kikimeisaiId;
        private BigDecimal kikikoseiId;
        private BigDecimal brandkoseiId;

        public BigDecimal getKotaikanriId() { return kotaikanriId; }
        public void setKotaikanriId(BigDecimal v) { this.kotaikanriId = v; }
        public BigDecimal getCheckFlg() { return checkFlg; }
        public void setCheckFlg(BigDecimal v) { this.checkFlg = v; }
        public BigDecimal getTmKotaimeisaiId() { return tmKotaimeisaiId; }
        public void setTmKotaimeisaiId(BigDecimal v) { this.tmKotaimeisaiId = v; }
        public BigDecimal getUmKotaimeisaiId() { return umKotaimeisaiId; }
        public void setUmKotaimeisaiId(BigDecimal v) { this.umKotaimeisaiId = v; }
        public BigDecimal getKikimeisaiId() { return kikimeisaiId; }
        public void setKikimeisaiId(BigDecimal v) { this.kikimeisaiId = v; }
        public BigDecimal getKikikoseiId() { return kikikoseiId; }
        public void setKikikoseiId(BigDecimal v) { this.kikikoseiId = v; }
        public BigDecimal getBrandkoseiId() { return brandkoseiId; }
        public void setBrandkoseiId(BigDecimal v) { this.brandkoseiId = v; }
    }

    public static class TankaData {
        private BigDecimal checkFlg;
        private BigDecimal umKikimeisaiId;
        private BigDecimal suryoNm;
        private BigDecimal umTankaId;
        private BigDecimal tmTankaId;
        private BigDecimal kikimeisaiId;
        private BigDecimal umMitsumoriId;
        private BigDecimal hyojunKin;
        private BigDecimal sikiriKin;
        private LocalDate kaisiDt;
        private BigDecimal controllerFlg;
        private BigDecimal brandkoseiId;

        public BigDecimal getCheckFlg() { return checkFlg; }
        public void setCheckFlg(BigDecimal v) { this.checkFlg = v; }
        public BigDecimal getUmKikimeisaiId() { return umKikimeisaiId; }
        public void setUmKikimeisaiId(BigDecimal v) { this.umKikimeisaiId = v; }
        public BigDecimal getSuryoNm() { return suryoNm; }
        public void setSuryoNm(BigDecimal v) { this.suryoNm = v; }
        public BigDecimal getUmTankaId() { return umTankaId; }
        public void setUmTankaId(BigDecimal v) { this.umTankaId = v; }
        public BigDecimal getTmTankaId() { return tmTankaId; }
        public void setTmTankaId(BigDecimal v) { this.tmTankaId = v; }
        public BigDecimal getKikimeisaiId() { return kikimeisaiId; }
        public void setKikimeisaiId(BigDecimal v) { this.kikimeisaiId = v; }
        public BigDecimal getUmMitsumoriId() { return umMitsumoriId; }
        public void setUmMitsumoriId(BigDecimal v) { this.umMitsumoriId = v; }
        public BigDecimal getHyojunKin() { return hyojunKin; }
        public void setHyojunKin(BigDecimal v) { this.hyojunKin = v; }
        public BigDecimal getSikiriKin() { return sikiriKin; }
        public void setSikiriKin(BigDecimal v) { this.sikiriKin = v; }
        public LocalDate getKaisiDt() { return kaisiDt; }
        public void setKaisiDt(LocalDate v) { this.kaisiDt = v; }
        public BigDecimal getControllerFlg() { return controllerFlg; }
        public void setControllerFlg(BigDecimal v) { this.controllerFlg = v; }
        public BigDecimal getBrandkoseiId() { return brandkoseiId; }
        public void setBrandkoseiId(BigDecimal v) { this.brandkoseiId = v; }
    }

    public static class KikanData {
        private LocalDate kaisiDt;
        private BigDecimal umMitsumoriId;
        private BigDecimal umKihonMitsumoriId;
        private BigDecimal umuFlg;

        public LocalDate getKaisiDt() { return kaisiDt; }
        public void setKaisiDt(LocalDate v) { this.kaisiDt = v; }
        public BigDecimal getUmMitsumoriId() { return umMitsumoriId; }
        public void setUmMitsumoriId(BigDecimal v) { this.umMitsumoriId = v; }
        public BigDecimal getUmKihonMitsumoriId() { return umKihonMitsumoriId; }
        public void setUmKihonMitsumoriId(BigDecimal v) { this.umKihonMitsumoriId = v; }
        public BigDecimal getUmuFlg() { return umuFlg; }
        public void setUmuFlg(BigDecimal v) { this.umuFlg = v; }
    }

    // ================================================================
    // ユーティリティメソッド
    // ================================================================

    private boolean isChecked(BigDecimal flg) {
        return flg != null && flg.compareTo(BigDecimal.ONE) == 0;
    }

    /**
     * 状態が「見積」かどうか判定
     * TODO: McmConstant.JOTAI_MITSUMORI の実値に合わせること
     */
    private boolean isJotaiMitsumori(String jotai) {
        return "10".equals(jotai);
    }

    /**
     * 承認状態が審査中・承認中・承認済みかどうか判定
     * TODO: McmConstant の定数値に合わせること
     */
    private boolean isShonshinInProgress(String syoninJotai) {
        return "20".equals(syoninJotai) ||
               "30".equals(syoninJotai) ||
               "40".equals(syoninJotai);
    }

    private List<BigDecimal> toBigDecimalList(String[] ids) {
        if (ids == null || ids.length == 0) return Collections.emptyList();
        List<BigDecimal> list = new ArrayList<>();
        for (String id : ids) {
            if (id != null && !id.isBlank()) {
                list.add(new BigDecimal(id.trim()));
            }
        }
        return list;
    }

    private int safeInt(BigDecimal val) {
        return val != null ? val.intValue() : 0;
    }
}
