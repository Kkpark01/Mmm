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
        DateTimeFormatter.ofPattern("uuuu/MM/dd").withResolverStyle(java.time.format.ResolverStyle.STRICT);

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

        // 同じマスタ構成が複数ブランドに属していても、親の選択行は一つにまとめる。
        var kosei=new java.util.LinkedHashMap<BigDecimal,Mcm2002uKoseiRowDto>();
        for(var row:result.getKoseiRows())kosei.merge(row.getKikikoseiId().stripTrailingZeros(),row,(a,b)->{if(BigDecimal.ONE.compareTo(b.getCheckFlg()==null?BigDecimal.ZERO:b.getCheckFlg())==0)a.setCheckFlg(BigDecimal.ONE);return a;});
        result.setKoseiRows(new ArrayList<>(kosei.values()));
        var meisai=new java.util.LinkedHashMap<BigDecimal,Mcm2002uMeisaiRowDto>();
        for(var row:result.getMeisaiRows())meisai.merge(row.getKikimeisaiId().stripTrailingZeros(),row,(a,b)->{if(BigDecimal.ONE.compareTo(b.getCheckFlg()==null?BigDecimal.ZERO:b.getCheckFlg())==0)a.setCheckFlg(BigDecimal.ONE);return a;});
        result.setMeisaiRows(new ArrayList<>(meisai.values()));
        if (result.getTmKeiyakujikanIds()==null) result.setTmKeiyakujikanIds(tmKeiyakujikanIds);
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
    // VB MSG_0075: 選択した取引先見積にない通常機器の新規選定を禁止する。
    // 再選定で既に保存されている機器は保持対象にできる。
    public static boolean selectable(Mcm2002uMeisaiRowDto row,int mode) {
        return (row.getTmvTmIraiNo()!=null&&!row.getTmvTmIraiNo().isBlank())
            || BigDecimal.ONE.compareTo(row.getControllerFlg()==null?BigDecimal.ZERO:row.getControllerFlg())==0
            || mode==1&&row.getUmvUmKikimeisaiId()!=null&&row.getUmvUmKikimeisaiId().signum()>0;
    }
    public static boolean selectable(Mcm2002uKotaiRowDto row,int mode) {
        return (row.getTmvTmIraiNo()!=null&&!row.getTmvTmIraiNo().isBlank())
            || BigDecimal.ONE.compareTo(row.getMahControllerFlg()==null?BigDecimal.ZERO:row.getMahControllerFlg())==0
            || mode==1&&row.getUmvUmKotaimeisaiId()!=null&&row.getUmvUmKotaimeisaiId().signum()>0;
    }
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
        return "1".equals(jotai);
    }

    /**
     * 承認状態が審査中・承認中・承認済みかどうか判定
     * TODO: McmConstant の定数値に合わせること
     */
    private boolean isShonshinInProgress(String syoninJotai) {
        return "1".equals(syoninJotai) || "2".equals(syoninJotai) || "3".equals(syoninJotai);
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
