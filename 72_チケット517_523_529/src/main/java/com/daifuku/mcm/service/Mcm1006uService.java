package com.daifuku.mcm.service;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.daifuku.mcm.constants.Mcm1006uConstants;
import com.daifuku.mcm.form.Mcm1006uForm;
import com.daifuku.mcm.repository.Mcm1006uRepository;

/**
 * 【変換元】Mcm1006u1Screen.vb / Mcm1006u2Screen.vb
 * MCM1006U 取引先契約機器選定 Service
 */
@Service
public class Mcm1006uService {

    private static final Logger log = LoggerFactory.getLogger(Mcm1006uService.class);
    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");

    @Autowired
    private Mcm1005uService mcm1005uService;

    /**
     * 【移植】元VB: Mcm1006u1Screen.nextButton / Mcm1006u2Screen.tekiyoButton.AuthorityIsThrough=False
     *   MCM1006Uは常にMCM1005U（取引先契約内容）から起動されるサブ画面のため、
     *   権限区分は呼び出し元と同じ「取引先契約関連」（機能ID=MCM1005U）で判定する。
     */
    public boolean canUpdate(String user) {
        return mcm1005uService.canUpdate(user);
    }
    private static final DateTimeFormatter DATE_FMT2 = DateTimeFormatter.ISO_LOCAL_DATE;

    @Autowired
    private Mcm1006uRepository repository;

    // =========================================================================
    // 1画面目: データロード
    // 【変換元】Mcm1006u1Screen_Load()
    // =========================================================================

    /**
     * 1画面目用見積データをロードし、TmKeiyakujikanIDに基づいてチェック状態を設定する。
     */
    public List<Mcm1006uForm.MitsumoriRowForm> loadMitsumori(
            BigDecimal plantId, BigDecimal torihikisakiId,
            List<BigDecimal> selectedIds) {
        List<Mcm1006uForm.MitsumoriRowForm> rows = repository.findMitsumori(plantId, torihikisakiId);
        if (selectedIds != null && !selectedIds.isEmpty()) {
            Set<BigDecimal> selectedSet = new LinkedHashSet<>(selectedIds);
            for (Mcm1006uForm.MitsumoriRowForm row : rows) {
                if (row.getTmKeiyakujikanId() != null && selectedSet.stream().anyMatch(id -> eq(id, row.getTmKeiyakujikanId()))) {
                    row.setCheckFlg(Mcm1006uConstants.CHECKBOX_ON);
                }
            }
        }
        return rows;
    }

    /**
     * 1画面目の見積単価グリッド（参照専用）をロードする。
     * 保守契約時間・点検・保守方法・サービス形態はMCM1005Uと同じ表示ラベルを付与する。
     * 【変換元】Mcm1006u1DataSet.MCM_TM_TANKA
     */
    public List<Map<String, Object>> loadStep1Tanka(BigDecimal plantId, BigDecimal torihikisakiId) {
        if (plantId == null || torihikisakiId == null) return new ArrayList<>();
        List<Map<String, Object>> rows = repository.findStep1Tanka(plantId, torihikisakiId);
        return mcm1005uService.equipmentDisplay(rows, mcm1005uService.maintenanceTimeLabels());
    }

    public List<Map<String, Object>> loadStep1Kikan(BigDecimal plantId, BigDecimal clientId) {
        if (plantId == null || clientId == null) return new ArrayList<>();
        return repository.findStep1Kikan(plantId, clientId);
    }

    /** 保存用コードと表示名を分離し、VB共通DataTableの名称を表示する。 */
    public List<Map<String, Object>> estimateLabels(List<Mcm1006uForm.MitsumoriRowForm> rows) {
        List<Map<String, Object>> values = new ArrayList<>();
        for (var row : rows) {
            Map<String, Object> value = new HashMap<>();
            value.put("HOSYUHOHO", row.getHosyuhoho());
            value.put("TENKENUMU", row.getTenkenumu());
            String day = row.getTenkenkanoyobi() == null ? "" : row.getTenkenkanoyobi();
            value.put("YOBI_LABEL", Map.of("1", "月～金", "2", "月～土", "3", "月～日", "4", "土・日").getOrDefault(day, day));
            String night = row.getYakantaioumu() == null ? "" : row.getYakantaioumu();
            value.put("YAKAN_LABEL", Map.of("0", "夜間対応なし", "1", "夜間対応あり").getOrDefault(night, night));
            values.add(value);
        }
        return mcm1005uService.equipmentDisplay(values, Map.of());
    }

    // =========================================================================
    // 1画面目: バリデーション（「次へ」ボタン）
    // 【変換元】Mcm1006u1Screen.vb nextButton_Click()
    // =========================================================================

    /**
     * 1画面目バリデーション。エラーがあればメッセージを返す（最初の1件）。
     * VBのDisplayMessage()に相当 — 1件エラーで処理停止。
     *
     * チェック順:
     *   1. 期間必須
     *   2. 期間大小（開始 <= 終了）
     *   3. 期間1年以内
     *   4. 見積未選択（MSG_0037）
     *   5. 重複個体管理（MSG_0038）
     *   6. 出精値引きチェック（MSG_0079/0080/0082相当）
     *
     * @return エラーメッセージ（nullなら正常）
     */
    public String validateStep1(Mcm1006uForm form) {
        form.setWarningMsg(null);

        // 1. 期間必須
        if (isEmpty(form.getKaisiDt()) || isEmpty(form.getSyuryoDt())) {
            return Mcm1006uConstants.MSG_ERR_KIKAN_REQUIRED;
        }

        LocalDate kaisiDate;
        LocalDate syuryoDate;
        try {
            kaisiDate  = parseDate(form.getKaisiDt());
            syuryoDate = parseDate(form.getSyuryoDt());
        } catch (DateTimeParseException e) {
            return Mcm1006uConstants.MSG_ERR_KIKAN_REQUIRED;
        }

        // 2. 期間大小チェック（開始 <= 終了）
        if (kaisiDate.isAfter(syuryoDate)) {
            return Mcm1006uConstants.MSG_ERR_KIKAN_RELATION;
        }

        // 3. 期間1年以内チェック
        // 【変換元】IsCheckRelation(syuryoDtDate, kaisiDtDate.AddYears(1).AddDays(-1))
        LocalDate kaisi1Year = kaisiDate.plusYears(1).minusDays(1);
        if (syuryoDate.isAfter(kaisi1Year)) {
            return Mcm1006uConstants.MSG_ERR_KIKAN_1YEAR;
        }

        // 4. 見積未選択チェック
        List<BigDecimal> checkedIds = getCheckedKeiyakujikanIds(form.getMitsumoriRows());
        if (checkedIds.isEmpty()) {
            return Mcm1006uConstants.MSG_ERR_MITSUMORI_NONE;
        }

        // 5. 重複個体管理チェック
        // VB共通DisplayMessageはMSG_0038Eで例外を投げるため、後続チェック・遷移を中止する。
        int duplicateCount = repository.countKotaiDuplicate(checkedIds);
        if (duplicateCount > 0) {
            return Mcm1006uConstants.MSG_ERR_KOTAI_DUPLICATE;
        }

        // 6. 出精値引きチェック
        List<BigDecimal> tmKikanIds = new ArrayList<>();
        int[] syusseinebikiGokei = {0};
        for (BigDecimal tmKeiyakujikanId : checkedIds) {
            String errMsg = checkTanka(tmKeiyakujikanId, tmKikanIds, syusseinebikiGokei,
                                       form.getKaisiDt(), syuryoDate);
            if (errMsg != null) {
                return errMsg;
            }
        }

        // バリデーション通過後: 集計した値をformに反映
        form.setTmKeiyakujikanIds(checkedIds);
        form.setTmKikanIds(tmKikanIds);
        form.setSyusseinebikiGokei(syusseinebikiGokei[0]);

        return null;
    }

    /**
     * 出精値引きチェック（1取引先見積契約時間分）
     * 【変換元】Mcm1006u1Screen.vb checkTanka()
     *
     * @param tmKeiyakujikanId チェック対象の取引先見積契約時間ID
     * @param tmKikanIds       取引先見積期間IDリスト（蓄積用、参照渡し相当）
     * @param syusseinebikiGokei 出精値引き合計（蓄積用、参照渡し相当）
     * @param kaisiDtView      入力開始日文字列
     * @param syuryoDtView     入力終了日
     * @return エラーメッセージ（nullなら正常）
     */
    private String checkTanka(
            BigDecimal tmKeiyakujikanId,
            List<BigDecimal> tmKikanIds,
            int[] syusseinebikiGokei,
            String kaisiDtView,
            LocalDate syuryoDtView) {

        List<Map<String, Object>> kikanData = repository.findKikanForSyusseinebiki(
            tmKeiyakujikanId, kaisiDtView, formatDate(syuryoDtView));

        // 該当する取引先見積期間が0件の場合エラー（VB: MSG_0080）
        if (kikanData.isEmpty()) {
            return Mcm1006uConstants.MSG_ERR_KIKAN_NO_DATA;
        }

        boolean syusseinebikiCheck = false;
        BigDecimal prevKeiyakujikanId = null;
        Map<BigDecimal, Boolean> icchiKikanMap = new HashMap<>();
        LocalDate inputKaisi = parseDate(kaisiDtView);
        LocalDate kaishiDtMin = null;
        LocalDate syuryoDtMax = null;

        for (Map<String, Object> dataRow : kikanData) {
            BigDecimal keiyakujikanIdData = toBD(dataRow.get("TM_KEIYAKUJIKAN_ID"));
            BigDecimal tmKikanIdData      = toBD(dataRow.get("TM_KIKAN_ID"));
            int syusseinebikiSql          = toInt(dataRow.get("SYUSSEINEBIKI_KIN"));
            String kaisiDtSql             = toStr(dataRow.get("KAISI_DT"));
            String syuryoDtSql            = toStr(dataRow.get("SYURYO_DT"));

            LocalDate kikanKaisi  = parseDate(kaisiDtSql);
            LocalDate kikanSyuryo = parseDate(syuryoDtSql);

            // 取得した期間の開始日の最小、終了日の最大を取得（VB: kaishiDtMin/syuryoDtMax）
            if (kaishiDtMin == null || kikanKaisi.isBefore(kaishiDtMin)) kaishiDtMin = kikanKaisi;
            if (syuryoDtMax == null || kikanSyuryo.isAfter(syuryoDtMax)) syuryoDtMax = kikanSyuryo;

            // 取引先見積期間ID蓄積
            if (tmKikanIdData != null && !tmKikanIds.contains(tmKikanIdData)) {
                tmKikanIds.add(tmKikanIdData);
            }

            // TM_KEIYAKUJIKAN_ID が変わった場合（前グループのチェック）
            if (prevKeiyakujikanId != null && !prevKeiyakujikanId.equals(keiyakujikanIdData)) {
                if (syusseinebikiCheck && !icchiKikanMap.containsKey(prevKeiyakujikanId)) {
                    return Mcm1006uConstants.MSG_ERR_SYUSSEI_KIKAN;
                }
                syusseinebikiCheck = false;
                icchiKikanMap.clear();
            }
            prevKeiyakujikanId = keiyakujikanIdData;

            // 出精値引きあり判定
            if (syusseinebikiSql > 0) {
                syusseinebikiCheck = true;
                syusseinebikiGokei[0] += syusseinebikiSql;
            }

            // 期間一致チェック（画面上の期間と完全一致、または「無期限・年間契約」の特例）
            // 【変換元】(kaisiDtSql = kaisiDtView And syuryoDtSql = syuryoDtView) Or
            //           (syuryoDtSql = "9999/12/31" And kaisiDtSql <= kaisiDtView And
            //            DateDiff("m", kaisiDtView, syuryoDtView) + 1 = 12)
            boolean exactMatch = kikanKaisi.equals(inputKaisi) && kikanSyuryo.equals(syuryoDtView);
            boolean openEndedAnnual = kikanSyuryo.equals(LocalDate.of(9999, 12, 31))
                    && !kikanKaisi.isAfter(inputKaisi)
                    && (ChronoUnit.MONTHS.between(
                            java.time.YearMonth.from(inputKaisi), java.time.YearMonth.from(syuryoDtView)) + 1) == 12;
            if (exactMatch || openEndedAnnual) {
                icchiKikanMap.put(keiyakujikanIdData, true);
            }
        }

        // 最終グループチェック
        if (syusseinebikiCheck && prevKeiyakujikanId != null
                && !icchiKikanMap.containsKey(prevKeiyakujikanId)) {
            return Mcm1006uConstants.MSG_ERR_SYUSSEI_KIKAN;
        }

        // 入力された契約期間を見積期間が網羅しているかチェック（VB: MSG_0082）
        // 画面上の開始日 < 取得した最小開始日、または 取得した最大終了日 < 画面上の終了日 の場合エラー
        if (inputKaisi.isBefore(kaishiDtMin) || syuryoDtMax.isBefore(syuryoDtView)) {
            return Mcm1006uConstants.MSG_ERR_KIKAN_COVERAGE;
        }

        return null;
    }

    // =========================================================================
    // 2画面目: データロード
    // 【変換元】Mcm1006u2Screen_Load()
    // =========================================================================

    /**
     * 2画面目データ（機器構成/機器明細/個体明細/単価）をロードしformに設定する。
     * SeniMotoKbn=TAB_HENKO の場合はDeliveryのチェック状態を復元する。
     */
    public void loadStep2Data(Mcm1006uForm form) {
        List<BigDecimal> keiyakujikanIds = form.getTmKeiyakujikanIds();
        List<BigDecimal> kikanIds        = form.getTmKikanIds();
        BigDecimal plantId               = form.getPlantId();

        // DBからロード
        List<Mcm1006uForm.KoseiRowForm>  koseiRows  = repository.findKosei(keiyakujikanIds, plantId);
        List<Mcm1006uForm.MeisaiRowForm> meisaiRows = repository.findMeisai(keiyakujikanIds, plantId);
        List<Mcm1006uForm.KotaiRowForm>  kotaiRows  = repository.findKotai(keiyakujikanIds, plantId);
        List<Mcm1006uForm.TankaRowForm>  tankaRows  = repository.findTanka(kikanIds, plantId);

        // 依頼番号（複数可）を構成・明細それぞれに補完する
        // 【変換元】Mcm1006u2Screen.vb getTmIraiNoKosei() / getTmIraiNoMeisai()
        populateIraiNos(keiyakujikanIds, koseiRows, meisaiRows);

        if (form.getSeniMotoKbn() == Mcm1006uConstants.SENIMOTO_TAB_HENKO || form.isRestorePendingSelection() || (form.getKoseiRows() != null && !form.getKoseiRows().isEmpty())) {
            // TAB_HENKO: Deliveryの保存済みチェック状態を復元する
            // 【変換元】Mcm1006u2Screen_Load() TAB_HENKO ブロック
            restoreCheckState(form, koseiRows, meisaiRows, kotaiRows, tankaRows);
        }
        // INSERT（新規）の場合はDBのCHECK_FLGをそのまま使用

        form.setKoseiRows(koseiRows);
        form.setMeisaiRows(meisaiRows);
        form.setKotaiRows(kotaiRows);
        form.setTankaRows(tankaRows);
    }

    /**
     * 機器構成・機器明細の各行に、対応する依頼番号（カンマ区切り、複数可）を設定する。
     * 依頼番号が1件も無い行は null のままとし、「依頼番号なし＝選択不可」の判定に用いる。
     */
    private void populateIraiNos(List<BigDecimal> keiyakujikanIds,
            List<Mcm1006uForm.KoseiRowForm> koseiRows,
            List<Mcm1006uForm.MeisaiRowForm> meisaiRows) {
        for (Mcm1006uForm.KoseiRowForm row : koseiRows) {
            String iraiNo = repository.findIraiNosKosei(keiyakujikanIds, row.getKikikoseiId());
            row.setTmIraiNo(iraiNo == null || iraiNo.isEmpty() ? null : iraiNo);
        }
        for (Mcm1006uForm.MeisaiRowForm row : meisaiRows) {
            String iraiNo = repository.findIraiNosMeisai(keiyakujikanIds, row.getKikimeisaiId());
            row.setTmIraiNo(iraiNo == null || iraiNo.isEmpty() ? null : iraiNo);
        }
    }

    /**
     * TAB_HENKO時のチェック状態復元
     * 【変換元】Mcm1006u2Screen_Load() — 連携値の状態に合わせてチェックをつける
     */
    private void restoreCheckState(
            Mcm1006uForm form,
            List<Mcm1006uForm.KoseiRowForm>  koseiRows,
            List<Mcm1006uForm.MeisaiRowForm> meisaiRows,
            List<Mcm1006uForm.KotaiRowForm>  kotaiRows,
            List<Mcm1006uForm.TankaRowForm>  tankaRows) {

        // まずすべてのチェックをOFFに
        koseiRows.forEach(r  -> r.setCheckFlg(Mcm1006uConstants.CHECKBOX_OFF));
        meisaiRows.forEach(r -> r.setCheckFlg(Mcm1006uConstants.CHECKBOX_OFF));
        kotaiRows.forEach(r  -> r.setCheckFlg(Mcm1006uConstants.CHECKBOX_OFF));

        // 機器構成のチェック状態復元（Delivery保存済みから）
        for (Mcm1006uForm.KoseiRowForm deliveryKosei : form.getKoseiRows()) {
            for (Mcm1006uForm.KoseiRowForm row : koseiRows) {
                if (eq(row.getKikikoseiId(), deliveryKosei.getKikikoseiId())) {
                    row.setCheckFlg(isEmpty(row.getTmIraiNo()) ? 0 : deliveryKosei.getCheckFlg());
                    row.setCheckFlgOld(deliveryKosei.getCheckFlgOld());
                    row.setTkKikikoseiId(deliveryKosei.getTkKikikoseiId());
                }
            }
        }

        // 機器明細のチェック状態復元
        for (Mcm1006uForm.MeisaiRowForm deliveryMeisai : form.getMeisaiRows()) {
            for (Mcm1006uForm.MeisaiRowForm row : meisaiRows) {
                if (eq(row.getKikikoseiId(), deliveryMeisai.getKikikoseiId())
                        && eq(row.getKikimeisaiId(), deliveryMeisai.getKikimeisaiId())) {
                    row.setCheckFlg(isEmpty(row.getTmIraiNo()) ? 0 : deliveryMeisai.getCheckFlg());
                    row.setCheckFlgOld(deliveryMeisai.getCheckFlgOld());
                    row.setTkKikikoseiId(deliveryMeisai.getTkKikikoseiId());
                    row.setTkKikimeisaiId(deliveryMeisai.getTkKikimeisaiId());
                }
            }
        }

        // 個体明細のチェック状態復元
        for (Mcm1006uForm.KotaiRowForm deliveryKotai : form.getKotaiRows()) {
            for (Mcm1006uForm.KotaiRowForm row : kotaiRows) {
                if (eq(row.getKikimeisaiId(), deliveryKotai.getKikimeisaiId())
                        && eq(row.getKotaikanriId(), deliveryKotai.getKotaikanriId())) {
                    row.setCheckFlg(isEmpty(row.getTmIraiNo()) ? 0 : deliveryKotai.getCheckFlg());
                    row.setCheckFlgOld(deliveryKotai.getCheckFlgOld());
                    row.setTkKikikoseiId(deliveryKotai.getTkKikikoseiId());
                    row.setTkKikimeisaiId(deliveryKotai.getTkKikimeisaiId());
                    row.setTkKotaimeisaiId(deliveryKotai.getTkKotaimeisaiId());
                }
            }
        }

        // 単価のチェック状態復元
        // 【変換元】Mcm1006u2Screen_Load() 単価ループ ＋ tankaDelRowList
        //   Deliveryにあった単価行が今回のDBロード結果（tankaRows）に見つからない場合、
        //   契約期間の見直しで単価行自体が対象外になったことを意味する。
        //   VBはこれを tankaDelRowList に集め、DB上の該当行を明示的に削除する。
        // 前へ→次へを繰り返しても削除予定の元行を失わない。再選択された行は復元対象へ戻す。
        List<Mcm1006uForm.TankaRowForm> candidates = new ArrayList<>(form.getTankaRows());
        for (var removed : form.getTankaDelRows()) {
            if (candidates.stream().noneMatch(r -> eq(r.getTmTankaId(), removed.getTmTankaId())
                    && eq(r.getKikikoseiId(), removed.getKikikoseiId())
                    && eq(r.getKikimeisaiId(), removed.getKikimeisaiId()))) candidates.add(removed);
        }
        List<Mcm1006uForm.TankaRowForm> tankaDelRows = new ArrayList<>();
        for (Mcm1006uForm.TankaRowForm deliveryTanka : candidates) {
            boolean found = false;
            for (Mcm1006uForm.TankaRowForm row : tankaRows) {
                if (eq(row.getKikimeisaiId(), deliveryTanka.getKikimeisaiId())
                        && eq(row.getKikikoseiId(), deliveryTanka.getKikikoseiId())
                        && eq(row.getTmTankaId(), deliveryTanka.getTmTankaId())) {
                    row.setCheckFlg(deliveryTanka.getCheckFlg());
                    row.setCheckFlgOld(deliveryTanka.getCheckFlgOld());
                    row.setTkKikikoseiId(deliveryTanka.getTkKikikoseiId());
                    row.setTkKikimeisaiId(deliveryTanka.getTkKikimeisaiId());
                    row.setTkTankaId(deliveryTanka.getTkTankaId());
                    row.setTmTankaId(deliveryTanka.getTmTankaId());
                    found = true;
                }
            }
            if (!found) {
                tankaDelRows.add(deliveryTanka);
            }
        }
        form.setTankaDelRows(tankaDelRows);

        // 機器明細OFFの場合は対応単価もOFF
        // 【変換元】meisaiOffDataRowList ループ
        for (Mcm1006uForm.MeisaiRowForm meisaiRow : meisaiRows) {
            if (meisaiRow.getCheckFlg() == Mcm1006uConstants.CHECKBOX_OFF) {
                for (Mcm1006uForm.TankaRowForm tankaRow : tankaRows) {
                    if (eq(tankaRow.getKikikoseiId(), meisaiRow.getKikikoseiId())
                            && eq(tankaRow.getKikimeisaiId(), meisaiRow.getKikimeisaiId())) {
                        tankaRow.setCheckFlg(Mcm1006uConstants.CHECKBOX_OFF);
                    }
                }
            }
        }
    }

    // =========================================================================
    // 2画面目: 「摘要へ」処理
    // 【変換元】Mcm1006u2Screen.vb tekiyoButton_Click()
    // =========================================================================

    /**
     * 「摘要へ」ボタン押下時の処理。
     * setTankaKin（月割り計算）を実行し、チェックON行の情報をformに集約する。
     * MCM1005Uへ渡す結果がformのkosei/meisai/kotai/tankaRowsに格納される。
     *
     * @return エラーメッセージ（nullなら正常）
     */
    public String processComplete(Mcm1006uForm form) {
        form.setWarningMsg(null);

        // 【不具合修正 #527】構成・明細・個体のいずれかにチェックONがない場合、MSG_0023を返して処理を中止する。
        //   （VB原本はメッセージ表示後も次画面へ進むが、単体テスト仕様1006-23の期待結果に合わせ、
        //   画面遷移せず入力内容を保持する。単価の月割り計算等でフォームを変更する前に中止する。）
        boolean allKindsChecked =
                form.getKoseiRows().stream().anyMatch(r -> r.getCheckFlg() == Mcm1006uConstants.CHECKBOX_ON)
             && form.getMeisaiRows().stream().anyMatch(r -> r.getCheckFlg() == Mcm1006uConstants.CHECKBOX_ON)
             && form.getKotaiRows().stream().anyMatch(r -> r.getCheckFlg() == Mcm1006uConstants.CHECKBOX_ON);
        if (!allKindsChecked) {
            return Mcm1006uConstants.MSG_WARN_NO_CHECK;
        }

        // 単価の期間・金額按分（月割り計算）
        // 【変換元】Mcm1006u2Screen.vb setTankaKin()
        prorateSelectedTanka(form);

        // 数量補正: 機器明細のSURYO_NMを個体明細チェックON行のSURYO_NM合計で更新
        // 【変換元】tekiyoButton_Click() 内の SURYO_NM修正ロジック
        for (Mcm1006uForm.MeisaiRowForm meisaiRow : form.getMeisaiRows()) {
            BigDecimal kotaiSum = BigDecimal.ZERO;
            for (Mcm1006uForm.KotaiRowForm kotaiRow : form.getKotaiRows()) {
                if (eq(kotaiRow.getKikimeisaiId(), meisaiRow.getKikimeisaiId())
                        && kotaiRow.getCheckFlg() == Mcm1006uConstants.CHECKBOX_ON) {
                    if (kotaiRow.getSuryoNm() != null) {
                        kotaiSum = kotaiSum.add(kotaiRow.getSuryoNm());
                    }
                }
            }
            meisaiRow.setSuryoNm(kotaiSum);
        }

        // 単価の数量も同様に更新
        for (Mcm1006uForm.TankaRowForm tankaRow : form.getTankaRows()) {
            BigDecimal kotaiSum = BigDecimal.ZERO;
            for (Mcm1006uForm.KotaiRowForm kotaiRow : form.getKotaiRows()) {
                if (eq(kotaiRow.getKikimeisaiId(), tankaRow.getKikimeisaiId())
                        && kotaiRow.getCheckFlg() == Mcm1006uConstants.CHECKBOX_ON) {
                    if (kotaiRow.getSuryoNm() != null) {
                        kotaiSum = kotaiSum.add(kotaiRow.getSuryoNm());
                    }
                }
            }
            tankaRow.setSuryoNm(kotaiSum);
        }

        return null;
    }

    /**
     * チェックONの単価行について、契約期間（画面上の開始日・終了日）に応じて
     * 標準価格・仕切価格・開始日・終了日を月割りで按分する。
     * 【変換元】Mcm1006u2Screen.vb setTankaKin()
     *
     * VBの3パターン:
     *   ①取得期間が画面期間に完全に含まれる → 変更なし
     *   ②取得開始 <= 画面開始              → 新開始=画面開始
     *   ③画面開始 <  取得開始              → 新開始=取得開始
     *   （新終了は「取得終了・画面終了のうち画面期間に収まる方」）
     * 按分後の金額は「元期間の月数(tsukiTanka)で割った月額 × 新期間の月数(tsukiSai)」。
     * 元期間の月数で割った際の端数(mod)は、取得開始日が画面開始日以上の場合のみ加算する
     * （無期限期間は翌年以降も端数を加算する）。
     */
    private void prorateSelectedTanka(Mcm1006uForm form) {
        LocalDate contractStart = parseDate(form.getKaisiDt());
        LocalDate contractEnd   = parseDate(form.getSyuryoDt());
        int viewStartYm = yyyymm(contractStart);
        int viewEndYm   = yyyymm(contractEnd);

        java.util.SortedSet<BigDecimal> periodIds = new java.util.TreeSet<>();

        for (Mcm1006uForm.TankaRowForm row : form.getTankaRows()) {
            if (row.getCheckFlg() != Mcm1006uConstants.CHECKBOX_ON) continue;

            LocalDate rowStart = parseDate(row.getKaisiDt());
            boolean noEnd = row.getSyuryoDt() == null || row.getSyuryoDt().isBlank();
            LocalDate rowEnd = noEnd ? LocalDate.of(9999, 12, 31) : parseDate(row.getSyuryoDt());
            noEnd = noEnd || rowEnd.equals(LocalDate.of(9999, 12, 31));
            int rowStartYm = yyyymm(rowStart);
            int rowEndYm   = noEnd ? 999912 : yyyymm(rowEnd);

            if (viewStartYm <= rowStartYm && rowEndYm <= viewEndYm) {
                // パターン①: 画面期間内に取得期間が完全に含まれる → 変更なし
            } else {
                LocalDate newKaisiDt;
                LocalDate newSyuryoDt;
                int tsukiSai;
                if (rowStartYm <= viewStartYm) {
                    // パターン②: 取得開始 <= 画面開始
                    newKaisiDt = contractStart;
                    newSyuryoDt = (rowEndYm <= viewEndYm) ? rowEnd : contractEnd;
                    tsukiSai = monthsBetweenInclusive(contractStart, newSyuryoDt);
                } else {
                    // パターン③: 画面開始 < 取得開始
                    newKaisiDt = rowStart;
                    newSyuryoDt = (rowEndYm <= viewEndYm) ? rowEnd : contractEnd;
                    tsukiSai = monthsBetweenInclusive(rowStart, newSyuryoDt);
                }

                // 元期間全体の月数（終了日が無期限の場合は12ヶ月固定。VB: DateDiff("m",...)相当）
                int tsukiTanka = noEnd ? 12 : monthsBetweenInclusive(rowStart, rowEnd);
                if (tsukiTanka == 0) {
                    throw new ArithmeticException("単価の期間月数が0です。TM_TANKA_ID=" + row.getTmTankaId());
                }

                BigDecimal hyojunKin = row.getHyojunKin() != null ? row.getHyojunKin() : BigDecimal.ZERO;
                BigDecimal sikiriKin = row.getSikiriKin() != null ? row.getSikiriKin() : BigDecimal.ZERO;

                // VBの無期限分岐は契約年度にかかわらず端数を加算する。
                boolean addRemainder = noEnd || rowStartYm >= viewStartYm;
                row.setHyojunKin(proratedAmount(hyojunKin, tsukiTanka, tsukiSai, addRemainder));
                row.setSikiriKin(proratedAmount(sikiriKin, tsukiTanka, tsukiSai, addRemainder));
                row.setKaisiDt(formatDate(newKaisiDt));
                row.setSyuryoDt(formatDate(newSyuryoDt));
            }

            if (row.getTmKikanId() != null) {
                periodIds.add(row.getTmKikanId());
            }
        }

        form.setTmKikanIds(new ArrayList<>(periodIds));
    }

    /** VBの HYOJUN_KIN Mod tsukiTanka による端数計算＋月割り再分配を行う。 */
    private BigDecimal proratedAmount(BigDecimal kin, int tsukiTanka, int tsukiSai, boolean addRemainder) {
        BigDecimal[] divRem = kin.divideAndRemainder(BigDecimal.valueOf(tsukiTanka));
        BigDecimal monthly = divRem[0];
        BigDecimal remainder = divRem[1];
        BigDecimal result = monthly.multiply(BigDecimal.valueOf(tsukiSai));
        if (addRemainder) {
            result = result.add(remainder);
        }
        return result;
    }

    /** 【変換元】GetYYYYMM() : yyyy/MM/dd → YYYYMM(int) */
    private int yyyymm(LocalDate d) {
        return d.getYear() * 100 + d.getMonthValue();
    }

    /** VBの DateDiff("m", from, to.AddDays(1)) 相当（両端含む月数）。 */
    private int monthsBetweenInclusive(LocalDate from, LocalDate to) {
        return (int) ChronoUnit.MONTHS.between(
                java.time.YearMonth.from(from), java.time.YearMonth.from(to.plusDays(1)));
    }

    // =========================================================================
    // チェック済みIDリスト取得
    // =========================================================================
    public List<BigDecimal> getCheckedKeiyakujikanIds(List<Mcm1006uForm.MitsumoriRowForm> rows) {
        List<BigDecimal> ids = new ArrayList<>();
        if (rows == null) return ids;
        for (Mcm1006uForm.MitsumoriRowForm row : rows) {
            if (row.getCheckFlg() == Mcm1006uConstants.CHECKBOX_ON
                    && row.getTmKeiyakujikanId() != null) {
                if (!ids.contains(row.getTmKeiyakujikanId())) {
                    ids.add(row.getTmKeiyakujikanId());
                }
            }
        }
        return ids;
    }

    // =========================================================================
    // ユーティリティ
    // =========================================================================
    private LocalDate parseDate(String s) {
        if (s == null || s.isBlank()) throw new DateTimeParseException("blank", s == null ? "" : s, 0);
        // 画面の yyyy/MM/dd とDBの yyyy-MM-dd HH:mm:ss.S を日付として扱う。
        String normalized = s.trim().replace('/', '-');
        if (normalized.length() == 10) return LocalDate.parse(normalized, DATE_FMT2);
        return java.time.LocalDateTime.parse(normalized.replace(' ', 'T'), java.time.format.DateTimeFormatter.ISO_LOCAL_DATE_TIME).toLocalDate();
    }
    private String formatDate(LocalDate d) {
        return d == null ? null : d.format(DATE_FMT);
    }

    private boolean isEmpty(String s) { return s == null || s.isBlank(); }

    private boolean eq(BigDecimal a, BigDecimal b) {
        if (a == null && b == null) return true;
        if (a == null || b == null) return false;
        return a.compareTo(b) == 0;
    }

    private BigDecimal toBD(Object v) {
        if (v == null) return null;
        if (v instanceof BigDecimal) return (BigDecimal) v;
        return new BigDecimal(v.toString());
    }

    private String toStr(Object v) { return v == null ? null : v.toString(); }

    private int toInt(Object v) {
        if (v == null) return 0;
        if (v instanceof Number) return ((Number) v).intValue();
        try { return Integer.parseInt(v.toString()); } catch (Exception e) { return 0; }
    }
}

