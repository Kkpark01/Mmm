/**
 * 【変換元】Mcm0021uScreen.vb（724行）
 *   MCM0021U プラント付替え サービスクラス
 *   元コード: SearchButton_Click / TsukekaemotoButton_Click /
 *             TsukekaesakiButton_Click / ReplaceButton_Click /
 *             getTkKeiyakuList / getUkKeiyakuList / tsukekae_check
 *
 *   全てのボタンイベント処理をService層に集約。
 *   @Transactional で付替え処理の原子性を保証。
 */
package com.daifuku.mcm.service;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.daifuku.mcm.constants.Mcm0021uConstants;
import com.daifuku.mcm.dto.Mcm0021uKikikoseiDto;
import com.daifuku.mcm.dto.Mcm0021uKikimeisaiSearchDto;
import com.daifuku.mcm.dto.Mcm0021uMotoRowDto;
import com.daifuku.mcm.dto.Mcm0021uNonyusakiDto;
import com.daifuku.mcm.dto.Mcm0021uSakiRowDto;
import com.daifuku.mcm.entity.KikikoseiEntity;
import com.daifuku.mcm.entity.KikimeisaiEntity;
import com.daifuku.mcm.form.Mcm0021uForm.MotoRowForm;
import com.daifuku.mcm.repository.KikikoseiRepository;
import com.daifuku.mcm.repository.KikikotaikanriRepository;
import com.daifuku.mcm.repository.KikimeisaiRepository;
import com.daifuku.mcm.repository.Mcm0021uRepository;

/**
 * MCM0021U（プラント付替え）画面のサービスクラス
 */
@Service
public class Mcm0021uService {

    private static final Logger log = LoggerFactory.getLogger(Mcm0021uService.class);

    private final Mcm0021uRepository mcm0021uRepository;
    private final KikikoseiRepository kikikoseiRepository;
    private final KikimeisaiRepository kikimeisaiRepository;
    private final KikikotaikanriRepository kikikotaikanriRepository;

    public Mcm0021uService(
            Mcm0021uRepository mcm0021uRepository,
            KikikoseiRepository kikikoseiRepository,
            KikimeisaiRepository kikimeisaiRepository,
            KikikotaikanriRepository kikikotaikanriRepository) {
        this.mcm0021uRepository = mcm0021uRepository;
        this.kikikoseiRepository = kikikoseiRepository;
        this.kikimeisaiRepository = kikimeisaiRepository;
        this.kikikotaikanriRepository = kikikotaikanriRepository;
    }

    // =================================================================
    // 1. 検索処理
    // 【変換元】SearchButton_Click()
    //   元コード: Me.Fill(MCM_MA_NONYUSAKI, ...) + MCM_MA_KIKIKOSEI + MCM_MA_KIKIMEISAI
    // =================================================================

    /**
     * 納入先を検索する
     * @return 検索結果リスト
     */
    public List<Mcm0021uNonyusakiDto> searchNonyusaki(
            String nonyusakiCd, String nonyusakiNk, String supportId, String plantNk) {
        return mcm0021uRepository.searchNonyusaki(nonyusakiCd, nonyusakiNk, supportId, plantNk);
    }

    /**
     * 機器構成+ブランドを検索する
     */
    public List<Mcm0021uKikikoseiDto> searchKikikosei(
            String nonyusakiCd, String nonyusakiNk, String supportId, String plantNk) {
        return mcm0021uRepository.searchKikikosei(nonyusakiCd, nonyusakiNk, supportId, plantNk);
    }

    /**
     * 機器明細を検索する
     */
    public List<Mcm0021uKikimeisaiSearchDto> searchKikimeisai(
            String nonyusakiCd, String nonyusakiNk, String supportId, String plantNk) {
        return mcm0021uRepository.searchKikimeisai(nonyusakiCd, nonyusakiNk, supportId, plantNk);
    }

    // =================================================================
    // 2. 付替え元データ取得
    // 【変換元】TsukekaemotoButton_Click()
    //   元コード: Me.Fill(MCM_MA_KIKIKOSEI_MOTO, plantId)
    // =================================================================

    /**
     * 付替え元の機器構成データを取得する
     */
    public List<Mcm0021uMotoRowDto> findMotoByPlantId(BigDecimal plantId) {
        return mcm0021uRepository.findMotoByPlantId(plantId);
    }

    /** #258: 他の構成に明細があっても、選択した構成に明細がなければ元に設定しない。 */
    public boolean hasKikiDetails(BigDecimal kikikoseiId) {
        return kikikoseiId != null && kikimeisaiRepository.countByKikikoseiId(kikikoseiId) > 0;
    }

    // =================================================================
    // 3. 付替え先データ取得
    // 【変換元】TsukekaesakiButton_Click()
    //   元コード: Me.Fill(MCM_MA_KIKIKOSEI_SAKI, plantId)
    //             + MCM_MA_BRAND_KOSEI_FOR_COMBOTableAdapter.Fill(...)
    // =================================================================

    /**
     * 付替え先の機器構成データを取得する
     */
    public List<Mcm0021uSakiRowDto> findSakiByPlantId(BigDecimal plantId) {
        return mcm0021uRepository.findSakiByPlantId(plantId);
    }

    /**
     * 付替え先プラントのブランド構成リストを取得する（コンボ用）
     */
    public List<Object[]> findBrandKoseiForCombo(BigDecimal plantId) {
        return mcm0021uRepository.findBrandKoseiForCombo(plantId);
    }

    // =================================================================
    // 4. 付替え前バリデーション
    // 【変換元】ReplaceButton_Click() の前半チェック部分
    //   元コード: plantIdMoto=0チェック / plantIdSaki=0チェック /
    //             tsukekae_check(同値チェック) / IDO_CHECK+IKOSAKI整合性チェック
    // =================================================================

    /**
     * 付替え実行前のバリデーション
     * @return エラーメッセージリスト（空なら正常）
     */
    public List<String> validateReplace(BigDecimal plantIdMoto, BigDecimal plantIdSaki,
                                         List<MotoRowForm> motoRows) {
        List<String> errors = new ArrayList<>();

        /*
         * 【変換元】付替え元が設定されているか確認
         *   元コード: If Me.plantIdMoto = 0 Then DisplayMessage(MSG_0047)
         */
        if (plantIdMoto == null || plantIdMoto.compareTo(BigDecimal.ZERO) == 0) {
            errors.add(Mcm0021uConstants.MSG_MOTO_NOT_SET);
            return errors;
        }

        /*
         * 【変換元】付替え先が設定されているか確認
         *   元コード: If Me.plantIdSaki = 0 Then DisplayMessage(MSG_0048)
         */
        if (plantIdSaki == null || plantIdSaki.compareTo(BigDecimal.ZERO) == 0) {
            errors.add(Mcm0021uConstants.MSG_SAKI_NOT_SET);
            return errors;
        }

        /*
         * 【変換元】付替え元と付替え先の同値チェック
         *   元コード: tsukekae_check(plantIdMoto, plantIdSaki)
         */
        if (plantIdMoto.compareTo(plantIdSaki) == 0) {
            errors.add(Mcm0021uConstants.MSG_SAME_PLANT);
            return errors;
        }

        /*
         * 【変換元】チェックONかつ移行先未選択のチェック（#261対応）
         *   元コード: If IDO_CHECK = ON And IsNull(IKOSEI_MOTO) Then MSG_0086
         *   付替え先プラントの未設定と、行ごとのブランド未選択を区別する。
         */
        boolean hasChecked = false;
        if (motoRows != null) {
            for (MotoRowForm row : motoRows) {
                if (row == null) continue;
                if (row.isIdoCheck() && (row.getIkosaki() == null || row.getIkosaki().signum() <= 0)) {
                    errors.add(Mcm0021uConstants.MSG_IKOSAKI_NOT_SELECTED);
                    return errors;
                }
                if (row.isIdoCheck()) {
                    hasChecked = true;
                }
            }
        }

        /*
         * 【変換元】チェックが1件もないチェック
         *   元コード: If checkFlg = False Then DisplayMessage(MSG_0130)
         */
        if (!hasChecked) {
            errors.add(Mcm0021uConstants.MSG_NO_CHECK);
            return errors;
        }

        return errors;
    }

    // =================================================================
    // 5. 契約不整合チェック
    // 【変換元】ReplaceButton_Click() の契約チェック部分
    //   チェックOFF行とON行で同じ契約を共有していないか確認
    //   元コード: getTkKeiyakuList() / getUkKeiyakuList()
    //   tkKeiyakuOnList.IndexOf(tkKeiyakuOffList(j)) >= 0 → 不整合
    // =================================================================

    /**
     * 契約不整合をチェックする
     * @return true=不整合あり（警告表示が必要）, false=問題なし
     */
    public boolean checkKeiyakuConflict(List<MotoRowForm> motoRows) {
        if (motoRows == null || motoRows.isEmpty()) {
            return false;
        }

        // チェックOFF行を抽出
        List<MotoRowForm> offRows = motoRows.stream()
                .filter(r -> !r.isIdoCheck()).collect(Collectors.toList());

        // チェックOFF行が0件なら不整合なし（全行移動）
        if (offRows.isEmpty()) {
            return false;
        }

        // チェックON行を抽出
        List<MotoRowForm> onRows = motoRows.stream()
                .filter(MotoRowForm::isIdoCheck).collect(Collectors.toList());

        /*
         * 【変換元】取引先契約チェック
         *   元コード: tkKeiyakuOnList / tkKeiyakuOffList → IndexOfで交差判定
         */
        Set<BigDecimal> tkOn = new HashSet<>();
        for (MotoRowForm row : onRows) {
            tkOn.addAll(mcm0021uRepository.getTkKeiyakuList(
                    row.getBrandkoseiId(), row.getKikikoseiId()));
        }
        Set<BigDecimal> tkOff = new HashSet<>();
        for (MotoRowForm row : offRows) {
            tkOff.addAll(mcm0021uRepository.getTkKeiyakuList(
                    row.getBrandkoseiId(), row.getKikikoseiId()));
        }
        for (BigDecimal id : tkOn) {
            if (tkOff.contains(id)) {
                return true;
            }
        }

        /*
         * 【変換元】店舗契約チェック
         *   元コード: ukKeiyakuOnList / ukKeiyakuOffList → IndexOfで交差判定
         */
        Set<BigDecimal> ukOn = new HashSet<>();
        for (MotoRowForm row : onRows) {
            ukOn.addAll(mcm0021uRepository.getUkKeiyakuList(
                    row.getBrandkoseiId(), row.getKikikoseiId()));
        }
        Set<BigDecimal> ukOff = new HashSet<>();
        for (MotoRowForm row : offRows) {
            ukOff.addAll(mcm0021uRepository.getUkKeiyakuList(
                    row.getBrandkoseiId(), row.getKikikoseiId()));
        }
        for (BigDecimal id : ukOn) {
            if (ukOff.contains(id)) {
                return true;
            }
        }

        return false;
    }

    // =================================================================
    // 6. 付替え実行
    // 【変換元】ReplaceButton_Click() の後半更新部分
    //   元コード: 機器構成LOOPをFor文でループし、
    //     onCount=0: skip（Continue For）
    //     offCount=0: PLANT_ID置換 + BRANDKOSEI_ID置換
    //     一部ON: 新規行作成（機器構成+機器明細）+ 個体管理移動 + セット数調整
    //   UpdateAll: MCM_MA_KIKIKOSEI_UPDATE, MCM_MA_KIKIMEISAI_UPDATE,
    //              MCM_MA_KIKIKOTAIKANRI_UPDATE
    // =================================================================

    /**
     * 付替え処理を実行する
     * @param plantIdMoto 付替え元プラントID
     * @param plantIdSaki 付替え先プラントID
     * @param motoRows 付替え元グリッドデータ（チェック状態+移行先ブランド含む）
     * @param loginUser ログインユーザID
     * @return エラーメッセージ（正常時はnull）
     */
    private static String idKey(BigDecimal id) {
        return id == null ? "" : id.stripTrailingZeros().toPlainString();
    }
    private static String rowKey(BigDecimal composition, BigDecimal brand) {
        return idKey(composition) + "/" + idKey(brand);
    }
    private record MovePlan(KikikoseiEntity source, List<MotoRowForm> rows,
                            boolean whole, List<KikimeisaiEntity> details, Map<String,Integer> counts) {}

    @Transactional(isolation = org.springframework.transaction.annotation.Isolation.SERIALIZABLE)
    public String executeReplace(BigDecimal plantIdMoto, BigDecimal plantIdSaki,
                                  List<MotoRowForm> motoRows, String loginUser) {
        var errors = validateReplace(plantIdMoto, plantIdSaki, motoRows);
        if (!errors.isEmpty()) return errors.get(0);

        // 最新の元データと照合する。完了済みの再送・別プラントのID・行の欠落で更新しない。
        var liveRows = findMotoByPlantId(plantIdMoto);
        var liveKeys = liveRows.stream().map(r -> rowKey(r.getKikikoseiId(), r.getBrandkoseiId()))
                .collect(Collectors.toSet());
        var postedKeys = new HashSet<String>();
        for (var row : motoRows) {
            if (row == null || row.getKikikoseiId() == null || row.getBrandkoseiId() == null
                    || !postedKeys.add(rowKey(row.getKikikoseiId(), row.getBrandkoseiId()))
                    || !liveKeys.contains(rowKey(row.getKikikoseiId(), row.getBrandkoseiId())))
                return Mcm0021uConstants.MSG_TARGET_CHANGED;
        }
        var validBrands = findBrandKoseiForCombo(plantIdSaki).stream()
                .map(b -> idKey(new BigDecimal(b[0].toString()))).collect(Collectors.toSet());
        for (var row : motoRows) if (row.isIdoCheck() && !validBrands.contains(idKey(row.getIkosaki())))
            return Mcm0021uConstants.MSG_IKOSAKI_NOT_SELECTED;

        var compositions = kikikoseiRepository.findByPlantIdOrderByHyojijunAsc(plantIdMoto).stream()
                .collect(Collectors.toMap(r -> idKey(r.getKikikoseiId()), r -> r));
        var groups = motoRows.stream().collect(Collectors.groupingBy(r -> idKey(r.getKikikoseiId()),
                java.util.LinkedHashMap::new, Collectors.toList()));
        List<MovePlan> plans = new ArrayList<>();
        // すべての数量・明細を更新前に検査する。後続行のエラーで先行行だけ確定させない。
        for (var group : groups.entrySet()) {
            var rows = group.getValue();
            var checked = rows.stream().filter(MotoRowForm::isIdoCheck).toList();
            if (checked.isEmpty()) continue;
            var source = compositions.get(group.getKey());
            if (source == null) return Mcm0021uConstants.MSG_TARGET_CHANGED;
            long actualRows = liveRows.stream().filter(r -> idKey(r.getKikikoseiId()).equals(group.getKey())).count();
            if (actualRows != rows.size()) return Mcm0021uConstants.MSG_TARGET_CHANGED;
            var details = kikimeisaiRepository.findByKikikoseiIdOrderByHyojijunAsc(source.getKikikoseiId());
            if (details.isEmpty()) return Mcm0021uConstants.MSG_NO_KIKI_DATA;
            boolean whole = checked.size() == rows.size();
            Map<String,Integer> counts = new java.util.LinkedHashMap<>();
            if (!whole) {
                long total = 0;
                for (var row : checked) {
                    var maxima = mcm0021uRepository.findKotaimeisaiVDataMax(source.getKikikoseiId(), row.getBrandkoseiId());
                    Integer count = null;
                    for (var maximum : maxima) {
                        int value = ((Number) maximum[1]).intValue();
                        if (value <= 0 || (count != null && count != value))
                            return String.format(Mcm0021uConstants.MSG_SET_MISMATCH, row.getBrandsyosaiNk(), source.getKikikoseiNk());
                        count = value;
                    }
                    if (count == null) return Mcm0021uConstants.MSG_NO_KIKI_DATA;
                    counts.put(idKey(row.getBrandkoseiId()), count);
                    total += count;
                }
                if (source.getSetNm() == null || source.getSetNm().compareTo(BigDecimal.valueOf(total)) <= 0)
                    return String.format(Mcm0021uConstants.MSG_SET_MISMATCH, checked.get(0).getBrandsyosaiNk(), source.getKikikoseiNk());
            }
            plans.add(new MovePlan(source, checked, whole, details, counts));
        }

        BigDecimal maxKoseiId = BigDecimal.ZERO, maxKoseiNo = BigDecimal.ZERO, maxMeisaiId = BigDecimal.ZERO;
        if (plans.stream().anyMatch(p -> !p.whole())) {
            maxKoseiId = mcm0021uRepository.getMaxKikikoseiId();
            maxKoseiNo = mcm0021uRepository.getMaxKikikoseiHyojijun(plantIdSaki);
            maxMeisaiId = mcm0021uRepository.getMaxKikimeisaiId();
        }
        LocalDateTime now = LocalDateTime.now();
        for (var plan : plans) {
            var source = plan.source();
            BigDecimal sourceId = source.getKikikoseiId();
            if (plan.whole()) {
                // VBの全件移動：構成IDを維持して所属プラントを変更し、全対象個体のブランドも変更。
                if (kikikoseiRepository.updatePlantId(sourceId, plantIdSaki, loginUser, now) != 1)
                    throw new IllegalStateException(Mcm0021uConstants.MSG_TARGET_CHANGED);
                for (var row : plan.rows()) {
                    if (mcm0021uRepository.moveIndividuals(sourceId, row.getBrandkoseiId(),
                            row.getIkosaki(), sourceId, loginUser, now) <= 0)
                        throw new IllegalStateException(Mcm0021uConstants.MSG_TARGET_CHANGED);
                }
            } else {
                // VBの一部移動：移動ブランドごとに構成と全明細を複製し、選択した個体のみ移す。
                for (var row : plan.rows()) {
                    int count = plan.counts().get(idKey(row.getBrandkoseiId()));
                    maxKoseiId = maxKoseiId.add(BigDecimal.ONE);
                    maxKoseiNo = maxKoseiNo.add(BigDecimal.ONE);
                    var destination = new KikikoseiEntity();
                    org.springframework.beans.BeanUtils.copyProperties(source, destination);
                    destination.setPlantId(plantIdSaki); destination.setKikikoseiId(maxKoseiId);
                    destination.setSetNm(BigDecimal.valueOf(count)); destination.setHyojijun(maxKoseiNo);
                    destination.setCreatedDt(now); destination.setCreatedBy(loginUser);
                    destination.setLastupdateDt(now); destination.setLastupdateBy(loginUser);
                    kikikoseiRepository.saveAndFlush(destination);
                    BigDecimal order = BigDecimal.ZERO;
                    for (var original : plan.details()) {
                        var detail = new KikimeisaiEntity();
                        org.springframework.beans.BeanUtils.copyProperties(original, detail);
                        maxMeisaiId = maxMeisaiId.add(BigDecimal.ONE); order = order.add(BigDecimal.ONE);
                        detail.setKikimeisaiId(maxMeisaiId); detail.setKikikoseiId(maxKoseiId); detail.setHyojijun(order);
                        detail.setCreatedDt(now); detail.setCreatedBy(loginUser);
                        detail.setLastupdateDt(now); detail.setLastupdateBy(loginUser);
                        kikimeisaiRepository.save(detail);
                    }
                    kikimeisaiRepository.flush();
                    if (mcm0021uRepository.moveIndividuals(sourceId, row.getBrandkoseiId(),
                            row.getIkosaki(), maxKoseiId, loginUser, now) <= 0)
                        throw new IllegalStateException(Mcm0021uConstants.MSG_TARGET_CHANGED);
                    if (kikikoseiRepository.decrementSetNm(sourceId, BigDecimal.valueOf(count), loginUser, now) != 1)
                        throw new IllegalStateException(Mcm0021uConstants.MSG_TARGET_CHANGED);
                }
            }
        }
        log.info("プラント付替え完了: moto={} -> saki={}, user={}", plantIdMoto, plantIdSaki, loginUser);
        return null;
    }
}
