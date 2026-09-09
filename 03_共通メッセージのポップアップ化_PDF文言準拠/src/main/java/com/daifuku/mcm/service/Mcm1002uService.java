package com.daifuku.mcm.service;

/**
 * 【変換元】Mcm1002u1Screen.vb + Mcm1002u2Screen.vb
 *   MCM1002U 需要家見積作成 Service
 *
 *   元コード: Mcm1002u2Screen_Load(), updateMethod(), changeJotai(),
 *            afterUpdate(), excelOutput()
 */

import com.daifuku.mcm.common.Mcm1002uConstants;
import com.daifuku.mcm.entity.*;
import com.daifuku.mcm.form.*;
import com.daifuku.mcm.form.Mcm1002u1Form.*;
import com.daifuku.mcm.form.Mcm1002u2Form.*;
import com.daifuku.mcm.repository.*;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Slf4j
public class Mcm1002uService {

    private final McmTmMitsumoriRepository mitsumoriRepo;
    private final McmTmKeiyakujikanRepository keiyakujikanRepo;
    private final McmTmKikikoseiRepository kikikoseiRepo;
    private final McmTmKikimeisaiRepository kikimeisaiRepo;
    private final McmTmKotaimeisaiRepository kotaimeisaiRepo;
    private final McmTmKikanRepository kikanRepo;
    private final McmTmTankaRepository tankaRepo;
    private final McmTmTenkenRepository tenkenRepo;

    // ================================================================
    // Screen1: 機器構成パターン取得（チェックフラグ付き）
    // 【変換元】Mcm1002u1DataSet - MCM_MA_KIKIKOSEITableAdapter.Fill
    //   元SQL: SELECT MAE.KIKIKOSEI_ID, ... NVL(TKV.TM_KIKIKOSEI_ID, 0),
    //          DECODE(NVL(TKV.TM_KIKIKOSEI_ID, 0), 0, 0, 1) AS CHECK_FLG
    //          FROM MCM_MA_KIKIKOSEI MAE LEFT JOIN (subquery) TKV ...
    // ================================================================
    @Transactional(readOnly = true)
    public List<KikikoseiRow> loadKikikoseiWithCheckFlags(Long plantId, Long tmKeiyakujikanId) {
        // TODO: Native Query or custom Repository method
        // SQL Server版: ISNULL → NVL, CASE WHEN → DECODE
        log.info("loadKikikoseiWithCheckFlags: plantId={}, tmKeiyakujikanId={}", plantId, tmKeiyakujikanId);
        return new ArrayList<>();
    }

    // ================================================================
    // Screen1: 機器明細取得
    // 【変換元】MCM_MA_KIKIMEISAITableAdapter.Fill
    //   元SQL: SELECT MAF.KIKIMEISAI_ID, ... FROM MCM_MA_KIKIMEISAI MAF,
    //          MCM_MA_ATSUKAIKIKI MAH, MCM_MA_KIKIBUNRUI MAM ...
    // ================================================================
    @Transactional(readOnly = true)
    public List<KikimeisaiRow> loadKikimeisai() {
        log.info("loadKikimeisai");
        return new ArrayList<>();
    }

    // ================================================================
    // Screen1: 個体明細取得
    // 【変換元】MCM_MA_KIKIKOTAIMEISAITableAdapter.Fill
    //   元SQL: SELECT UVB.* FROM MCM_MA_KOTAIMEISAI_V UVB LEFT JOIN ...
    // ================================================================
    @Transactional(readOnly = true)
    public List<KotaimeisaiRow> loadKotaimeisai(Long plantId, Long tmKeiyakujikanId) {
        log.info("loadKotaimeisai: plantId={}, tmKeiyakujikanId={}", plantId, tmKeiyakujikanId);
        return new ArrayList<>();
    }

    // ================================================================
    // Screen2: Form構築
    // 【変換元】Mcm1002u2Screen_Load
    // ================================================================
    public Mcm1002u2Form buildScreen2Form(Mcm1002u1Form screen1Form,
                                           Mcm1002uDeliveryDto delivery) {
        Mcm1002u2Form form = new Mcm1002u2Form();
        if (delivery != null) {
            form.setNonyusakiCd(delivery.getNonyusakiCd());
            form.setNonyusakiNk(delivery.getNonyusakiNk());
            form.setSupportId(delivery.getSupportId());
            form.setPlantNk(delivery.getPlantNk());
        }

        // 既存データがある場合（修正モード）
        if (delivery != null && delivery.getTmKeiyakujikanId() != null
                && delivery.getTmKeiyakujikanId() > 0) {
            List<McmTmMitsumoriEntity> existing =
                mitsumoriRepo.findByKeiyakujikanId(delivery.getTmKeiyakujikanId());
            for (McmTmMitsumoriEntity e : existing) {
                MitsumoriRow row = new MitsumoriRow();
                row.setTmIraiId(e.getTmIraiId());
                row.setTmIraiNo(e.getTmIraiNo());
                row.setTorihikisakiId(e.getTorihikisakiId());
                row.setTorihikisakiCd(e.getTorihikisakiCd());
                row.setTorihikisakiNk(e.getTorihikisakiNk());
                row.setH8("1".equals(e.getH8()));
                row.setH24("1".equals(e.getH24()));
                row.setTenkenumu("1".equals(e.getTenkenumu()));
                row.setTenkenkanoyobi(e.getTenkenkanoyobi());
                row.setYakantaioumu("1".equals(e.getYakantaioumu()));
                row.setHosyuhoho(e.getHosyuhoho());
                row.setPackFlg(BigDecimal.ONE.equals(e.getPackFlg()));
                row.setTenpoFlg(BigDecimal.ONE.equals(e.getTenpoFlg()));
                form.getMitsumoriRows().add(row);
            }
        }
        return form;
    }

    // ================================================================
    // バリデーション: 8H / 24H チェック
    // 【変換元】updateMethod() 内の時間帯チェック
    //   元コード: If 8H=0 And 24H=0 → MSG_0024
    //            If 8H=1 And 24H=1 → MSG_0145
    // ================================================================
    public List<String> validate8H24H(Mcm1002u2Form form) {
        List<String> errors = new ArrayList<>();
        for (int i = 0; i < form.getMitsumoriRows().size(); i++) {
            MitsumoriRow row = form.getMitsumoriRows().get(i);
            if (!row.isH8() && !row.isH24()) {
                errors.add("保守契約時間帯を選択してください。");
            }
            if (row.isH8() && row.isH24()) {
                errors.add("行" + (i + 1) + ": 保守契約時間帯は8Hまたは24Hのいずれか一方を選択してください");
            }
        }
        return errors;
    }

    // ================================================================
    // 登録処理（メイン）
    // 【変換元】Mcm1002u2Screen.updateMethod()
    //   元コード: 約300行の巨大メソッド
    //   ・契約時間の削除/追加
    //   ・ID採番（MAX+1）→ JPA IDENTITYで自動化
    //   ・親子ID付替え → Cascadeで自動化
    //   ・状態変更（依頼書発行時）
    // ================================================================
    @Transactional
    public void saveAll(Mcm1002u2Form form,
                        Mcm1002uDeliveryDto delivery,
                        String loginUser,
                        boolean printFlg) {

        int mitsumoriFlg = delivery != null ? delivery.getMitsumoriFlg() : Mcm1002uConstants.SHINKI;
        String jotai = delivery != null ? delivery.getJotai() : null;
        LocalDateTime now = LocalDateTime.now();

        // 修正モードの状態チェック
        if (mitsumoriFlg == Mcm1002uConstants.SYUSEI) {
            if (Mcm1002uConstants.JOTAI_KAIYAKU_NK.equals(jotai)
                || Mcm1002uConstants.JOTAI_KEIYAKU_NK.equals(jotai)
                || Mcm1002uConstants.JOTAI_HAKI_NK.equals(jotai)) {
                throw new IllegalStateException(
                    "状態が「" + jotai + "」のため登録できません");
            }
        }

        // 見積行ごとに処理
        for (MitsumoriRow row : form.getMitsumoriRows()) {
            McmTmMitsumoriEntity mitsumori;

            if (row.getTmIraiId() != null && row.getTmIraiId() > 0) {
                // 既存データ更新
                mitsumori = mitsumoriRepo.findById(row.getTmIraiId())
                    .orElseThrow(() -> new RuntimeException("見積データが見つかりません: " + row.getTmIraiId()));
            } else {
                // 新規作成
                mitsumori = new McmTmMitsumoriEntity();
                mitsumori.setTmIraiNo(generateIraiNo());
                mitsumori.setCreatedDt(now);
                mitsumori.setCreatedBy(loginUser);
                mitsumori.setMitsumoriDt(now);
                mitsumori.setIraitantosya(loginUser);
            }

            // Form → Entity マッピング
            mitsumori.setH8(row.isH8() ? "1" : "0");
            mitsumori.setH24(row.isH24() ? "1" : "0");
            mitsumori.setTenkenumu(row.isTenkenumu() ? "1" : "0");
            mitsumori.setTenkenkanoyobi(row.getTenkenkanoyobi());
            mitsumori.setYakantaioumu(row.isYakantaioumu() ? "1" : "0");
            mitsumori.setHosyuhoho(row.getHosyuhoho());
            mitsumori.setPackFlg(row.isPackFlg() ? BigDecimal.ONE : BigDecimal.ZERO);
            mitsumori.setTenpoFlg(row.isTenpoFlg() ? BigDecimal.ONE : BigDecimal.ZERO);
            mitsumori.setLastupdateDt(now);
            mitsumori.setLastupdateBy(loginUser);

            // -----------------------------------------------
            // 契約時間帯の同期（8H / 24H）
            // 【変換元】updateMethod() 内の契約時間DataTable削除/追加
            // -----------------------------------------------
            syncKeiyakujikan(mitsumori, row, now, loginUser, printFlg, jotai);

            // 保存（Cascade で子テーブルも一括保存）
            mitsumoriRepo.save(mitsumori);
        }
    }

    /**
     * 契約時間帯の同期処理
     * 【変換元】updateMethod() 内の契約時間DataTableの行削除/追加ロジック
     */
    private void syncKeiyakujikan(McmTmMitsumoriEntity mitsumori,
                                   MitsumoriRow row,
                                   LocalDateTime now,
                                   String loginUser,
                                   boolean printFlg,
                                   String jotai) {

        List<McmTmKeiyakujikanEntity> existing = mitsumori.getKeiyakujikanList();
        Map<String, McmTmKeiyakujikanEntity> byJikantai = existing.stream()
            .collect(Collectors.toMap(
                McmTmKeiyakujikanEntity::getKeiyakujikantai,
                e -> e,
                (a, b) -> a));

        // 8H 処理
        if (row.isH8()) {
            if (!byJikantai.containsKey(Mcm1002uConstants.JIKAN_8)) {
                McmTmKeiyakujikanEntity newK = createKeiyakujikan(
                    mitsumori, Mcm1002uConstants.JIKAN_8, now, loginUser);
                existing.add(newK);
            }
        } else {
            existing.removeIf(k -> Mcm1002uConstants.JIKAN_8.equals(k.getKeiyakujikantai()));
        }

        // 24H 処理
        if (row.isH24()) {
            if (!byJikantai.containsKey(Mcm1002uConstants.JIKAN_24)) {
                McmTmKeiyakujikanEntity newK = createKeiyakujikan(
                    mitsumori, Mcm1002uConstants.JIKAN_24, now, loginUser);
                existing.add(newK);
            }
        } else {
            existing.removeIf(k -> Mcm1002uConstants.JIKAN_24.equals(k.getKeiyakujikantai()));
        }

        // 依頼書発行時の状態変更
        if (printFlg && Mcm1002uConstants.JOTAI_SAKUSEICHU_NK.equals(jotai)) {
            changeJotai(existing);
        }
    }

    private McmTmKeiyakujikanEntity createKeiyakujikan(
            McmTmMitsumoriEntity mitsumori, String jikantai,
            LocalDateTime now, String loginUser) {
        McmTmKeiyakujikanEntity k = new McmTmKeiyakujikanEntity();
        k.setMitsumori(mitsumori);
        k.setKeiyakujikantai(jikantai);
        k.setJotai(Mcm1002uConstants.JOTAI_SAKUSEICHU);
        k.setIns(Mcm1002uConstants.INS_NEW);
        k.setCreatedDt(now);
        k.setCreatedBy(loginUser);
        return k;
    }

    // ================================================================
    // 状態変更：作成中 → 依頼
    // 【変換元】Mcm1002u2Screen.changeJotai()
    //   元コード: keiyakujikanRow.JOTAI = McmConstant.JOTAI_IRAI
    // ================================================================
    public void changeJotai(List<McmTmKeiyakujikanEntity> keiyakujikanList) {
        for (McmTmKeiyakujikanEntity k : keiyakujikanList) {
            if (Mcm1002uConstants.JOTAI_SAKUSEICHU.equals(k.getJotai())) {
                k.setJotai(Mcm1002uConstants.JOTAI_IRAI);
            }
        }
    }

    // ================================================================
    // 依頼NO生成
    // 【変換元】updateMethod() 内の iraiNoStr 生成
    //   元コード: Mcm1002uConstant.IRAI_NO_Pre & Format(Now, "YYYYMM") & 連番3桁
    // ================================================================
    public String generateIraiNo() {
        String prefix = Mcm1002uConstants.IRAI_NO_PREFIX
            + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMM"));
        Integer maxNo = mitsumoriRepo.findMaxIraiNoByPrefix(prefix);
        int nextNo = (maxNo != null ? maxNo : 0) + 1;
        return prefix + String.format("%03d", nextNo);
    }

    // ================================================================
    // ドロップダウン用マスタ
    // ================================================================
    public Map<String, String> getHosyuhohoOptions() {
        Map<String, String> map = new LinkedHashMap<>();
        map.put("", "");
        map.put("01", "フルメンテナンス");
        map.put("02", "POG");
        // TODO: MCM_MA_CODEから取得に変更
        return map;
    }

    public Map<String, String> getTenkenkanoyobiOptions() {
        Map<String, String> map = new LinkedHashMap<>();
        map.put("1", "平日");
        map.put("2", "土日祝");
        map.put("3", "指定なし");
        // TODO: MCM_MA_CODEから取得に変更
        return map;
    }

    // ================================================================
    // Excel帳票出力
    // 【変換元】Mcm1002u2Screen.excelOutput() → Mcm1001pExcel
    // ================================================================
    // TODO: Apache POI で実装
    public void generateExcelReport(Long tmIraiId) {
        log.info("TODO: Excel帳票出力 tmIraiId={}", tmIraiId);
    }
}