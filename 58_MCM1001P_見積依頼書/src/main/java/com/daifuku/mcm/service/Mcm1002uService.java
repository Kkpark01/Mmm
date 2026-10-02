package com.daifuku.mcm.service;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.stream.Collectors;

import org.springframework.context.MessageSource;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

/**
 * 【変換元】Mcm1002u1Screen.vb + Mcm1002u2Screen.vb
 *   MCM1002U 需要家見積作成 Service
 *
 *   元コード: Mcm1002u2Screen_Load(), updateMethod(), changeJotai(),
 *            afterUpdate(), excelOutput()
 */

import com.daifuku.mcm.common.Mcm1002uConstants;
import com.daifuku.mcm.dto.Mcm1001pDeliveryDto;
import com.daifuku.mcm.dto.Mcm1001pHeaderDto;
import com.daifuku.mcm.dto.Mcm1001pViewDto;
import com.daifuku.mcm.entity.McmTmKeiyakujikanEntity;
import com.daifuku.mcm.entity.McmTmKikikoseiEntity;
import com.daifuku.mcm.entity.McmTmKikimeisaiEntity;
import com.daifuku.mcm.entity.McmTmKotaimeisaiEntity;
import com.daifuku.mcm.entity.McmTmMitsumoriEntity;
import com.daifuku.mcm.form.Mcm1002u1Form;
import com.daifuku.mcm.form.Mcm1002u1Form.KotaimeisaiRow;
import com.daifuku.mcm.form.Mcm1002u2Form;
import com.daifuku.mcm.form.Mcm1002u2Form.MitsumoriRow;
import com.daifuku.mcm.form.Mcm1002uDeliveryDto;
import com.daifuku.mcm.repository.Mcm1002uRepository;
import com.daifuku.mcm.repository.McmTmKeiyakujikanRepository;
import com.daifuku.mcm.repository.McmTmKikanRepository;
import com.daifuku.mcm.repository.McmTmKikikoseiRepository;
import com.daifuku.mcm.repository.McmTmKikimeisaiRepository;
import com.daifuku.mcm.repository.McmTmKotaimeisaiRepository;
import com.daifuku.mcm.repository.McmTmMitsumoriRepository;
import com.daifuku.mcm.repository.McmTmTankaRepository;
import com.daifuku.mcm.repository.McmTmTenkenRepository;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;

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
    private final Mcm1002uRepository mcm1002uRepo;
    private final Mcm2004uService permissions;
    private final MessageSource messageSource;

    /** 「取引先見積関連」の作成権限で登録を許可する。依頼書の取得は参照権限でも可能。 */
    public boolean canUpdate(String user) {
        return "2".equals(permissions.getAuthority(user, "MCM1002U"));
    }

    /**
     * plantIdをキーにMCM_MA_PLANT/MCM_MA_NONYUSAKIマスタを結合検索し、
     * 見積依頼ヘッダ（納入先・プラント関連項目）を取得する。
     * 【変換元】Mcm1002u1Screen_Load - MCM_TM_MITSUMORITableAdapter.Fill()
     */
    @Transactional(readOnly = true)
    public Map<String, Object> findMitsumoriHeader(BigDecimal plantId) {
        if (plantId == null) return null;
        return mcm1002uRepo.findMitsumoriHeaderByPlantId(plantId);
    }

    // ================================================================
    // Screen1: 機器構成パターン取得（チェックフラグ付き）
    // 【変換元】Mcm1002u1DataSet - MCM_MA_KIKIKOSEITableAdapter.Fill
    //   元SQL: SELECT MAE.KIKIKOSEI_ID, ... NVL(TKV.TM_KIKIKOSEI_ID, 0),
    //          DECODE(NVL(TKV.TM_KIKIKOSEI_ID, 0), 0, 0, 1) AS CHECK_FLG
    //          FROM MCM_MA_KIKIKOSEI MAE LEFT JOIN (subquery) TKV ...
    // ================================================================
    @Transactional(readOnly = true)
    public List<Mcm1002u1Form.KikikoseiRow> loadKikikoseiWithCheckFlags(Long plantId, Long tmKeiyakujikanId) {
        log.info("loadKikikoseiWithCheckFlags: plantId={}, tmKeiyakujikanId={}", plantId, tmKeiyakujikanId);
        if (plantId == null) return new ArrayList<>();
        BigDecimal plant = BigDecimal.valueOf(plantId);
        BigDecimal keiyakujikan = tmKeiyakujikanId != null ? BigDecimal.valueOf(tmKeiyakujikanId) : null;
        return mcm1002uRepo.findKikikosei(plant, keiyakujikan);
    }

    // ================================================================
    // Screen1: 機器明細取得
    // 【変換元】MCM_MA_KIKIMEISAITableAdapter.Fill
    //   元SQL: SELECT MAF.KIKIMEISAI_ID, ... FROM MCM_MA_KIKIMEISAI MAF,
    //          MCM_MA_ATSUKAIKIKI MAH, MCM_MA_KIKIBUNRUI MAM ...
    // ================================================================
    @Transactional(readOnly = true)
    public List<Mcm1002u1Form.KikimeisaiRow> loadKikimeisai(Long plantId) {
        log.info("loadKikimeisai: plantId={}", plantId);
        if (plantId == null) return new ArrayList<>();
        return mcm1002uRepo.findKikimeisai(BigDecimal.valueOf(plantId));
    }

    // ================================================================
    // Screen1: 個体明細取得
    // 【変換元】MCM_MA_KIKIKOTAIMEISAITableAdapter.Fill
    //   元SQL: SELECT UVB.* FROM MCM_MA_KOTAIMEISAI_V UVB LEFT JOIN ...
    // ================================================================
    @Transactional(readOnly = true)
    public List<KotaimeisaiRow> loadKotaimeisai(Long plantId, Long tmKeiyakujikanId) {
        log.info("loadKotaimeisai: plantId={}, tmKeiyakujikanId={}", plantId, tmKeiyakujikanId);
        if (plantId == null) return new ArrayList<>();
        BigDecimal plant = BigDecimal.valueOf(plantId);
        BigDecimal keiyakujikan = tmKeiyakujikanId != null ? BigDecimal.valueOf(tmKeiyakujikanId) : null;
        return mcm1002uRepo.findKotaimeisai(plant, keiyakujikan);
    }

    // ================================================================
    // Screen1 → Screen2: 取引先ごとの見積依頼行を組み立てる
    // 【変換元】Mcm1002u1Screen.vb - NextButton_Click()
    //   元コード: 個体明細でCHECK_FLG=ONの行をループし、取引先ID単位で
    //            Mcm1002u2Form.MitsumoriRow を生成する（配列組立の簡略移植）。
    //   Web版差異: VB版はID採番（負数の仮ID）や既存TM_IRAI_ID引継ぎを含む複雑な
    //   配列操作を行うが、Web版ではDB永続化はsaveAll()側でJPA Cascadeにより
    //   行われるため、ここでは「取引先単位のグルーピング」のみを担う。
    // ================================================================
    public List<MitsumoriRow> buildMitsumoriRowsFromSelection(Mcm1002u1Form screen1Form) {
        List<MitsumoriRow> result = new ArrayList<>();
        if (screen1Form == null || screen1Form.getKotaimeisaiRows() == null) {
            return result;
        }

        // チェック済みの個体明細を取引先ID単位でグルーピング
        LinkedHashMap<String, KotaimeisaiRow> firstRowByTorihikisaki = new LinkedHashMap<>();
        for (KotaimeisaiRow row : screen1Form.getKotaimeisaiRows()) {
            if (!row.isChecked()) {
                continue;
            }
            String key = row.getTorihikisakiId() != null ? row.getTorihikisakiId().toString() : "";
            firstRowByTorihikisaki.putIfAbsent(key, row);
        }

        for (KotaimeisaiRow row : firstRowByTorihikisaki.values()) {
            MitsumoriRow mr = new MitsumoriRow();
            mr.setTorihikisakiId(row.getTorihikisakiId());
            mr.setTorihikisakiCd(row.getTorihikisakiCd());
            mr.setTorihikisakiNk(row.getTorihikisakiNk());
            // 【不一致修正】元VB: Mcm1002u1Screen.vb NextButton_Click() 取引先Array格納部
            //   torihikisaki(TMA_TORITEL_NO/TMA_TORIFAX_NO/TMA_TORIJIGYOSYO_NK/...)は
            //   個体明細マスタ結合結果(currentKotaiRow)由来。従来コピーされておらず
            //   移植漏れとなっていたため追加。
            mr.setToritelNo(row.getToritelNo());
            mr.setTorifaxNo(row.getTorifaxNo());
            mr.setTorijigyosyoNk(row.getTorijigyosyoNk());
            mr.setTorisyutantosyaNk(row.getTorisyutantosyaNk());
            mr.setToriassistantNk(row.getToriassistantNk());

            // 【不一致修正】元VB: Mcm1002u1Screen.vb NextButton_Click() 取引先Array格納部
            //   torihikisaki(TMA_NONYUSAKI_ID/TMA_PLANT_ID 等)は、plantIdをキーにした
            //   マスタ検索結果（currentMitsumoriRow = Screen1のMCM_TM_MITSUMORIヘッダ）由来。
            //   これがMCM_TM_MITSUMORI新規行の必須項目であり、従来一切コピーされておらず
            //   PLANT_ID等がNULLで保存される不具合の原因になっていたため追加。
            if (screen1Form != null) {
                mr.setNonyusakiId(screen1Form.getNonyusakiId());
                mr.setNonyusakiCd(screen1Form.getNonyusakiCd());
                mr.setNonyusakiNk(screen1Form.getNonyusakiNk());
                mr.setNonyusakijusyo1Nk(screen1Form.getNonyusakijusyo1Nk());
                mr.setNonyusakijusyo2Nk(screen1Form.getNonyusakijusyo2Nk());
                mr.setPlantId(screen1Form.getPlantId());
                mr.setSupportId(screen1Form.getSupportId());
                mr.setPlantNk(screen1Form.getPlantNk());
                mr.setNonyubusyoNk(screen1Form.getNonyubusyoNk());
                mr.setNonyutantosyaNk(screen1Form.getNonyutantosyaNk());
                mr.setNonyutelNo(screen1Form.getNonyutelNo());
                mr.setNonyufaxNo(screen1Form.getNonyufaxNo());
            }

            // 【不一致修正】元VB: Mcm1002u2Screen.vb 見積テーブル追加設定部
            //   mitsumoriNewRow.H8 = 1（チェックON）
            //   mitsumoriNewRow.H24 = 0（OFF）
            //   mitsumoriNewRow.TENKENKANOYOBI = 1（「月～金」）
            //   がJavaのフィールドデフォルト(false/null)のままで移植漏れしていたため追加。
            mr.setH8(true);
            mr.setH24(false);
            mr.setTenkenkanoyobi("1");
            result.add(mr);
        }
        return result;
    }

    // ================================================================
    // Screen2: Form構築
    // 【変換元】Mcm1002u2Screen_Load
    // 【不一致修正】@Transactionalが無いと、Repositoryから返却されたEntityの
    //   遅延ロードコレクション(keiyakujikanList等)にコントローラー側でアクセスした際、
    //   セッションが既に閉じているためLazyInitializationExceptionが発生する。
    // ================================================================
    @Transactional(readOnly = true)
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
                row.setToritelNo(e.getToritelNo());
                row.setTorifaxNo(e.getTorifaxNo());
                row.setTorijigyosyoNk(e.getTorijigyosyoNk());
                row.setTorisyutantosyaNk(e.getTorisyutantosyaNk());
                row.setToriassistantNk(e.getToriassistantNk());
                row.setNonyusakiId(e.getNonyusakiId());
                row.setNonyusakiCd(e.getNonyusakiCd());
                row.setNonyusakiNk(e.getNonyusakiNk());
                row.setNonyusakijusyo1Nk(e.getNonyusakijusyo1Nk());
                row.setNonyusakijusyo2Nk(e.getNonyusakijusyo2Nk());
                row.setPlantId(e.getPlantId());
                row.setSupportId(e.getSupportId());
                row.setPlantNk(e.getPlantNk());
                row.setNonyubusyoNk(e.getNonyubusyoNk());
                row.setNonyutantosyaNk(e.getNonyutantosyaNk());
                row.setNonyutelNo(e.getNonyutelNo());
                row.setNonyufaxNo(e.getNonyufaxNo());
                // 【不一致修正】H8/H24はMCM_TM_MITSUMORIの実カラムではなく、
                //   MCM_TM_KEIYAKUJIKAN.KEIYAKUJIKANTAI('8'/'24')の行の存在有無から判定する。
                boolean has8h = e.getKeiyakujikanList().stream()
                    .anyMatch(k -> Mcm1002uConstants.JIKAN_8.equals(k.getKeiyakujikantai()));
                boolean has24h = e.getKeiyakujikanList().stream()
                    .anyMatch(k -> Mcm1002uConstants.JIKAN_24.equals(k.getKeiyakujikantai()));
                row.setH8(has8h);
                row.setH24(has24h);
                row.setTenkenumu("1".equals(e.getTenkenumu()));
                row.setTenkenkanoyobi(e.getTenkenkanoyobi());
                row.setYakantaioumu("1".equals(e.getYakantaioumu()));
                row.setHosyuhoho(e.getHosyuhoho());
                row.setPackFlg(BigDecimal.ONE.equals(e.getPackFlg()));
                row.setTenpoFlg(BigDecimal.ONE.equals(e.getTenpoFlg()));

                /*
                 * 【#471不一致修正】元VB: Mcm1002u2Screen_Load() 複製処理
                 *   If mitsumoriFlg = FUKUSEI Then 見積データテーブルを新規に置き換える
                 *     mitsumoriRow.TM_IRAI_NO = "" / INS = INS_NEW
                 *     MITSUMORI_DT = システム日付 / IRAITANTOSYA = ログインユーザー
                 *     SetKAITOKIZITSU_DTNull()
                 *   従来は複製元のTM_IRAI_IDをそのまま保持していたため、登録時に
                 *   複製元の見積を上書き更新してしまっていた。複製時はIDを持たない
                 *   新規行として扱い、saveAll()で新規採番させる
                 *   （見積日・依頼担当者はsaveAll()の新規作成分岐で設定される）。
                 */
                if (delivery.getMitsumoriFlg() == Mcm1002uConstants.FUKUSEI) {
                    row.setTmIraiId(null);
                    row.setTmIraiNo("");
                    row.setKaitokizitsuDt(null);
                }
                form.getMitsumoriRows().add(row);
            }

            // 【追加】機器構成パターン・機器構成明細（修正モード：DBの既存登録内容を表示）
            // 【変換元】MCM_TM_KIKIKOSEITableAdapter.Fill / MCM_TM_KIKIMEISAITableAdapter.Fill
            List<McmTmKikikoseiEntity> kikikoseiEntities =
                kikikoseiRepo.findByKeiyakujikanId(delivery.getTmKeiyakujikanId());
            for (McmTmKikikoseiEntity ke : kikikoseiEntities) {
                // 【注意】画面側でのJS文字列比較(絞り込み表示)がスケール差異でずれないよう正規化する。
                BigDecimal normalizedKoseiId = ke.getKikikoseiId() != null
                    ? ke.getKikikoseiId().stripTrailingZeros() : null;
                Mcm1002u2Form.KikikoseiRow kr = new Mcm1002u2Form.KikikoseiRow();
                kr.setKikikoseiId(normalizedKoseiId);
                kr.setKikikoseiNk(ke.getKikikoseiNk());
                kr.setTehaiseiban(ke.getTehaiseiban());
                form.getKikikoseiRows().add(kr);

                for (McmTmKikimeisaiEntity me : kikimeisaiRepo
                        .findByKikikosei_TmKikikoseiIdOrderByHyojijun(ke.getTmKikikoseiId())) {
                    Mcm1002u2Form.KikimeisaiRow mr = new Mcm1002u2Form.KikimeisaiRow();
                    mr.setKikikoseiId(normalizedKoseiId);
                    mr.setSeizomakerNk(me.getSeizomakerNk());
                    mr.setKikihinmeiNk(me.getKikihinmeiNk());
                    mr.setKikikatashiki(me.getKikikatashiki());
                    mr.setSuryoNm(me.getSuryoNm());
                    form.getKikimeisaiRows().add(mr);
                }
            }

            // 【不一致修正】既存契約(TM_KEIYAKUJIKAN_ID)はあるが、DB側にまだ機器構成が
            //   1件も登録されていない場合（Screen1で今回初めて機器構成を選んで
            //   「次へ」を押したケース）は、DBの検索結果が0件のままになってしまう。
            //   その場合はScreen1で選択された内容にフォールバックする。
            if (form.getKikikoseiRows().isEmpty() && screen1Form != null) {
                appendKikikoseiFromScreen1(form, screen1Form);
            }
        } else if (screen1Form != null) {
            // 新規作成モード：Screen1で選択された個体明細から取引先単位の
            // 見積依頼行を新規に組み立てる
            // 【変換元】Mcm1002u1Screen.vb - NextButton_Click() 取引先Array格納部
            form.getMitsumoriRows().addAll(buildMitsumoriRowsFromSelection(screen1Form));

            // 【追加】機器構成パターン・機器構成明細（新規作成モード：Screen1で
            //   チェックされた行をそのまま表示）
            appendKikikoseiFromScreen1(form, screen1Form);
        }
        return form;
    }

    /** 参照／確定済みの発行後も、disabledで送信されない項目を登録済みの値で表示する。 */
    @Transactional(readOnly = true)
    public void restoreSavedReportRows(Mcm1002u2Form form) {
        for (MitsumoriRow row : form.getMitsumoriRows()) {
            if (row.getTmIraiId() == null || row.getTmIraiId() <= 0) continue;
            McmTmMitsumoriEntity m = mitsumoriRepo.findById(row.getTmIraiId())
                    .orElseThrow(() -> new IllegalArgumentException("見積依頼の出力対象が見つかりません。"));
            var hours = keiyakujikanRepo.findByMitsumori_TmIraiId(row.getTmIraiId());
            row.setTmIraiNo(m.getTmIraiNo());
            row.setH8(hours.stream().anyMatch(k -> Mcm1002uConstants.JIKAN_8.equals(k.getKeiyakujikantai())));
            row.setH24(hours.stream().anyMatch(k -> Mcm1002uConstants.JIKAN_24.equals(k.getKeiyakujikantai())));
            row.setKaitokizitsuDt(formatDate(m.getKaitokizitsuDt()));
            row.setTenkenumu("1".equals(m.getTenkenumu()));
            row.setYakantaioumu("1".equals(m.getYakantaioumu()));
            row.setHosyuhoho(m.getHosyuhoho());
            row.setTenkenkanoyobi(m.getTenkenkanoyobi());
            row.setPackFlg(m.getPackFlg() != null && m.getPackFlg().compareTo(BigDecimal.ONE) == 0);
            row.setTenpoFlg(m.getTenpoFlg() != null && m.getTenpoFlg().compareTo(BigDecimal.ONE) == 0);
        }
    }

    /** POST対象ではない上部ヘッダー・表示専用の機器表を復元し、編集行はそのまま保持する。 */
    @Transactional(readOnly = true)
    public void restoreReportDisplay(Mcm1002u2Form form, Mcm1002u1Form screen1Form,
                                     Mcm1002uDeliveryDto delivery) {
        Mcm1002u2Form display = buildScreen2Form(screen1Form, delivery);
        if (form.getNonyusakiCd() == null) form.setNonyusakiCd(display.getNonyusakiCd());
        if (form.getNonyusakiNk() == null) form.setNonyusakiNk(display.getNonyusakiNk());
        if (form.getSupportId() == null) form.setSupportId(display.getSupportId());
        if (form.getPlantNk() == null) form.setPlantNk(display.getPlantNk());
        form.setKikikoseiRows(display.getKikikoseiRows());
        form.setKikimeisaiRows(display.getKikimeisaiRows());
    }

    // ================================================================
    // 【#471】期間情報複写あり時の8H/24H値の固定
    // 【変換元】Mcm1002u2Screen_Load()
    //   If Me.copyFlg Then
    //     KEIYAKUJIKANTAI8H_MITSUMORI_CheckBox.ReadOnly = True
    //     KEIYAKUJIKANTAI24H_MITSUMORI_CheckBox.ReadOnly = True
    //   Web版は画面側で非活性にするが、送信値の改ざんに備え、登録・依頼書発行時に
    //   画面表示時と同じ組み立て（buildScreen2Form）で得た値へ上書きする。
    //   copyFlg=false（通常新規・修正・複写なし複製・MCM1008U経由）の場合は何もしない。
    // ================================================================
    @Transactional(readOnly = true)
    public void restoreLockedJikantai(Mcm1002u2Form form,
                                      Mcm1002u1Form screen1Form,
                                      Mcm1002uDeliveryDto delivery) {
        if (form == null || delivery == null || !delivery.isCopyFlg()) {
            return;
        }
        List<MitsumoriRow> expectedRows = buildScreen2Form(screen1Form, delivery).getMitsumoriRows();
        List<MitsumoriRow> postedRows = form.getMitsumoriRows();
        for (int i = 0; i < postedRows.size(); i++) {
            MitsumoriRow posted = postedRows.get(i);
            MitsumoriRow expected = expectedRows.stream()
                .filter(r -> sameId(r.getTorihikisakiId(), posted.getTorihikisakiId()))
                .findFirst()
                .orElse(i < expectedRows.size() ? expectedRows.get(i) : null);
            if (expected == null) {
                continue;
            }
            posted.setH8(expected.isH8());
            posted.setH24(expected.isH24());
        }
    }

    // ================================================================
    // Screen1で選択された機器構成パターン・機器構成明細をScreen2Formに反映する
    //   機器構成はチェック済みの機器明細を1件以上含む構成のみ表示する。
    // ================================================================
    private void appendKikikoseiFromScreen1(Mcm1002u2Form form, Mcm1002u1Form screen1Form) {
        for (Mcm1002u1Form.KikikoseiRow kose : screen1Form.getKikikoseiRows()) {
            if (!kose.isChecked()) {
                continue;
            }
            // 【注意】BigDecimal#equals()はスケール(小数桁数)まで一致しないとfalseになるため、
            //   数値としての比較にはcompareTo()を使用する。
            boolean hasCheckedMeisai = screen1Form.getKikimeisaiRows().stream()
                .anyMatch(m -> m.isChecked()
                    && m.getKikikoseiId() != null && kose.getKikikoseiId() != null
                    && m.getKikikoseiId().compareTo(kose.getKikikoseiId()) == 0);
            if (!hasCheckedMeisai) {
                continue;
            }
            Mcm1002u2Form.KikikoseiRow kr = new Mcm1002u2Form.KikikoseiRow();
            // 【注意】機器構成明細側とはBigDecimalのスケールが異なる場合があり、
            //   画面側でのJS文字列比較(絞り込み表示)がずれるため、stripTrailingZerosで正規化する。
            kr.setKikikoseiId(kose.getKikikoseiId() != null ? kose.getKikikoseiId().stripTrailingZeros() : null);
            kr.setKikikoseiNk(kose.getKikikoseiNk());
            kr.setTehaiseiban(kose.getTehaiseiban());
            form.getKikikoseiRows().add(kr);
        }
        for (Mcm1002u1Form.KikimeisaiRow me : screen1Form.getKikimeisaiRows()) {
            if (!me.isChecked()) {
                continue;
            }
            Mcm1002u2Form.KikimeisaiRow mr = new Mcm1002u2Form.KikimeisaiRow();
            mr.setKikikoseiId(me.getKikikoseiId() != null ? me.getKikikoseiId().stripTrailingZeros() : null);
            mr.setSeizomakerNk(me.getSeizomakerNk());
            mr.setKikihinmeiNk(me.getKikihinmeiNk());
            mr.setKikikatashiki(me.getKikikatashiki());
            mr.setSuryoNm(me.getSuryoNm());
            form.getKikimeisaiRows().add(mr);
        }
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
                errors.add(messageSource.getMessage("MSG_0145E", null, Locale.JAPANESE));
            }

            // 【変換元】MCM0013U - dateError() 準拠：YYYY/MM/DD形式チェック
            String dateErr = dateError("回答希望日", row.getKaitokizitsuDt());
            if (dateErr != null) {
                errors.add("行" + (i + 1) + ": " + dateErr);
            }
        }
        return errors;
    }

    /**
     * 日付文字列(YYYY/MM/DD)の書式チェック
     * 【変換元】Mcm0013uService.dateError() 準拠
     */
    private static String dateError(String label, String value) {
        if (value == null || value.trim().isEmpty()) {
            return null;
        }
        if (!value.matches("\\d{4}/\\d{2}/\\d{2}")) {
            return label + "はYYYY/MM/DDの書式で入力してください。";
        }
        try {
            java.time.LocalDate.parse(value, DateTimeFormatter.ofPattern("yyyy/MM/dd"));
        } catch (java.time.DateTimeException e) {
            return label + "はYYYY/MM/DDの書式で入力してください。";
        }
        return null;
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
    public Long saveAll(Mcm1002u2Form form,
                        Mcm1002u1Form screen1Form,
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

        // 【不一致修正】VB版はTM_IRAI_ID/TM_KIKIKOSEI_ID/TM_KIKIMEISAI_ID/
        //   TM_KOTAIMEISAI_ID/TM_KEIYAKUJIKAN_IDをDBの自動採番(IDENTITY)に
        //   任せず、アプリ側でMAX(該当列)+1を計算して明示的にセットしていた
        //   (Mcm1002u2Screen.vb updateMethod()の新規ID採番部)。
        //   このトランザクション内で複数件新規作成される可能性があるため、
        //   採番済みIDが重複しないようAtomicLongで1件ずつインクリメントする。
        BigDecimal maxKeiyakujikanId = keiyakujikanRepo.findMaxKeiyakujikanId();
        IdAllocator idAllocator = new IdAllocator(
            mitsumoriRepo.findMaxIraiId(),
            kikikoseiRepo.findMaxKikikoseiId(),
            kikimeisaiRepo.findMaxKikimeisaiId(),
            kotaimeisaiRepo.findMaxKotaimeisaiId(),
            maxKeiyakujikanId != null ? maxKeiyakujikanId.longValue() : 0L);

        // 登録後、呼び出し元(Controller)がdelivery.tmKeiyakujikanIdを更新して
        // 次回表示を修正モードへ切り替えるための代表ID（複数見積依頼行がある場合は先頭のもの）
        Long resultKeiyakujikanId = delivery != null ? delivery.getTmKeiyakujikanId() : null;
        if (resultKeiyakujikanId != null && resultKeiyakujikanId <= 0) {
            resultKeiyakujikanId = null;
        }
        // 【#471不一致修正】複製時のdelivery.tmKeiyakujikanIdは複製元のIDのため、
        //   これを返すと登録後の再表示（修正モード）で複製元の見積が表示されてしまう。
        //   複製時は今回新規作成した契約時間IDを返す。
        if (mitsumoriFlg == Mcm1002uConstants.FUKUSEI) {
            resultKeiyakujikanId = null;
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
                mitsumori.setTmIraiId(idAllocator.nextIraiId());
                mitsumori.setTmIraiNo(generateIraiNo());
                mitsumori.setCreatedDt(now);
                mitsumori.setCreatedBy(loginUser);
                mitsumori.setMitsumoriDt(now);
                mitsumori.setIraitantosya(loginUser);

                /*
                 * 【不一致修正】元VB: Mcm1002u2Screen.vb updateMethod() 見積テーブル追加設定部
                 *   mitsumoriNewRow.NONYUSAKI_ID/NONYUSAKI_CD/NONYUSAKI_NK/NONYUSAKIJUSYO1_NK/
                 *   NONYUSAKIJUSYO2_NK/PLANT_ID/SUPPORT_ID/PLANT_NK/NONYUBUSYO_NK/
                 *   NONYUTANTOSYA_NK/NONYUTEL_NO/NONYUFAX_NO/TORIHIKISAKI_ID/TORIHIKISAKI_CD/
                 *   TORIHIKISAKI_NK/TORITEL_NO/TORIFAX_NO/TORIJIGYOSYO_NK/TORISYUTANTOSYA_NK/
                 *   TORIASSISTANT_NKを新規行に設定していたが、Web版では従来これらが
                 *   一切setされておらずNULLで登録されていた（PLANT_ID等がNULLになる不具合の
                 *   直接原因。行削除ボタンでも同様の欠落があったため今回まとめて追加）。
                 */
                mitsumori.setNonyusakiId(row.getNonyusakiId());
                mitsumori.setNonyusakiCd(row.getNonyusakiCd());
                mitsumori.setNonyusakiNk(row.getNonyusakiNk());
                mitsumori.setNonyusakijusyo1Nk(row.getNonyusakijusyo1Nk());
                mitsumori.setNonyusakijusyo2Nk(row.getNonyusakijusyo2Nk());
                mitsumori.setPlantId(row.getPlantId());
                mitsumori.setSupportId(row.getSupportId());
                mitsumori.setPlantNk(row.getPlantNk());
                mitsumori.setNonyubusyoNk(row.getNonyubusyoNk());
                mitsumori.setNonyutantosyaNk(row.getNonyutantosyaNk());
                mitsumori.setNonyutelNo(row.getNonyutelNo());
                mitsumori.setNonyufaxNo(row.getNonyufaxNo());
                mitsumori.setTorihikisakiId(row.getTorihikisakiId());
                mitsumori.setTorihikisakiCd(row.getTorihikisakiCd());
                mitsumori.setTorihikisakiNk(row.getTorihikisakiNk());
                mitsumori.setToritelNo(row.getToritelNo());
                mitsumori.setTorifaxNo(row.getTorifaxNo());
                mitsumori.setTorijigyosyoNk(row.getTorijigyosyoNk());
                mitsumori.setTorisyutantosyaNk(row.getTorisyutantosyaNk());
                mitsumori.setToriassistantNk(row.getToriassistantNk());
            }

            // 【#373不一致修正】新規作成した見積依頼の採番済みTM_IRAI_IDを画面Form側にも
            //   反映する。saveAll()実行前はrow.tmIraiIdがnullのままのため、依頼書発行時の
            //   Excel帳票データ組み立て（buildIraisyoDelivery）が対象を特定できなかった。
            row.setTmIraiId(mitsumori.getTmIraiId());
            row.setTmIraiNo(mitsumori.getTmIraiNo());

            // Form → Entity マッピング
            // 【不一致修正】H8/H24はMCM_TM_MITSUMORIの実カラムではないため、
            //   Entityへの直接setterは行わない。永続化はsyncKeiyakujikan()が
            //   MCM_TM_KEIYAKUJIKANへの行追加/削除として行う。
            mitsumori.setTenkenumu(row.isTenkenumu() ? "1" : "0");
            mitsumori.setTenkenkanoyobi(row.getTenkenkanoyobi());
            mitsumori.setYakantaioumu(row.isYakantaioumu() ? "1" : "0");
            mitsumori.setHosyuhoho(row.getHosyuhoho());
            mitsumori.setPackFlg(row.isPackFlg() ? BigDecimal.ONE : BigDecimal.ZERO);
            mitsumori.setTenpoFlg(row.isTenpoFlg() ? BigDecimal.ONE : BigDecimal.ZERO);
            // 【#373不一致修正】回答希望日(KAITOKIZITSU_DT)がこれまでEntityへ反映されておらず、
            //   常にNULLで保存されていた（依頼書発行のExcel帳票ヘッダに反映されない不具合の
            //   一因でもあるため合わせて修正）。画面の文字列(yyyy/MM/dd)をLocalDateTimeへ変換する。
            mitsumori.setKaitokizitsuDt(parseDateTime(row.getKaitokizitsuDt()));
            mitsumori.setLastupdateDt(now);
            mitsumori.setLastupdateBy(loginUser);

            // -----------------------------------------------
            // 契約時間帯の同期（8H / 24H）
            // 【変換元】updateMethod() 内の契約時間DataTable削除/追加
            // -----------------------------------------------
            syncKeiyakujikan(mitsumori, row, now, loginUser, printFlg, jotai, idAllocator);

            // -----------------------------------------------
            // 機器構成・機器明細・個体明細の保存
            // 【変換元】Mcm1002u1Screen.vb - NextButton_Click()（機器構成/機器明細/取引先Array格納）
            //         Mcm1002u2Screen.vb - updateMethod()（新規ID採番→親子ID付替え→一括保存）
            //   Web版差異: VB版は仮ID採番と付替えをアプリ側で行うが、IDそのものは
            //   VB版同様にMAX+1で明示的に採番する（DBはIDENTITY自動採番列ではない）。
            //   修正時は既存の子データを一旦クリアしてScreen1の選択内容で作り直す
            //   （差分更新はサポートしない）。
            // -----------------------------------------------
            if (screen1Form != null) {
                syncKikikosei(mitsumori, screen1Form, row, now, loginUser, idAllocator);
            }

            // 保存（Cascade で子テーブルも一括保存）
            mitsumoriRepo.save(mitsumori);

            // 【不一致修正】元VB: afterUpdate() でmitsumoriFlgを「修正」に変更し、
            //   以後は登録済みの契約時間IDを用いて画面を再表示する。
            //   Web版では呼び出し元(Controller)がdelivery.tmKeiyakujikanIdを
            //   更新して次回表示を修正モードに切り替えるため、代表となる
            //   契約時間IDをここで収集して返す。
            if (resultKeiyakujikanId == null && !mitsumori.getKeiyakujikanList().isEmpty()) {
                resultKeiyakujikanId = mitsumori.getKeiyakujikanList().get(0).getTmKeiyakujikanId().longValue();
            }
        }
        return resultKeiyakujikanId;
    }

    /**
     * 見積依頼(取引先)行に紐づく機器構成・機器明細・個体明細を、Screen1で選択された
     * 内容に基づいて再構築する。
     * 【変換元】Mcm1002u1Screen.vb - NextButton_Click() 機器構成/機器明細Array格納部
     */
    private void syncKikikosei(McmTmMitsumoriEntity mitsumori,
                                Mcm1002u1Form screen1Form,
                                MitsumoriRow mitsumoriRow,
                                LocalDateTime now,
                                String loginUser,
                                IdAllocator idAllocator) {

        // この見積依頼(取引先)に属する、チェック済みの個体明細を抽出する
        List<Mcm1002u1Form.KotaimeisaiRow> targetKotai = screen1Form.getKotaimeisaiRows().stream()
            .filter(Mcm1002u1Form.KotaimeisaiRow::isChecked)
            .filter(k -> sameId(k.getTorihikisakiId(), mitsumoriRow.getTorihikisakiId()))
            .collect(Collectors.toList());

        if (targetKotai.isEmpty()) {
            return;
        }

        // 既存の機器構成一式をクリアして作り直す（cascade + orphanRemovalでDBからも削除される）
        mitsumori.getKikikoseiList().clear();

        // 機器構成ID → Entity のキャッシュ（同一機器構成の重複作成を避ける）
        Map<String, McmTmKikikoseiEntity> koseiCache = new LinkedHashMap<>();
        // 機器明細ID → Entity のキャッシュ（同一機器明細の重複作成を避ける）
        Map<String, McmTmKikimeisaiEntity> meisaiCache = new LinkedHashMap<>();

        for (Mcm1002u1Form.KotaimeisaiRow kotai : targetKotai) {

            // 対応する機器明細（マスタ選択情報）を取得
            Mcm1002u1Form.KikimeisaiRow meisaiSrc = screen1Form.getKikimeisaiRows().stream()
                .filter(m -> sameId(m.getKikimeisaiId(), kotai.getKikimeisaiId()))
                .findFirst().orElse(null);
            if (meisaiSrc == null) {
                continue;
            }

            // 対応する機器構成（マスタ選択情報）を取得
            Mcm1002u1Form.KikikoseiRow koseiSrc = screen1Form.getKikikoseiRows().stream()
                .filter(k -> sameId(k.getKikikoseiId(), meisaiSrc.getKikikoseiId()))
                .findFirst().orElse(null);
            if (koseiSrc == null) {
                continue;
            }

            // 機器構成Entity（同じ機器構成IDなら再利用）
            String koseiKey = String.valueOf(koseiSrc.getKikikoseiId());
            McmTmKikikoseiEntity koseiEntity = koseiCache.get(koseiKey);
            if (koseiEntity == null) {
                koseiEntity = new McmTmKikikoseiEntity();
                koseiEntity.setTmKikikoseiId(idAllocator.nextKikikoseiId());
                koseiEntity.setMitsumori(mitsumori);
                koseiEntity.setKikikoseiId(koseiSrc.getKikikoseiId());
                koseiEntity.setKikikoseiNk(koseiSrc.getKikikoseiNk());
                koseiEntity.setSetNm(koseiSrc.getSetNm());
                koseiEntity.setTehaiseiban(koseiSrc.getTehaiseiban());
                koseiEntity.setBiko(koseiSrc.getBiko());
                koseiEntity.setHyojijun(koseiSrc.getHyojijun());
                koseiEntity.setCreatedDt(now);
                koseiEntity.setCreatedBy(loginUser);
                mitsumori.getKikikoseiList().add(koseiEntity);
                koseiCache.put(koseiKey, koseiEntity);
            }

            // 機器明細Entity（同じ機器明細IDなら再利用）
            String meisaiKey = koseiKey + "-" + meisaiSrc.getKikimeisaiId();
            McmTmKikimeisaiEntity meisaiEntity = meisaiCache.get(meisaiKey);
            if (meisaiEntity == null) {
                meisaiEntity = new McmTmKikimeisaiEntity();
                meisaiEntity.setTmKikimeisaiId(idAllocator.nextKikimeisaiId());
                meisaiEntity.setKikikosei(koseiEntity);
                meisaiEntity.setSeizomakerId(meisaiSrc.getSeizomakerId());
                meisaiEntity.setSeizomakerNk(meisaiSrc.getSeizomakerNk());
                meisaiEntity.setKikimeisaiId(meisaiSrc.getKikimeisaiId());
                meisaiEntity.setKikihinmeiNk(meisaiSrc.getKikihinmeiNk());
                meisaiEntity.setKikikatashiki(meisaiSrc.getKikikatashiki());
                meisaiEntity.setSuryoNm(meisaiSrc.getSuryoNm());
                meisaiEntity.setAtsukaikikiId(meisaiSrc.getAtsukaikikiId());
                meisaiEntity.setBiko(meisaiSrc.getBiko());
                meisaiEntity.setHyojijun(meisaiSrc.getHyojijun());
                meisaiEntity.setCreatedDt(now);
                meisaiEntity.setCreatedBy(loginUser);
                koseiEntity.getKikimeisaiList().add(meisaiEntity);
                meisaiCache.put(meisaiKey, meisaiEntity);
            }

            // 個体明細Entity（個体明細は1個体1行）
            McmTmKotaimeisaiEntity kotaiEntity = new McmTmKotaimeisaiEntity();
            kotaiEntity.setTmKotaimeisaiId(idAllocator.nextKotaimeisaiId());
            kotaiEntity.setKikimeisai(meisaiEntity);
            kotaiEntity.setKotaikanriId(kotai.getKotaikanriId());
            kotaiEntity.setKotaiNk(kotai.getKotaiNk());
            kotaiEntity.setSerialNo(kotai.getSerialNo());
            kotaiEntity.setSetchibasyo(kotai.getSetchibasyo());
            kotaiEntity.setItijinonyuDt(parseDateTime(kotai.getItijinonyuDt()));
            kotaiEntity.setTekkyobiDt(parseDateTime(kotai.getTekkyoDt()));
            kotaiEntity.setKeiyakukigenDt(parseDateTime(kotai.getKeiyakukigenDt()));
            kotaiEntity.setEnchokeiyakukigenDt(parseDateTime(kotai.getEnchokeiyakukigenDt()));
            kotaiEntity.setBrandId(kotai.getBrandId());
            kotaiEntity.setBrandNk(kotai.getBrandNk());
            kotaiEntity.setBrandkoseiId(kotai.getBrandkoseiId());
            kotaiEntity.setBrandsyosaiNk(kotai.getBrandsyosaiNk());
            kotaiEntity.setCreatedDt(now);
            kotaiEntity.setCreatedBy(loginUser);
            meisaiEntity.getKotaimeisaiList().add(kotaiEntity);
        }
    }

    /** BigDecimal同士をスケール差異を無視して比較する（null安全） */
    private boolean sameId(BigDecimal a, BigDecimal b) {
        return a != null && b != null && a.compareTo(b) == 0;
    }

    /** "yyyy/MM/dd"等の日付文字列をLocalDateTimeへ変換する（変換不可・空文字はnull） */
    private LocalDateTime parseDateTime(String dateStr) {
        if (dateStr == null || dateStr.isBlank()) {
            return null;
        }
        try {
            return java.time.LocalDate.parse(dateStr.replace("/", "-")).atStartOfDay();
        } catch (Exception e) {
            return null;
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
                                   String jotai,
                                   IdAllocator idAllocator) {

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
                    mitsumori, Mcm1002uConstants.JIKAN_8, now, loginUser, idAllocator);
                existing.add(newK);
            }
        } else {
            existing.removeIf(k -> Mcm1002uConstants.JIKAN_8.equals(k.getKeiyakujikantai()));
        }

        // 24H 処理
        if (row.isH24()) {
            if (!byJikantai.containsKey(Mcm1002uConstants.JIKAN_24)) {
                McmTmKeiyakujikanEntity newK = createKeiyakujikan(
                    mitsumori, Mcm1002uConstants.JIKAN_24, now, loginUser, idAllocator);
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
            LocalDateTime now, String loginUser, IdAllocator idAllocator) {
        McmTmKeiyakujikanEntity k = new McmTmKeiyakujikanEntity();
        k.setTmKeiyakujikanId(idAllocator.nextKeiyakujikanId());
        // 【不一致修正】setMitsumori()はTODOスタブで実装が空のため、
        //   実際にINSERT対象となるFK列(tmIraiId)へ直接値をセットする。
        //   これが無いとMCM_TM_KEIYAKUJIKAN.TM_IRAI_IDがNULLで保存され、
        //   MCM_TM_MITSUMORIとの結合検索(findByKeiyakujikanId)が0件になり、
        //   登録後に「見積情報」グリッドが空になる不具合が起きていた。
        k.setTmIraiId(mitsumori.getTmIraiId() != null
            ? BigDecimal.valueOf(mitsumori.getTmIraiId()) : null);
        k.setMitsumori(mitsumori);
        k.setKeiyakujikantai(jikantai);
        k.setJotai(Mcm1002uConstants.JOTAI_SAKUSEICHU);
        k.setCreatedDt(now);
        k.setCreatedBy(loginUser);
        return k;
    }

    /**
     * 見積依頼登録処理内で使用するID採番ヘルパー。
     * 【変換元】Mcm1002u2Screen.vb updateMethod() の maxTma/maxTmc/maxTmd/maxKotaiId/
     *   maxKeiyakujikanId のMAX+1手動採番ロジック。VB版はDB自動採番(IDENTITY)を使わず、
     *   トランザクション開始時に各テーブルのMAX値を取得し、以後1件ずつインクリメントして
     *   新規行に採番する。Web版でも同じ方式を採用する（1トランザクション内で複数件
     *   新規作成される場合でもID重複が起きないようにするため）。
     */
    private static final class IdAllocator {
        private long iraiId;
        private long kikikoseiId;
        private long kikimeisaiId;
        private long kotaimeisaiId;
        private long keiyakujikanId;

        IdAllocator(Long maxIraiId, Long maxKikikoseiId, Long maxKikimeisaiId,
                    Long maxKotaimeisaiId, Long maxKeiyakujikanId) {
            this.iraiId = maxIraiId != null ? maxIraiId : 0L;
            this.kikikoseiId = maxKikikoseiId != null ? maxKikikoseiId : 0L;
            this.kikimeisaiId = maxKikimeisaiId != null ? maxKikimeisaiId : 0L;
            this.kotaimeisaiId = maxKotaimeisaiId != null ? maxKotaimeisaiId : 0L;
            this.keiyakujikanId = maxKeiyakujikanId != null ? maxKeiyakujikanId : 0L;
        }

        Long nextIraiId() { return ++iraiId; }
        Long nextKikikoseiId() { return ++kikikoseiId; }
        Long nextKikimeisaiId() { return ++kikimeisaiId; }
        Long nextKotaimeisaiId() { return ++kotaimeisaiId; }
        BigDecimal nextKeiyakujikanId() { return BigDecimal.valueOf(++keiyakujikanId); }
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
    // 【変換元】McmHoshuhohoDataTable_Item.xml（20_COMMON/DATATABLE）
    //   VB版はDBマスタではなく、アプリ内固定のXML設定（CPComboBoxDataTable派生）から
    //   選択肢を読み込んでいた。Web版でも同じキー・表示値をそのまま移植する。
    public Map<String, String> getHosyuhohoOptions() {
        Map<String, String> map = new LinkedHashMap<>();
        map.put("", "");                  // AddBlankRow=True 相当（先頭に空行）
        map.put("F", "ｵﾝｻｲﾄ");
        map.put("S", "ｾﾝﾄﾞﾊﾞｯｸ");
        map.put("C", "ｺﾝﾃｯｸ製品");
        map.put("H", "持ち帰り");
        map.put("I", "ｽﾎﾟｯﾄ");
        map.put("T", "TEL対応");
        return map;
    }

    public Map<String, String> getTenkenkanoyobiOptions() {
        Map<String, String> map = new LinkedHashMap<>();
        map.put("", "");                  // AddBlankRow=True 相当（先頭に空行）
        map.put("1", "月～金");
        map.put("2", "月～土");
        map.put("3", "月～日");
        map.put("4", "土・日");
        return map;
    }

    // ================================================================
    // Excel帳票出力用データ組み立て
    // 【変換元】Mcm1002u2Screen.vb - IraisyoHakkoButton_Click() 帳票出力処理
    //   元コード: 見積DataTableを取引先(TM_IRAI_ID)単位でループし、契約時間
    //            (8H/24H)ごとにヘッダー情報＋一覧情報を組み立てて
    //            excelOutput()を呼び出す（取引先1件＝Excel1ファイル）。
    // ================================================================
    @Transactional(readOnly = true)
    public List<Mcm1001pDeliveryDto> buildIraisyoDeliveryList(List<Long> tmIraiIds) {
        List<Mcm1001pDeliveryDto> result = new ArrayList<>();
        if (tmIraiIds == null) {
            return result;
        }
        for (Long tmIraiId : new java.util.LinkedHashSet<>(tmIraiIds)) {
            if (tmIraiId == null || tmIraiId <= 0) {
                continue;
            }
            McmTmMitsumoriEntity mitsumori = mitsumoriRepo.findById(tmIraiId).orElse(null);
            if (mitsumori == null) {
                throw new IllegalArgumentException("見積依頼の出力対象が見つかりません。");
            }

            Mcm1001pDeliveryDto delivery = new Mcm1001pDeliveryDto();
            delivery.setTmIraiNo(mitsumori.getTmIraiNo());
            delivery.setTorihikisakiCd(mitsumori.getTorihikisakiCd());

            List<Mcm1001pHeaderDto> headerList = new ArrayList<>();
            List<Mcm1001pViewDto> allMeisaiList = new ArrayList<>();
            int cursor = 0;

            // 【変換元】keiyakujikanTable = MCM_TM_KEIYAKUJIKAN.Select(TM_IRAI_ID = ...)
            List<McmTmKeiyakujikanEntity> keiyakujikanList =
                keiyakujikanRepo.findByMitsumori_TmIraiId(tmIraiId);

            keiyakujikanList = new ArrayList<>(keiyakujikanList);
            keiyakujikanList.sort(java.util.Comparator.comparing(
                    k -> nz(k.getKeiyakujikantai())));
            if (keiyakujikanList.isEmpty()) {
                throw new IllegalArgumentException("見積依頼の契約時間が未登録です。");
            }
            for (McmTmKeiyakujikanEntity keiyakujikan : keiyakujikanList) {

                // 【変換元】MCM_TM_KIKIKOSEI_EXCELTableAdapter.GetData(tmKeiyakujikanId)
                List<Map<String, Object>> koseiRows = mcm1002uRepo
                    .findKikikoseiExcelByKeiyakujikanId(keiyakujikan.getTmKeiyakujikanId());

                int startIndex = cursor;
                for (Map<String, Object> koseiRow : koseiRows) {
                    Mcm1001pViewDto view = new Mcm1001pViewDto();
                    view.setKikikoseiNk(toStr(koseiRow.get("KIKIKOSEI_NK")));
                    view.setSeizomakerNk(toStr(koseiRow.get("SEIZOMAKER_NK")));
                    view.setKikihinmeiNk(toStr(koseiRow.get("KIKIHINMEI_NK")));
                    view.setKikikatashiki(toStr(koseiRow.get("KIKIKATASHIKI")));
                    Object suryo = koseiRow.get("SURYO_NM");
                    view.setSuryoNm(suryo != null ? new BigDecimal(suryo.toString()) : null);
                    view.setTehaiseiban(toStr(koseiRow.get("TEHAISEIBAN")));
                    allMeisaiList.add(view);
                    cursor++;
                }
                int endIndex = cursor - 1; // 0件はend < start。次の時間帯の明細を含めない。

                Mcm1001pHeaderDto header = new Mcm1001pHeaderDto();
                header.setIraiNo(mitsumori.getTmIraiNo());
                header.setTorihikisakiNk(mitsumori.getTorihikisakiNk());
                header.setNonyusakiNk(mitsumori.getNonyusakiNk());
                header.setNonyusakijusyo1Nk(mitsumori.getNonyusakijusyo1Nk());
                header.setNonyusakijusyo2Nk(mitsumori.getNonyusakijusyo2Nk());
                header.setKaitokizitsuDt(formatDate(mitsumori.getKaitokizitsuDt()));
                header.setMitsumoriDt(formatDate(mitsumori.getMitsumoriDt()));
                header.setHosyuhoho(getHosyuhohoOptions().getOrDefault(nz(mitsumori.getHosyuhoho()), ""));
                header.setTenkenumu("1".equals(mitsumori.getTenkenumu()) ? "点検あり" : "0".equals(mitsumori.getTenkenumu()) ? "点検なし" : "");
                header.setKeiyakujikantai(nz(keiyakujikan.getKeiyakujikantai()) + "H");
                header.setTenkenkanobi(getTenkenkanoyobiOptions()
                    .getOrDefault(nz(mitsumori.getTenkenkanoyobi()), ""));
                header.setYakanTenken("1".equals(mitsumori.getYakantaioumu()) ? "夜間対応あり" : "0".equals(mitsumori.getYakantaioumu()) ? "夜間対応なし" : "");
                header.setStartIndex(startIndex);
                header.setEndIndex(endIndex);
                headerList.add(header);
            }

            delivery.setHeaderList(headerList);
            delivery.setMeisaiList(allMeisaiList);
            result.add(delivery);
        }
        result.sort(java.util.Comparator.comparing(d -> nz(d.getTmIraiNo())));
        return result;
    }

    private static String toStr(Object v) {
        return v == null ? null : v.toString();
    }

    private static String nz(String v) {
        return v == null ? "" : v;
    }

    private static String formatDate(LocalDateTime dt) {
        if (dt == null) {
            return "";
        }
        return dt.format(DateTimeFormatter.ofPattern("yyyy/MM/dd"));
    }
}