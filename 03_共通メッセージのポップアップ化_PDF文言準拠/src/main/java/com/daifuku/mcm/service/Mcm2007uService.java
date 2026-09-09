package com.daifuku.mcm.service;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.daifuku.mcm.form.Mcm2007uForm;
import com.daifuku.mcm.form.Mcm2007uForm.KihonBrandRowForm;
import com.daifuku.mcm.form.Mcm2007uForm.MitsumoriRowForm;
import com.daifuku.mcm.repository.Mcm2007uRepository;

/**
 * 【変換元】Mcm2007uScreen1.vb / Mcm2007uScreen2.vb
 * MCM2007U ユーザ契約内容変更サービス（DB書き込みなし）
 *
 * Step1: 店舗見積+ブランド選択一覧取得・バリデーション
 * Step2: 機器構成・明細・個体選択一覧取得
 */
@Service
public class Mcm2007uService {

    @Autowired
    private Mcm2007uRepository repo;

    private static final DateTimeFormatter DT_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");

    // ===================================================================
    // Step1: 店舗見積一覧ロード
    // 【変換元】Mcm2007uScreen1.vb — Form_Load
    // ===================================================================

    public void loadStep1(Mcm2007uForm form) {
        List<MitsumoriRowForm> mitsumoriList = repo.findMitsumoriList(form.getPlantId());

        if (!mitsumoriList.isEmpty()) {
            List<BigDecimal> allMitsumoriIds = mitsumoriList.stream()
                .map(MitsumoriRowForm::getUmMitsumoriId)
                .collect(Collectors.toList());

            List<KihonBrandRowForm> allBrands = repo.findKihonBrandList(allMitsumoriIds);

            // 見積IDでブランドをグルーピング
            for (MitsumoriRowForm m : mitsumoriList) {
                List<KihonBrandRowForm> brands = allBrands.stream()
                    .filter(b -> m.getUmMitsumoriId().equals(b.getUmMitsumoriId()))
                    .collect(Collectors.toList());
                m.setBrandRows(brands);
            }
        }
        form.setMitsumoriRows(mitsumoriList);
    }

    /** Mcm2007uScreen.vb getUmMitsumoriId + 初期チェック。2004Uからの新規作成時のみ。 */
    public void selectInitialEstimate(Mcm2007uForm form, BigDecimal estimateId) {
        BigDecimal selected = repo.findInitialMitsumoriId(estimateId);
        if (selected == null) return;
        for (var row : form.getMitsumoriRows()) {
            boolean matches = row.getUmMitsumoriId() != null && row.getUmMitsumoriId().compareTo(selected) == 0;
            for (var brand : row.getBrandRows()) brand.setCheckFlg(matches);
            if (matches) {
                form.setKaisiDt(row.getKaisiDt());
                form.setSyuryoDt(row.getSyuryoDt() == null || row.getSyuryoDt().isBlank()
                    ? LocalDate.parse(row.getKaisiDt(), DT_FMT).plusYears(1).minusDays(1).format(DT_FMT)
                    : row.getSyuryoDt());
            }
        }
    }

    // ===================================================================
    // Step1→Step2: バリデーション
    // 【変換元】NextButton_Click
    // ===================================================================

    public List<String> validateStep1(Mcm2007uForm form) {
        List<String> errors = new ArrayList<>();

        // 終了日必須
        if (form.getSyuryoDt() == null || form.getSyuryoDt().isBlank()) {
            errors.add("終了日を入力してください。");
            return errors;
        }

        // 日付形式チェック
        LocalDate kaisi = null, syuryo = null;
        try {
            kaisi = LocalDate.parse(form.getKaisiDt(), DT_FMT);
        } catch (Exception e) {
            errors.add("開始日付の書式を指定して下さい。(YYYY/MM/DD)");
        }
        try {
            syuryo = LocalDate.parse(form.getSyuryoDt(), DT_FMT);
        } catch (Exception e) {
            errors.add("終了日付の書式を指定して下さい。(YYYY/MM/DD)");
        }
        if (!errors.isEmpty()) return errors;

        // 開始日≦終了日
        if (kaisi != null && syuryo != null && syuryo.isBefore(kaisi)) {
            errors.add("開始日は終了日よりも前の日付で入力して下さい。");
        }

        // 期間1年以内
        if (kaisi != null && syuryo != null) {
            if (syuryo.isAfter(kaisi.plusYears(1))) {
                errors.add("契約期間は1年以内で設定してください。");
            }
        }

        // チェック件数0件チェック
        long checkedCount = form.getMitsumoriRows().stream()
            .flatMap(m -> m.getBrandRows().stream())
            .filter(KihonBrandRowForm::isCheckFlg)
            .count();
        if (checkedCount == 0) {
            errors.add("ブランドを1件以上選択してください。");
        }

        return errors;
    }

    // ===================================================================
    // Step2: 機器構成・明細・個体一覧ロード
    // 【変換元】Mcm2007uScreen2.vb — Form_Load
    // ===================================================================

    public void loadStep2(Mcm2007uForm form) {
        // Step1で選択された見積IDリストを取得
        List<BigDecimal> selectedMitsumoriIds = form.getMitsumoriRows().stream()
            .flatMap(m -> m.getBrandRows().stream())
            .filter(KihonBrandRowForm::isCheckFlg)
            .map(KihonBrandRowForm::getUmMitsumoriId)
            .distinct()
            .collect(Collectors.toList());

        form.setKoseiRows(repo.findKoseiList(selectedMitsumoriIds, form.getPlantId()));
        form.setMeisaiRows(repo.findMeisaiList(selectedMitsumoriIds, form.getPlantId()));
        form.setKotaiRows(repo.findKotaiList(selectedMitsumoriIds, form.getPlantId()));
    }

    // ===================================================================
    // Step2→Apply: バリデーション（重複個体チェック）
    // 【変換元】ApplyButton_Click
    // ===================================================================

    public List<String> validateStep2(Mcm2007uForm form) {
        List<String> errors = new ArrayList<>();

        // 重複個体チェック（既存契約との重複）
        if (form.getUkKeiyakuId() != null) {
            List<BigDecimal> checkedKotaiIds = form.getKotaiRows().stream()
                .filter(Mcm2007uForm.KotaiRowForm::isCheckFlg)
                .map(Mcm2007uForm.KotaiRowForm::getKotaikanriId)
                .collect(Collectors.toList());

            if (!checkedKotaiIds.isEmpty()) {
                List<String> duplicates = repo.findDuplicateKotai(checkedKotaiIds, form.getUkKeiyakuId());
                for (String name : duplicates) {
                    errors.add("個体「" + name + "」は既にこの契約に登録されています。");
                }
            }
        }

        return errors;
    }
}
