package com.daifuku.mcm.service;

import java.math.BigDecimal;
import com.daifuku.mcm.form.Mcm2006uForm;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;
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

    private static final DateTimeFormatter DT_FMT = DateTimeFormatter.ofPattern("uuuu/MM/dd").withResolverStyle(java.time.format.ResolverStyle.STRICT);

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
                    .filter(b -> same(m.getUmMitsumoriId(), b.getUmMitsumoriId()))
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
        if (kaisi != null && syuryo != null && !syuryo.isAfter(kaisi)) {
            errors.add("開始日は終了日よりも前の日付で入力して下さい。");
        }

        // 期間1年以内
        if (kaisi != null && syuryo != null) {
            if (syuryo.isAfter(kaisi.plusYears(1).minusDays(1))) {
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

    /** VBの既存ブランドのチェック状態を復元する。 */
    public void selectExisting(Mcm2007uForm form, Mcm2006uForm.KikanTabForm tab) {
        form.setUkKikanId(tab.getUkKikanId());
        for (var m : form.getMitsumoriRows()) for (var b : m.getBrandRows()) b.setCheckFlg(
            tab.getBrandRows().stream().anyMatch(old -> same(old.getUmMitsumoriId(), b.getUmMitsumoriId())
                && same(old.getBrandkoseiId(), b.getBrandkoseiId())));
    }
    /** 変更開始時の選択個体を維持する。見積の初期値と既存契約を合成する。 */
    public void selectExistingEquipment(Mcm2007uForm form, Mcm2006uForm.KikanTabForm tab) {
        if (tab == null) return;
        for (var k : form.getKoseiRows()) tab.getKoseiRows().stream().filter(old -> same(old.getBrandkoseiId(),k.getBrandkoseiId()) && same(old.getKikikoseiId(),k.getKikikoseiId())).findFirst().ifPresent(old -> {
            org.springframework.beans.BeanUtils.copyProperties(old,k); k.setCheckFlg(true);
        });
        for (var m : form.getMeisaiRows()) tab.getMeisaiRows().stream().filter(old -> same(old.getBrandkoseiId(),m.getBrandkoseiId()) && same(old.getKikimeisaiId(),m.getKikimeisaiId())).findFirst().ifPresent(old -> {
            org.springframework.beans.BeanUtils.copyProperties(old,m); m.setCheckFlg(true);
        });
        for (var t : form.getKotaiRows()) tab.getKotaiRows().stream().filter(old -> same(old.getBrandkoseiId(),t.getBrandkoseiId()) && same(old.getKikimeisaiId(),t.getKikimeisaiId()) && same(old.getKotaikanriId(),t.getKotaikanriId())).findFirst().ifPresent(old -> {
            org.springframework.beans.BeanUtils.copyProperties(old,t); t.setCheckFlg(true);
        });
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
        var brands = form.getMitsumoriRows().stream().flatMap(m -> m.getBrandRows().stream())
            .filter(KihonBrandRowForm::isCheckFlg).toList();
        for(var brand : brands) for(var source : repo.findEstimateTerms(brand.getUmKihonBrandId(),brand.getUmMitsumoriId()))
            for(var row : form.getMeisaiRows()) if(same(row.getBrandkoseiId(),brand.getBrandkoseiId()) && same(row.getKikimeisaiId(),source.getKikimeisaiId())) {
                row.setKeiyakunaiyo(source.getKeiyakunaiyo());row.setKeiyakuNo(source.getKeiyakuNo());row.setServicekeitai(source.getServicekeitai());
                row.setTorihosyujikanId(source.getTorihosyujikanId());row.setDaifukuhosyujikanId(source.getDaifukuhosyujikanId());
                row.setTenkenkaisu(source.getTenkenkaisu());row.setTenkenyobi(source.getTenkenyobi());row.setHosyuhoho(source.getHosyuhoho());
            }
        form.setKoseiRows(form.getKoseiRows().stream().filter(k -> brands.stream()
            .anyMatch(b -> same(b.getBrandkoseiId(), k.getBrandkoseiId()))).collect(Collectors.toList()));
        form.setMeisaiRows(form.getMeisaiRows().stream().filter(m -> form.getKoseiRows().stream()
            .anyMatch(k -> same(k.getKikikoseiId(), m.getKikikoseiId()) && same(k.getBrandkoseiId(), m.getBrandkoseiId()))).collect(Collectors.toList()));
        form.setKotaiRows(form.getKotaiRows().stream().filter(t -> form.getMeisaiRows().stream()
            .anyMatch(m -> same(m.getKikimeisaiId(), t.getKikimeisaiId()) && same(m.getBrandkoseiId(), t.getBrandkoseiId()))).collect(Collectors.toList()));
    }

    // ===================================================================
    // Step2→Apply: バリデーション（重複個体チェック）
    // 【変換元】ApplyButton_Click
    // ===================================================================

    public List<String> validateStep2(Mcm2007uForm form) {
        List<String> errors = new ArrayList<>(validateStep1(form));
        var seen = new java.util.HashSet<String>();
        for (var row : form.getMeisaiRows()) if (row.isCheckFlg() && form.getKoseiRows().stream()
                .noneMatch(k -> k.isCheckFlg() && same(k.getKikikoseiId(), row.getKikikoseiId()) && same(k.getBrandkoseiId(), row.getBrandkoseiId())))
            errors.add("機器明細に対応する機器構成を選択してください。");
        for (var row : form.getKotaiRows()) if (row.isCheckFlg()) {
            if (form.getMeisaiRows().stream().noneMatch(m -> m.isCheckFlg() && same(m.getKikimeisaiId(), row.getKikimeisaiId()) && same(m.getBrandkoseiId(), row.getBrandkoseiId())))
                errors.add("個体に対応する機器明細を選択してください。");
            if (row.getKotaikanriId() == null || !seen.add(row.getKikimeisaiId().stripTrailingZeros().toPlainString()+":"+row.getKotaikanriId().stripTrailingZeros().toPlainString()))
                errors.add("個体が重複して選択されています。");
        }
        // 同じ契約の前期間の個体は、期間変更で引き継ぐため重複エラーにしない。
        return errors.stream().distinct().toList();
    }
}
