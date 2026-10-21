package com.daifuku.mcm.service;

import java.math.BigDecimal;
import com.daifuku.mcm.form.Mcm2006uForm;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
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

    private static final org.slf4j.Logger PERF = org.slf4j.LoggerFactory.getLogger(Mcm2007uService.class);

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

            var previewRows = repo.findPreviewRows(allMitsumoriIds);

            // 見積IDでブランドをグルーピング
            for (MitsumoriRowForm m : mitsumoriList) {
                List<KihonBrandRowForm> brands = allBrands.stream()
                    .filter(b -> same(m.getUmMitsumoriId(), b.getUmMitsumoriId()))
                    .collect(Collectors.toList());
                m.setBrandRows(brands);
                m.setPreviewRows(previewRows.stream().filter(row -> m.getUmMitsumoriId().stripTrailingZeros().toPlainString().equals(row.get("UM_MITSUMORI_ID"))).collect(Collectors.toList()));
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
        List<String> errors = new ArrayList<>(validatePeriod(form));
        if (!errors.isEmpty()) return errors;

        // チェック件数0件チェック（VB: MSG_0078 "店舗見積を選択してください。"）
        long checkedCount = form.getMitsumoriRows().stream()
            .flatMap(m -> m.getBrandRows().stream())
            .filter(KihonBrandRowForm::isCheckFlg)
            .count();
        if (checkedCount == 0) {
            errors.add("店舗見積を選択してください。");
        }

        return errors;
    }

    /**
     * 期間（開始日・終了日）のみのチェック。
     * 【変換元】
     * ・開始日/終了日とも、VB版のCPTextBox（ValidateKey="KAISI_DT"）がフォーカスアウト時に
     *   フレームワーク側で自動的に必須・書式チェックを行う（CPTextBox.vb Leaveイベント→CPValidateUtilty.IsErrorCell）。
     *   ValidateKeyが両テキストボックスとも"KAISI_DT"のため、項目名は両方とも「開始日」になる
     *   （CPValidate.xml: KAISI_DT="開始日", mask="yyyy/MM/dd"）。
     * ・nextButton_Click内の明示チェックは終了日のみ、項目名を「期間：終了日」として行う。
     * 上記いずれも「次へ」ボタン押下時に検知できるよう、本メソッドに統合している。
     */
    public List<String> validatePeriod(Mcm2007uForm form) {
        List<String> errors = new ArrayList<>();

        // 開始日必須（VB: FWM_0001 "{0}は必ず入力してください。" {0}="開始日"）
        if (form.getKaisiDt() == null || form.getKaisiDt().isBlank()) {
            errors.add("開始日は必ず入力してください。");
            return errors;
        }

        // 終了日必須（VB: FWM_0001 "{0}は必ず入力してください。" {0}="期間：終了日"）
        if (form.getSyuryoDt() == null || form.getSyuryoDt().isBlank()) {
            errors.add("期間：終了日は必ず入力してください。");
            return errors;
        }

        // 開始日書式チェック（VB: FWM_0006 "{0}は{1}の書式で入力してください。" {0}="開始日"）
        LocalDate kaisi = null, syuryo = null;
        try {
            kaisi = LocalDate.parse(form.getKaisiDt(), DT_FMT);
        } catch (Exception e) {
            errors.add("開始日はYYYY/MM/DDの書式で入力してください。");
        }
        if (!errors.isEmpty()) return errors;

        // 終了日書式チェック（VB: FWM_0006 "{0}は{1}の書式で入力してください。" {0}="期間：終了日"）
        try {
            syuryo = LocalDate.parse(form.getSyuryoDt(), DT_FMT);
        } catch (Exception e) {
            errors.add("期間：終了日はYYYYMMDDの書式で入力してください。");
        }
        if (!errors.isEmpty()) return errors;

        // 開始日≦終了日（VB: MSG_0027）
        if (kaisi != null && syuryo != null && !syuryo.isAfter(kaisi)) {
            errors.add("開始日は終了日よりも前の日付で入力して下さい。");
        }

        // 期間1年以内（VB: MSG_0123 "期間は1年以内で設定してください。"）
        if (kaisi != null && syuryo != null) {
            if (syuryo.isAfter(kaisi.plusYears(1).minusDays(1))) {
                errors.add("期間は1年以内で設定してください。");
            }
        }

        return errors;
    }

    public boolean hasDuplicateCandidates(Mcm2007uForm form) {
        return repo.hasDuplicateIndividuals(form.getMitsumoriRows().stream()
            .filter(MitsumoriRowForm::isCheckFlg).map(MitsumoriRowForm::getUmMitsumoriId).distinct().toList());
    }

    /** VBの既存ブランドのチェック状態を復元する。 */
    public void selectExisting(Mcm2007uForm form, Mcm2006uForm.KikanTabForm tab) {
        form.setUkKikanId(tab.getUkKikanId());
        for (var m : form.getMitsumoriRows()) for (var b : m.getBrandRows()) b.setCheckFlg(
            tab.getBrandRows().stream().anyMatch(old -> same(old.getUmMitsumoriId(), b.getUmMitsumoriId())
                && same(old.getBrandkoseiId(), b.getBrandkoseiId())));
    }
    /** VBの変更開始時は既存契約の選定状態だけを復元し、候補の値は選択見積・マスターから読み込む。 */
    public void selectExistingEquipment(Mcm2007uForm form, Mcm2006uForm.KikanTabForm tab) {
        if (tab == null) return;
        var kosei = new HashSet<Pair>();
        var meisai = new HashSet<Pair>();
        var kotai = new HashSet<Triple>();
        for (var row : tab.getKoseiRows()) add(kosei, pair(row.getBrandkoseiId(), row.getKikikoseiId()));
        for (var row : tab.getMeisaiRows()) add(meisai, pair(row.getBrandkoseiId(), row.getKikimeisaiId()));
        for (var row : tab.getKotaiRows()) add(kotai, triple(row.getBrandkoseiId(), row.getKikimeisaiId(), row.getKotaikanriId()));
        form.getKoseiRows().forEach(r -> r.setCheckFlg(kosei.contains(pair(r.getBrandkoseiId(), r.getKikikoseiId()))));
        form.getMeisaiRows().forEach(r -> r.setCheckFlg(meisai.contains(pair(r.getBrandkoseiId(), r.getKikimeisaiId()))));
        form.getKotaiRows().forEach(r -> r.setCheckFlg(kotai.contains(triple(r.getBrandkoseiId(), r.getKikimeisaiId(), r.getKotaikanriId()))));
    }

    // ===================================================================
    // Step2: 機器構成・明細・個体一覧ロード
    // 【変換元】Mcm2007uScreen2.vb — Form_Load
    // ===================================================================

    public void loadStep2(Mcm2007uForm form) {
        long started = System.nanoTime();
        // Step1で選択された見積IDリストを取得
        List<BigDecimal> selectedMitsumoriIds = form.getMitsumoriRows().stream()
            .flatMap(m -> m.getBrandRows().stream())
            .filter(KihonBrandRowForm::isCheckFlg)
            .map(KihonBrandRowForm::getUmMitsumoriId)
            .distinct()
            .collect(Collectors.toList());

        form.setKoseiRows(repo.findKoseiList(selectedMitsumoriIds, form.getPlantId()));
        long koseiDone = System.nanoTime();
        form.setMeisaiRows(repo.findMeisaiList(selectedMitsumoriIds, form.getPlantId()));
        long meisaiDone = System.nanoTime();
        form.setKotaiRows(repo.findKotaiList(selectedMitsumoriIds, form.getPlantId()));
        long kotaiDone = System.nanoTime();
        var brands = form.getMitsumoriRows().stream().flatMap(m -> m.getBrandRows().stream())
            .filter(KihonBrandRowForm::isCheckFlg).toList();
        var terms = new LinkedHashMap<Pair, List<Mcm2007uForm.MeisaiRowForm>>();
        for (var source : repo.findEstimateTerms(brands)) {
            Pair key = pair(source.basicBrandId(), source.estimateId());
            if (key != null) terms.computeIfAbsent(key, ignored -> new ArrayList<>()).add(source.row());
        }
        // 同一組の同一明細が複数あるデータだけは、元SQLの取得順をそのまま使う。
        // 通常の一意な明細は一括取得し、ブランド単位のSQL往復をなくす。
        var checkedTerms = new HashSet<Pair>();
        for (var brand : brands) {
            Pair key = pair(brand.getUmKihonBrandId(), brand.getUmMitsumoriId());
            var rows = terms.get(key);
            if (checkedTerms.add(key) && rows != null && hasRepeatedDetail(rows))
                terms.put(key, repo.findEstimateTerms(brand.getUmKihonBrandId(), brand.getUmMitsumoriId()));
        }
        long termsDone = System.nanoTime();
        var targets = new HashMap<Pair, List<Mcm2007uForm.MeisaiRowForm>>();
        for (var row : form.getMeisaiRows()) {
            Pair key = pair(row.getBrandkoseiId(), row.getKikimeisaiId());
            if (key != null) targets.computeIfAbsent(key, ignored -> new ArrayList<>()).add(row);
        }
        // 元のブランド選択順と条件行の順を維持し、後の選択見積の条件が勝つ規則を変えない。
        for (var brand : brands) for (var source : terms.getOrDefault(
                pair(brand.getUmKihonBrandId(), brand.getUmMitsumoriId()), List.of()))
            for (var row : targets.getOrDefault(pair(brand.getBrandkoseiId(), source.getKikimeisaiId()), List.of())) {
                row.setKeiyakunaiyo(source.getKeiyakunaiyo());row.setKeiyakuNo(source.getKeiyakuNo());row.setServicekeitai(source.getServicekeitai());
                row.setTorihosyujikanId(source.getTorihosyujikanId());row.setDaifukuhosyujikanId(source.getDaifukuhosyujikanId());
                row.setTenkenkaisu(source.getTenkenkaisu());row.setTenkenyobi(source.getTenkenyobi());row.setHosyuhoho(source.getHosyuhoho());
            }
        var selectedBrandIds = new HashSet<String>();
        for (var brand : brands) if (brand.getBrandkoseiId() != null) selectedBrandIds.add(id(brand.getBrandkoseiId()));
        form.setKoseiRows(form.getKoseiRows().stream().filter(k -> k.getBrandkoseiId() != null
            && selectedBrandIds.contains(id(k.getBrandkoseiId()))).collect(Collectors.toList()));
        var koseiKeys = new HashSet<Pair>();
        for (var row : form.getKoseiRows()) add(koseiKeys, pair(row.getBrandkoseiId(), row.getKikikoseiId()));
        form.setMeisaiRows(form.getMeisaiRows().stream().filter(m -> koseiKeys.contains(
            pair(m.getBrandkoseiId(), m.getKikikoseiId()))).collect(Collectors.toList()));
        var meisaiKeys = new HashSet<Pair>();
        for (var row : form.getMeisaiRows()) add(meisaiKeys, pair(row.getBrandkoseiId(), row.getKikimeisaiId()));
        form.setKotaiRows(form.getKotaiRows().stream().filter(t -> meisaiKeys.contains(
            pair(t.getBrandkoseiId(), t.getKikimeisaiId()))).collect(Collectors.toList()));
        long hierarchyDone = System.nanoTime();
        var sources=repo.findSelectionSources(selectedMitsumoriIds);
        form.setStep2Brands(repo.findPlantBrands(form.getPlantId()));
        long sourcesDone = System.nanoTime();
        var numbers = indexNumbers(sources);
        var firstBrand = new HashMap<String, KihonBrandRowForm>();
        for (var brand : brands) firstBrand.putIfAbsent(id(brand.getBrandkoseiId()), brand);
        for(var row:form.getStep2Brands()){
            var selected=firstBrand.get(row.get("BRANDKOSEI_ID"));
            if(selected != null){
                row.putAll(selected.getDisplay());
                row.put("UM_MITSUMORI_NO",numbers.getOrDefault(new NumberKey(1,row.get("BRANDKOSEI_ID"),null,null,null),""));
            }
        }
        for(var row:form.getKoseiRows()) {String value=numbers.getOrDefault(new NumberKey(2,id(row.getBrandkoseiId()),id(row.getKikikoseiId()),null,null),"");setNumbers(row.getDisplay(),value);row.setQuoted(row.isQuoted()&&!value.isBlank());}
        for(var row:form.getMeisaiRows()) {String value=numbers.getOrDefault(new NumberKey(3,id(row.getBrandkoseiId()),id(row.getKikikoseiId()),id(row.getKikimeisaiId()),null),"");setNumbers(row.getDisplay(),value);row.setQuoted(row.isQuoted()&&!value.isBlank());}
        for(var row:form.getKotaiRows()) {String value=numbers.getOrDefault(new NumberKey(4,id(row.getBrandkoseiId()),id(row.getKikikoseiId()),id(row.getKikimeisaiId()),id(row.getKotaikanriId())),"");setNumbers(row.getDisplay(),value);row.setQuoted(row.isQuoted()&&!value.isBlank());}
        form.setSelectedBrandIndex(-1);form.setSelectedKoseiIndex(-1);
        long done = System.nanoTime();
        PERF.info("MCM2007U loadStep2 totalMs={} koseiMs={} meisaiMs={} kotaiMs={} termsMs={} hierarchyMs={} sourcesMs={} numbersMs={} brands={} kosei={} meisai={} kotai={} sources={}",
            ms(started,done),ms(started,koseiDone),ms(koseiDone,meisaiDone),ms(meisaiDone,kotaiDone),ms(kotaiDone,termsDone),
            ms(termsDone,hierarchyDone),ms(hierarchyDone,sourcesDone),ms(sourcesDone,done),brands.size(),form.getKoseiRows().size(),form.getMeisaiRows().size(),form.getKotaiRows().size(),sources.size());
    }

    private static String id(BigDecimal value){return value==null?"":value.stripTrailingZeros().toPlainString();}
    private record Pair(String first, String second) { }
    private record Triple(String first, String second, String third) { }
    private record NumberKey(int depth, String brand, String kosei, String detail, String individual) { }
    private static Pair pair(BigDecimal a, BigDecimal b) { return a == null || b == null ? null : new Pair(id(a),id(b)); }
    private static Triple triple(BigDecimal a, BigDecimal b, BigDecimal c) { return a == null || b == null || c == null ? null : new Triple(id(a),id(b),id(c)); }
    private static <T> void add(Set<T> keys, T key) { if (key != null) keys.add(key); }
    private static long ms(long from, long to) { return (to-from)/1_000_000; }
    private static void setNumbers(Map<String,String> display, String value) { display.put("UM_MITSUMORI_NO",value); }
    private static boolean hasRepeatedDetail(List<Mcm2007uForm.MeisaiRowForm> rows) {
        var seen = new HashSet<String>();
        return rows.stream().anyMatch(r -> r.getKikimeisaiId() != null && !seen.add(id(r.getKikimeisaiId())));
    }
    private static Map<NumberKey,String> indexNumbers(List<Map<String,String>> sources) {
        var grouped = new HashMap<NumberKey,LinkedHashSet<String>>();
        for (var source : sources) {
            String number=source.get("UM_MITSUMORI_NO");
            if (number==null || number.isBlank()) continue;
            String b=source.get("BRANDKOSEI_ID"),k=source.get("KIKIKOSEI_ID"),m=source.get("KIKIMEISAI_ID"),t=source.get("KOTAIKANRI_ID");
            for (var key : List.of(new NumberKey(1,b,null,null,null),new NumberKey(2,b,k,null,null),
                    new NumberKey(3,b,k,m,null),new NumberKey(4,b,k,m,t)))
                grouped.computeIfAbsent(key,ignored -> new LinkedHashSet<>()).add(number);
        }
        var result = new HashMap<NumberKey,String>();
        grouped.forEach((key,values) -> result.put(key,String.join(", ",values)));
        return result;
    }
    /** VBの個体選定数から適用時のセット数・数量を再計算する。 */
    public void prepareDeliveryQuantities(Mcm2007uForm form){
        for(var k:form.getKoseiRows())if(k.isCheckFlg()){
            var counts=form.getMeisaiRows().stream().filter(m->m.isCheckFlg()&&same(k.getBrandkoseiId(),m.getBrandkoseiId())&&same(k.getKikikoseiId(),m.getKikikoseiId())&&"1".equals(m.getDisplay().get("KOTAIKANRI_FLG")))
                .map(m->form.getKotaiRows().stream().filter(t->t.isCheckFlg()&&same(t.getBrandkoseiId(),m.getBrandkoseiId())&&same(t.getKikikoseiId(),m.getKikikoseiId())&&same(t.getKikimeisaiId(),m.getKikimeisaiId())).count()).toList();
            k.setSetNm(counts.isEmpty()||counts.stream().distinct().count()!=1?"1":Long.toString(counts.get(0)));
        }
        for(var m:form.getMeisaiRows())if(m.isCheckFlg()){
            var individuals=form.getKotaiRows().stream().filter(t->same(t.getBrandkoseiId(),m.getBrandkoseiId())&&same(t.getKikikoseiId(),m.getKikikoseiId())&&same(t.getKikimeisaiId(),m.getKikimeisaiId())).toList();
            // 個体展開されない機器の数量は既存値を保持する。
            if(!individuals.isEmpty())m.setSuryoNm(individuals.stream().filter(Mcm2007uForm.KotaiRowForm::isCheckFlg).map(t->number(t.getDisplay().get("SURYO_NM"))).reduce(BigDecimal.ZERO,BigDecimal::add).stripTrailingZeros().toPlainString());
        }
    }
    private static BigDecimal number(String value){return value==null||value.isBlank()?BigDecimal.ZERO:new BigDecimal(value);}

    // ===================================================================
    // Step2→Apply: バリデーション（重複個体チェック）
    // 【変換元】ApplyButton_Click
    // ===================================================================

    public List<String> validateStep2(Mcm2007uForm form) {
        List<String> errors = new ArrayList<>(validateStep1(form));
        var seen = new java.util.HashSet<String>();
        // VB: MSG_0081 "前画面で、選択した取引先見積に含まれていない為、選択する事は出来ません。"
        if(form.getKoseiRows().stream().anyMatch(r->r.isCheckFlg()&&!r.isQuoted())||form.getMeisaiRows().stream().anyMatch(r->r.isCheckFlg()&&!r.isQuoted())||form.getKotaiRows().stream().anyMatch(r->r.isCheckFlg()&&!r.isQuoted()))errors.add("前画面で、選択した取引先見積に含まれていない為、選択する事は出来ません。");
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
