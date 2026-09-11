package com.daifuku.mcm.service;

import java.math.BigDecimal;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.daifuku.mcm.constants.Mcm2006uConstants;
import com.daifuku.mcm.form.Mcm2006uForm;
import com.daifuku.mcm.form.Mcm2006uForm.BrandRowForm;
import com.daifuku.mcm.form.Mcm2006uForm.KikanTabForm;
import com.daifuku.mcm.form.Mcm2006uForm.KotaiRowForm;
import com.daifuku.mcm.form.Mcm2006uForm.KoseiRowForm;
import com.daifuku.mcm.form.Mcm2006uForm.MeisaiRowForm;
import com.daifuku.mcm.form.Mcm2006uForm.SeibanRowForm;
import com.daifuku.mcm.form.Mcm2007uForm;
import com.daifuku.mcm.repository.Mcm2006uRepository;

/**
 * 【変換元】Mcm2006uTabControl.vb — SearchRead / UpdateButtonTabNaiyo / View
 * MCM2006U ユーザ契約内容サービス
 *
 * Web版省略事項:
 *   - SP_UKストアドプロシージャ
 *   - ロック解除ボタン
 *   - タブ間日付自動連携（隣接タブの開始日/終了日を自動補完する機能）
 */
@Service
public class Mcm2006uService {

    @Autowired
    private Mcm2006uRepository repo;

    private static final DateTimeFormatter DT_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");

    // ===================================================================
    // ロード: 既存契約参照
    // 【変換元】SearchRead() — 全テーブルFill
    // ===================================================================

    public Mcm2006uForm load(BigDecimal ukKeiyakuId) {
        Mcm2006uForm form = repo.findKeiyaku(ukKeiyakuId);
        if (form.getUkKeiyakuId() == null) throw new IllegalStateException("契約情報が見つかりません。再検索してください。");
        form.setSeibanRows(repo.findSeibanList(ukKeiyakuId));

        List<KikanTabForm> tabs = repo.findKikanList(ukKeiyakuId);
        for (KikanTabForm tab : tabs) {
            tab.setBrandRows(repo.findBrandRows(tab.getUkKikanId()));
            tab.setKoseiRows(repo.findKoseiRows(tab.getUkKikanId()));
            tab.setMeisaiRows(repo.findMeisaiRows(tab.getUkKikanId()));
            tab.setKotaiRows(repo.findKotaiRows(tab.getUkKikanId()));
            tab.setTenkenRows(repo.findTenkenRows(tab.getUkKikanId()));
            tab.setTabFlg(calcTabFlg(tab, tabs));
        }
        form.setKikanTabs(tabs);
        if (!tabs.isEmpty()) {
            form.setSelectedKikanId(tabs.get(tabs.size()-1).getUkKikanId());
        }
        return form;
    }

    // ===================================================================
    // ロード: MCM2007U戻り後に期間タブを追加
    // 【変換元】TabHenkoButtonClick — SendDelivery戻り後のタブ追加
    // ===================================================================

    public void applyPendingKikan(Mcm2006uForm form, KikanTabForm pending) {
        var previous = form.getSelectedTab();
        if (previous != null) {
            // 変更開始日以降を新しい期間へ移し、同じ開始日の変更は期間自体を置換する。
            var start = date(pending.getKaisiDt());
            if (start.isBefore(date(previous.getKaisiDt())) || start.isAfter(date(previous.getSyuryoDt()).plusDays(1)))
                throw new IllegalStateException("変更摘要開始日が選択期間の範囲外です。");
            if (start.equals(date(previous.getKaisiDt()))) {
                if (previous.getUkKikanId() != null && previous.getUkKikanId().signum() > 0) {
                    pending.setUkKikanId(previous.getUkKikanId());
                    pending.setReplaceEquipment(true);
                }
                form.getKikanTabs().remove(previous);
            } else previous.setSyuryoDt(format(start.minusDays(1)));
            pending.setHosyuGkin(previous.getHosyuGkin());
            pending.setHosyuhoho(previous.getHosyuhoho());
            pending.setBiko(previous.getBiko());
            pending.setIraitenpoId(previous.getIraitenpoId());
            pending.setIraimeisho1Nk(previous.getIraimeisho1Nk());
            pending.setIraimeisho2Nk(previous.getIraimeisho2Nk());
            pending.setIraimeisho3Nk(previous.getIraimeisho3Nk());
            pending.setIraimeisho4Nk(previous.getIraimeisho4Nk());
            pending.setIraitenporyakuNk(previous.getIraitenporyakuNk());
            pending.setIraitantoNk(previous.getIraitantoNk());
        }
        if (pending.getUkKikanId() == null) pending.setUkKikanId(form.getKikanTabs().stream().map(KikanTabForm::getUkKikanId).filter(java.util.Objects::nonNull).min(BigDecimal::compareTo).orElse(BigDecimal.ZERO).min(BigDecimal.ZERO).subtract(BigDecimal.ONE));
        pending.setUkKeiyakuId(form.getUkKeiyakuId());
        pending.setTabFlg(Mcm2006uConstants.TAB_FLG_NEW);
        pending.setYukoFlg(Mcm2006uConstants.YUKO_FLG_OFF);
        if(previous==null) {
            if(form.getNonyubusyoNk()==null)form.setNonyubusyoNk(pending.getNonyubusyoNk());
            if(form.getNonyutantosyaNk()==null)form.setNonyutantosyaNk(pending.getNonyutantosyaNk());
            if(form.getNonyutelNo()==null)form.setNonyutelNo(pending.getNonyutelNo());
            if(form.getNonyufaxNo()==null)form.setNonyufaxNo(pending.getNonyufaxNo());
        }
        pending.setNonyubusyoNk(form.getNonyubusyoNk());
        pending.setNonyutantosyaNk(form.getNonyutantosyaNk());
        pending.setNonyutelNo(form.getNonyutelNo()); pending.setNonyufaxNo(form.getNonyufaxNo());
        form.getKikanTabs().add(pending);
        form.getKikanTabs().sort(java.util.Comparator.comparing(t -> date(t.getKaisiDt())));
        form.setSelectedKikanId(pending.getUkKikanId());
    }

    // ===================================================================
    // 保存
    // 【変換元】UpdateButton_Click → UpdateButtonTabNaiyo()
    //
    // 処理順:
    //   1. MCM_UK_KEIYAKU (INSERT/UPDATE)
    //   2. MCM_UK_SEIBAN (INSERT/UPDATE)
    //   3. 各タブ: MCM_UK_KIKAN + MCM_UK_BRAND + MCM_UK_KIKIKOSEI + MCM_UK_KIKIMEISAI + MCM_UK_KOTAIMEISAI
    //   4. 次回更新日設定 (最終タブのSYURYO_DT + 1日)
    // ===================================================================

    @Transactional
    public void save(Mcm2006uForm form, String loginUser) {

        if (form.getPlantId() == null || form.getNonyusakiId() == null || form.getKikanTabs().isEmpty())
            throw new IllegalStateException("契約の対象と期間を選択してください。");
        if (isReadOnly(form.getJotai(), form.getSeniMotoKbn())) throw new IllegalStateException(DENIED);
        if (form.getUkKeiyakuId() != null) {
            repo.lock(form.getUkKeiyakuId());
            var current = repo.findKeiyaku(form.getUkKeiyakuId());
            if(!java.util.Objects.equals(current.getVersion(),form.getVersion()))throw new IllegalStateException("他のユーザによって更新されています。再検索してください。");
            if (current.getUkKeiyakuId() == null || isReadOnly(current.getJotai(), form.getSeniMotoKbn())) throw new IllegalStateException(EXPIRED);
        }
        form.getKikanTabs().sort(java.util.Comparator.comparing(t -> date(t.getKaisiDt())));
        LocalDate end = null;
        for (var tab : form.getKikanTabs()) {
            var start = date(tab.getKaisiDt()); var finish = date(tab.getSyuryoDt());
            if (!finish.isAfter(start) || finish.isAfter(start.plusYears(1).minusDays(1)))
                throw new IllegalStateException("期間は開始日より後、1年以内の終了日を指定してください。");
            if (end != null && !start.equals(end.plusDays(1))) throw new IllegalStateException("契約期間が連続するように設定してください。");
            if (tab.getHosyuGkin() != null && tab.getHosyuGkin().signum() < 0) throw new IllegalStateException("保守額には0以上の値を入力してください。");
            end = finish;
        }
        for(var period:form.getKikanTabs()) for(var brand:period.getBrandRows())
            if(brand.getUkBrandId()==null && !repo.validSource(brand,form.getPlantId()))
                throw new IllegalStateException("選択した見積の状態が変更されています。選択し直してください。");
        // 1. MCM_UK_KEIYAKU
        boolean isNewKeiyaku = !repo.existsKeiyaku(form.getUkKeiyakuId());
        if (isNewKeiyaku) {
            if (form.getUkKeiyakuId() == null) {
                form.setUkKeiyakuId(repo.nextKeiyakuId());
            }
            if (form.getJotai() == null) {
                form.setJotai(Mcm2006uConstants.JOTAI_KEIYAKU);
            }
            if (form.getKeiyakuDt() == null) form.setKeiyakuDt(form.getKikanTabs().get(0).getKaisiDt());
            if (form.getShokaiKeiyakuDt() == null) form.setShokaiKeiyakuDt(form.getKikanTabs().get(0).getKaisiDt());
            form.setAutoFlg("1");
            repo.insertKeiyaku(form, loginUser);
        } else {
            repo.updateKeiyaku(form, loginUser);
        }

        // 2. MCM_UK_SEIBAN
        for (SeibanRowForm seiban : form.getSeibanRows()) {
            seiban.setUkKeiyakuId(form.getUkKeiyakuId());
            if (seiban.getUkSeibanId() == null) {
                seiban.setUkSeibanId(repo.nextSeibanId());
                repo.insertSeiban(seiban, loginUser);
            } else {
                repo.updateSeiban(seiban, loginUser);
            }
        }

        // 3. 各期間タブ
        List<KikanTabForm> tabs = form.getKikanTabs();
        String lastSyuryoDt = null;
        for (KikanTabForm tab : tabs) {
            tab.setUkKeiyakuId(form.getUkKeiyakuId());
            saveKikan(tab, loginUser);
            lastSyuryoDt = tab.getSyuryoDt();
        }

        // 4. 次回更新日 = 最終タブSYURYO_DT + 1日
        if (lastSyuryoDt != null && !lastSyuryoDt.isBlank()) {
            repo.updateJikaikosinDt(form.getUkKeiyakuId(), lastSyuryoDt, loginUser);
        }
    }

    // ===================================================================
    // 期間タブ1件保存
    // ===================================================================

    private void saveKikan(KikanTabForm tab, String loginUser) {
        boolean isNewKikan = !repo.existsKikan(tab.getUkKikanId());
        if (isNewKikan) {
            if (tab.getUkKikanId() == null || tab.getUkKikanId().signum() <= 0) {
                tab.setUkKikanId(repo.nextKikanId());
            }
            if (tab.getYukoFlg() == null) {
                tab.setYukoFlg(Mcm2006uConstants.YUKO_FLG_OFF);
            }
            repo.insertKikan(tab, loginUser);
        } else {
            repo.updateKikan(tab, loginUser);
        }

        if (tab.isReplaceEquipment()) repo.deleteEquipment(tab.getUkKikanId());
        // ブランド行を保存（新規INSERTのみ: MCM2007U経由で追加されたもの）
        BigDecimal ukBrandIdCounter = null;
        for (BrandRowForm brand : tab.getBrandRows()) {
            brand.setUkKikanId(tab.getUkKikanId());
            if (brand.getUkBrandId() == null) {
                if (ukBrandIdCounter == null) {
                    ukBrandIdCounter = repo.nextBrandId();
                }
                brand.setUkBrandId(ukBrandIdCounter);
                ukBrandIdCounter = ukBrandIdCounter.add(BigDecimal.ONE);
                repo.insertBrand(brand, loginUser);
                repo.insertMitsumoriLink(brand);
                // 対応する機器構成・明細・個体を保存
                saveKoseiForBrand(brand, tab, loginUser);
            }
        }
    }

    private void saveKoseiForBrand(BrandRowForm brand, KikanTabForm tab, String loginUser) {
        BigDecimal ukKikoseiIdCounter = repo.nextKikoseiId();
        for (KoseiRowForm kosei : tab.getKoseiRows()) {
            if (!same(brand.getBrandkoseiId(), kosei.getBrandkoseiId())) continue;
            if (kosei.getUkKikoseiId() == null) {
                kosei.setUkBrandId(brand.getUkBrandId());
                kosei.setUkKikoseiId(ukKikoseiIdCounter);
                ukKikoseiIdCounter = ukKikoseiIdCounter.add(BigDecimal.ONE);
                repo.insertKosei(kosei, loginUser);
                saveMeisaiForKosei(kosei, tab, loginUser);
            }
        }
    }

    private void saveMeisaiForKosei(KoseiRowForm kosei, KikanTabForm tab, String loginUser) {
        BigDecimal ukKikimeisaiIdCounter = repo.nextKikimeisaiId();
        for (MeisaiRowForm meisai : tab.getMeisaiRows()) {
            if (!same(kosei.getKikikoseiId(), meisai.getKikikoseiId()) || !same(kosei.getBrandkoseiId(), meisai.getBrandkoseiId())) continue;
            if (meisai.getUkKikimeisaiId() == null) {
                meisai.setUkKikoseiId(kosei.getUkKikoseiId());
                meisai.setUkKikimeisaiId(ukKikimeisaiIdCounter);
                ukKikimeisaiIdCounter = ukKikimeisaiIdCounter.add(BigDecimal.ONE);
                repo.insertMeisai(meisai, loginUser);
                saveKotaiForMeisai(meisai, tab, loginUser);
            }
        }
    }

    private void saveKotaiForMeisai(MeisaiRowForm meisai, KikanTabForm tab, String loginUser) {
        BigDecimal ukKotaimeisaiIdCounter = repo.nextKotaimeisaiId();
        for (KotaiRowForm kotai : tab.getKotaiRows()) {
            if (!same(meisai.getKikimeisaiId(), kotai.getKikimeisaiId()) || !same(meisai.getBrandkoseiId(), kotai.getBrandkoseiId())) continue;
            if (kotai.getUkKotaimeisaiId() == null) {
                kotai.setUkKikimeisaiId(meisai.getUkKikimeisaiId());
                kotai.setUkKikoseiId(meisai.getUkKikoseiId());
                kotai.setUkKotaimeisaiId(ukKotaimeisaiIdCounter);
                ukKotaimeisaiIdCounter = ukKotaimeisaiIdCounter.add(BigDecimal.ONE);
                repo.insertKotai(kotai, loginUser);
            }
        }
    }

    // ===================================================================
    // ReadOnly判定
    // 【変換元】View() — 廃棄(JOTAI=3) or 解約(JOTAI=4) or SENIMOTO_SHONIN
    // ===================================================================

    public boolean isReadOnly(String jotai, int seniMotoKbn) {
        if ((seniMotoKbn == Mcm2006uConstants.SENIMOTO_SHONIN || seniMotoKbn == 4)) return true;
        return Mcm2006uConstants.JOTAI_HAKI.equals(jotai)
            || Mcm2006uConstants.JOTAI_KAIYAKU.equals(jotai);
    }

    /** タブが編集可かどうか (TAB_FLG_NON/OLD のタブはReadOnly) */
    public boolean isTabReadOnly(KikanTabForm tab) {
        return tab.getTabFlg() == Mcm2006uConstants.TAB_FLG_NON;
    }

    // ===================================================================
    // MCM2007UのDeliveryからKikanTabFormを構築
    // 【変換元】Mcm2006uScreen.vb — TabHenkoButtonClick 後の処理
    // ===================================================================

    public KikanTabForm buildKikanFromDelivery(Mcm2007uForm delivery) {
        var checked=getCheckedBrands(delivery);
        KikanTabForm tab = checked.isEmpty()?new KikanTabForm():repo.findSourceDefaults(checked.get(0).getUmMitsumoriId());
        tab.setHosyuGkin(checked.stream().map(Mcm2007uForm.KihonBrandRowForm::getHoshuKin).filter(java.util.Objects::nonNull).reduce(BigDecimal.ZERO,BigDecimal::add));
        if(!checked.isEmpty())tab.setHosyuhoho(checked.get(0).getHosyuhoho());
        tab.setKaisiDt(delivery.getKaisiDt());
        tab.setSyuryoDt(delivery.getSyuryoDt());
        tab.setNonyusakiId(delivery.getNonyusakiId());
        tab.setNonyusakiCd(delivery.getNonyusakiCd());
        tab.setNonyusakiNk(delivery.getNonyusakiNk());
        tab.setNonyusakijusyo1Nk(delivery.getNonyusakijusyo1Nk());
        tab.setNonyusakijusyo2Nk(delivery.getNonyusakijusyo2Nk());
        tab.setPlantId(delivery.getPlantId());
        tab.setSupportId(delivery.getSupportId());
        tab.setPlantNk(delivery.getPlantNk());
        tab.setTabFlg(Mcm2006uConstants.TAB_FLG_NEW);
        tab.setYukoFlg(Mcm2006uConstants.YUKO_FLG_OFF);

        // ブランド行 (MCM2007U Step1 で選択されたもの)
        List<BrandRowForm> brandRows = new ArrayList<>();
        for (Mcm2007uForm.KihonBrandRowForm src : getCheckedBrands(delivery)) {
            BrandRowForm b = new BrandRowForm();
            b.setBrandkoseiId(src.getBrandkoseiId());
            b.setUmMitsumoriId(src.getUmMitsumoriId());
            b.setUmKihonBrandId(src.getUmKihonBrandId());
            b.setBrandNk(src.getBrandNk());
            b.setBrandsyosaiNk(src.getBrandsyosaiNk());
            b.setKeiyakujikantai(src.getKeiyakujikantai());
            b.setHosyuhoho(src.getHosyuhoho());
            brandRows.add(b);
        }
        tab.setBrandRows(brandRows);
        if (!brandRows.isEmpty()) tab.setKeiyakujikantai(brandRows.get(0).getKeiyakujikantai());

        // 機器構成行 (MCM2007U Step2 で選択されたもの)
        List<KoseiRowForm> koseiRows = new ArrayList<>();
        for (Mcm2007uForm.KoseiRowForm src : delivery.getKoseiRows()) {
            if (!src.isCheckFlg()) continue;
            KoseiRowForm k = new KoseiRowForm();
            k.setKikikoseiId(src.getKikikoseiId());
            k.setBrandkoseiId(src.getBrandkoseiId());
            k.setKikikoseiNk(src.getKikikoseiNk());
            k.setSetNm(src.getSetNm());
            k.setTani(src.getTani());
            k.setTehaiseiban(src.getTehaiseiban());
            k.setControllerFlg(src.getControllerFlg());
            k.setHosyuhoho(src.getHosyuhoho());
            k.setHyojijun(src.getHyojijun());
            koseiRows.add(k);
        }
        tab.setKoseiRows(koseiRows);

        // 機器明細行 (MCM2007U Step2 で選択されたもの)
        List<MeisaiRowForm> meisaiRows = new ArrayList<>();
        for (Mcm2007uForm.MeisaiRowForm src : delivery.getMeisaiRows()) {
            if (!src.isCheckFlg()) continue;
            MeisaiRowForm m = new MeisaiRowForm();
            m.setKikimeisaiId(src.getKikimeisaiId());
            m.setKikikoseiId(src.getKikikoseiId());
            m.setBrandkoseiId(src.getBrandkoseiId());
            m.setSeizomakerId(src.getSeizomakerId());
            m.setSeizomankerNk(src.getSeizomankerNk());
            m.setKikihinmeiNk(src.getKikihinmeiNk());
            m.setKikikatashiki(src.getKikikatashiki());
            m.setSuryoNm(src.getSuryoNm());
            m.setKeiyakunaiyo(src.getKeiyakunaiyo());
            m.setKeiyakuNo(src.getKeiyakuNo());
            m.setServicekeitai(src.getServicekeitai());
            m.setTorihosyujikanId(src.getTorihosyujikanId());
            m.setDaifukuhosyujikanId(src.getDaifukuhosyujikanId());
            m.setTenkenkaisu(src.getTenkenkaisu());
            m.setTenkenyobi(src.getTenkenyobi());
            m.setHosyuhoho(src.getHosyuhoho());
            m.setMaeHyojijun(src.getMaeHyojijun());
            m.setHyojijun(src.getHyojijun());
            meisaiRows.add(m);
        }
        tab.setMeisaiRows(meisaiRows);

        // 個体行 (MCM2007U Step2 で選択されたもの)
        List<KotaiRowForm> kotaiRows = new ArrayList<>();
        for (Mcm2007uForm.KotaiRowForm src : delivery.getKotaiRows()) {
            if (!src.isCheckFlg()) continue;
            KotaiRowForm kt = new KotaiRowForm();
            kt.setKotaikanriId(src.getKotaikanriId());
            kt.setKikimeisaiId(src.getKikimeisaiId());
            kt.setKikikoseiId(src.getKikikoseiId());
            kt.setBrandkoseiId(src.getBrandkoseiId());
            kt.setKotaiNk(src.getKotaiNk());
            kt.setSerialNo(src.getSerialNo());
            kt.setItizinonnyuDt(src.getItizinonnyuDt());
            kt.setSetchibasyo(src.getSetchibasyo());
            kotaiRows.add(kt);
        }
        tab.setKotaiRows(kotaiRows);

        return tab;
    }

    private List<Mcm2007uForm.KihonBrandRowForm> getCheckedBrands(Mcm2007uForm delivery) {
        List<Mcm2007uForm.KihonBrandRowForm> result = new ArrayList<>();
        for (Mcm2007uForm.MitsumoriRowForm m : delivery.getMitsumoriRows()) {
            for (Mcm2007uForm.KihonBrandRowForm b : m.getBrandRows()) {
                if (b.isCheckFlg()) result.add(b);
            }
        }
        return result;
    }

    // ===================================================================
    // TabFlg計算
    // 現在日付を基準に各タブの役割を決定する
    // ===================================================================

    private int calcTabFlg(KikanTabForm tab, List<KikanTabForm> allTabs) {
        if (tab.getSyuryoDt() == null || tab.getSyuryoDt().isBlank()) {
            return Mcm2006uConstants.TAB_FLG_NON;
        }
        LocalDate today = LocalDate.now();
        LocalDate kaisi = parseDate(tab.getKaisiDt());
        LocalDate syuryo = parseDate(tab.getSyuryoDt());

        // 最新タブ（他のタブよりSYURYO_DTが最大）かどうか
        boolean isLatest = allTabs.stream().noneMatch(other ->
            other != tab && parseDate(other.getSyuryoDt()).isAfter(syuryo));

        if (!isLatest) {
            // 最新でないタブはNON（ReadOnly）
            return Mcm2006uConstants.TAB_FLG_NON;
        }
        // 最新タブ: 現在日付との関係で分類
        if (!today.isAfter(syuryo) && !today.isBefore(kaisi)) {
            return Mcm2006uConstants.TAB_FLG_NOW;   // 現在期間
        } else if (kaisi.isAfter(today)) {
            return Mcm2006uConstants.TAB_FLG_FUTER; // 将来期間
        } else {
            return Mcm2006uConstants.TAB_FLG_OLD;   // 過去期間（変更可）
        }
    }

    private LocalDate parseDate(String dt) {
        if (dt == null || dt.isBlank()) return LocalDate.MIN;
        try { return LocalDate.parse(dt, DT_FMT); }
        catch (Exception e) { return LocalDate.MIN; }
    }
}
