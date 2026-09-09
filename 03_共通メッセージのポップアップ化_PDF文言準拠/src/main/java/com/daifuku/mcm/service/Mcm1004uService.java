/**
 * 【変換元】Mcm1004uScreen.vb + Mcm1004uTabControl.vb
 *   需要家契約一覧 Service - 検索/バリデーション/更新ロジック
 */
package com.daifuku.mcm.service;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.daifuku.mcm.common.AppConstants;
import com.daifuku.mcm.common.BaseService;
import com.daifuku.mcm.entity.McmTmKikanEntity;
import com.daifuku.mcm.form.Mcm1004uForm;
import com.daifuku.mcm.form.Mcm1004uTabForm;
import com.daifuku.mcm.repository.McmTmKeiyakujikanRepository;
import com.daifuku.mcm.repository.McmTmKikanRepository;
import com.daifuku.mcm.repository.McmTmMitsumoriRepository;
import com.daifuku.mcm.repository.McmTmTankaRepository;
import com.daifuku.mcm.repository.McmTmTenkenRepository;

import lombok.RequiredArgsConstructor;

@Service
@RequiredArgsConstructor
public class Mcm1004uService extends BaseService {

    private final McmTmKikanRepository kikanRepo;
    private final McmTmKeiyakujikanRepository keiyakuRepo;
    private final McmTmMitsumoriRepository mitsuRepo;
    private final McmTmTankaRepository tankaRepo;
    private final McmTmTenkenRepository tenkenRepo;

    // ========== 画面読込 ==========

    /**
     * 【変換元】Mcm1004uScreen.Search() + View()
     *   画面データ読込 - 見積+契約時間帯取得、viewFlg判定、タブ構築
     */
    public Mcm1004uForm loadScreen(BigDecimal tmKeiyakujikanId, int seniMotoKbn, boolean lockRelease) {
        Mcm1004uForm form = new Mcm1004uForm();
        form.setTmKeiyakujikanId(tmKeiyakujikanId);
        form.setSeniMotoKbn(seniMotoKbn);

        // 【変換元】MCM_TM_MITSUMORITableAdapter.Fill + MCM_TM_KEIYAKUJIKANTableAdapter.Fill
        List<Object[]> mitsumoriData = mitsuRepo.findMitsumoriWithKeiyakujikan(tmKeiyakujikanId);
        if (mitsumoriData.isEmpty()) { form.setDataNotFound(true); return form; }

        // TODO: mitsumoriData[0]からヘッダ情報をformにセット

        // 【変換元】viewFlg判定
        boolean viewFlg = determineViewFlag(
            form.getJotai(), tmKeiyakujikanId, form.getPlantId(), lockRelease);
        form.setViewFlg(viewFlg);

        // 【変換元】期間タブ構築
        List<McmTmKikanEntity> kikanList = kikanRepo
            .findByKeiyakujikan_TmKeiyakujikanIdOrderByKaisiDt(tmKeiyakujikanId);
        List<Mcm1004uTabForm> tabs = new ArrayList<>();
        if (kikanList.isEmpty()) {
            Mcm1004uTabForm tab = new Mcm1004uTabForm();
            tab.setTabLabel("初回");
            tab.setTmKikanId(Long.valueOf(0));
            tabs.add(tab);
        } else {
            for (int i = 0; i < kikanList.size(); i++) {
                McmTmKikanEntity kikan = kikanList.get(i);
                Mcm1004uTabForm tab = new Mcm1004uTabForm();
                tab.setTabLabel(i == 0 ? "初回" : (i + 1) + "回");
                tab.setTmKikanId(kikan.getTmKikanId());
                tab.setKaisiDt(kikan.getKaisiDt() != null ? kikan.getKaisiDt().toLocalDate() : null);
                tab.setSyuryoDt(kikan.getSyuryoDt() != null ? kikan.getSyuryoDt().toLocalDate() : null);
                tab.setTmMitsumoriNo(kikan.getTmMitsumoriNo());
                tab.setKaitoDt(kikan.getKaitoDt() != null ? kikan.getKaitoDt().toLocalDate() : null);
                tab.setHyojungokeiKin(kikan.getHyojungokeiKin());
                tab.setSikirisyokeiKin(kikan.getSikirisyokeiKin());
                tab.setSyusseinebikiKin(kikan.getSyusseinebikiKin());
                tab.setSikirigokeiKin(kikan.getSikirigokeiKin());
                tab.setBiko(kikan.getBiko());
                // TODO: TANKA, TENKENを読み込みtab.tankaRows/tenkenRowsにセット
                tabs.add(tab);
            }
        }
        form.setTabs(tabs);
        return form;
    }

    // ========== 編集可否チェック ==========

    /** 【変換元】checkUmMitsumori - ユーザ見積存在チェック */
    public boolean checkUmMitsumori(BigDecimal tmKeiyakujikanId) {
        // TODO: nativeQuery実装
        return false;
    }

    /** 【変換元】checkTkKiyaku - 取引先契約存在チェック */
    public boolean checkTkKiyaku(BigDecimal tmKeiyakujikanId) {
        // TODO: nativeQuery実装
        return false;
    }

    /** 【変換元】checkPlanId - プラント付替チェック */
    public boolean checkPlantId(BigDecimal tmKeiyakujikanId, BigDecimal plantId) {
        // TODO: nativeQuery実装
        return true;
    }

    private boolean determineViewFlag(String jotai, BigDecimal tmKeiyakujikanId,
                                      BigDecimal plantId, boolean lockRelease) {
        if (lockRelease) return true;
        if (!AppConstants.JOTAI_IRAI.equals(jotai)
            && !AppConstants.JOTAI_MITSUMORI.equals(jotai)
            && !AppConstants.JOTAI_SAKUSEICHU.equals(jotai)) return false;
        if (checkUmMitsumori(tmKeiyakujikanId)) return false;
        if (checkTkKiyaku(tmKeiyakujikanId)) return false;
        if (!checkPlantId(tmKeiyakujikanId, plantId)) return false;
        return true;
    }

    // ========== バリデーション ==========

    /** 【変換元】checkKikan - 期間バリデーション */
    public List<String> validatePeriods(List<Mcm1004uTabForm> tabs) {
        List<String> errors = new ArrayList<>();
        if (tabs == null || tabs.isEmpty()) return errors;
        for (int i = 0; i < tabs.size(); i++) {
            Mcm1004uTabForm t = tabs.get(i);
            if (t.getKaisiDt() == null) { errors.add("期間の開始日は必須です"); return errors; }
            if (i < tabs.size() - 1 && t.getSyuryoDt() == null) {
                errors.add(t.getTabLabel() + "タブの終了日は必須項目です。"); return errors;
            }
            if (t.getSyuryoDt() != null && t.getKaisiDt().isAfter(t.getSyuryoDt())) {
                errors.add("開始日は終了日以前に設定してください"); return errors;
            }
        }
        return errors;
    }

    public List<String> validateForTabAdd(Mcm1004uForm form) {
        List<String> errors = new ArrayList<>();
        if (form.getTabs() == null || form.getTabs().isEmpty()) return errors;
        Mcm1004uTabForm last = form.getTabs().get(form.getTabs().size() - 1);
        if (last.getKaisiDt() == null) errors.add("開始日は必須です");
        if (last.getSyuryoDt() == null) errors.add("タブを追加するには期間の終了日の入力が必要です。");
        return errors;
    }

    // ========== タブ追加 ==========

    /** 【変換元】TabTsuikaButton_Click タブ追加ロジック */
    public void addNewTab(Mcm1004uForm form) {
        Mcm1004uTabForm lastTab = form.getTabs().get(form.getTabs().size() - 1);
        Mcm1004uTabForm newTab = new Mcm1004uTabForm();
        newTab.setTabLabel((form.getTabs().size() + 1) + "回");
        newTab.setTmKikanId(Long.valueOf(0));
        // 変換元: syuryoDtDate.AddDays(1)
        if (lastTab.getSyuryoDt() != null) {
            newTab.setKaisiDt(lastTab.getSyuryoDt().plusDays(1));
        }
        // TODO: 前タブの単価/点検データコピー
        form.getTabs().add(newTab);
    }

    // ========== 登録処理 ==========

    /** 【変換元】UpdateButton_Click + UpdateButtonTabNaiyo */
    @Transactional
    public void updateAll(Mcm1004uForm form, String loginUser) {
        LocalDateTime now = LocalDateTime.now();
        // 変換元: 状態更新
        if (AppConstants.JOTAI_IRAI.equals(form.getJotai())
            || AppConstants.JOTAI_SAKUSEICHU.equals(form.getJotai())) {
            keiyakuRepo.findByTmKeiyakujikanId(form.getTmKeiyakujikanId()).ifPresent(k -> {
                k.setJotai(AppConstants.JOTAI_MITSUMORI);
                k.setLastupdateDt(now); k.setLastupdateBy(loginUser);
                keiyakuRepo.save(k);
            });
        }
        // 変換元: タブ毎更新 (UpdateButtonTabNaiyo)
        for (Mcm1004uTabForm tab : form.getTabs()) {
            if (tab.getTmKikanId().compareTo(Long.valueOf(0)) == 0) {
                BigDecimal newId = kikanRepo.findMaxId().orElse(BigDecimal.ZERO).add(BigDecimal.ONE);
                // TODO: KIKAN/TANKA/TENKEN INSERT with newId
            } else {
                // TODO: KIKAN/TANKA/TENKEN UPDATE
            }
        }
    }
}
