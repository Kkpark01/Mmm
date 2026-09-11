package com.daifuku.mcm.service;

import java.math.BigDecimal;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.constants.Mcm2003uConstants;
import com.daifuku.mcm.form.Mcm2003uForm;
import com.daifuku.mcm.repository.Mcm2003uRepository;

/**
 * 【変換元】Mcm2003uScreen.vb（MitsumoriButton_Click / HakkoButton_Click / SinseiButton_Click）
 * MCM2003U 店舗見積内容基本設定サービス
 *
 * Web版省略事項:
 *   - Excel帳票出力（setChohyo）
 *   - メール送信（申請時）
 *   - SP_UM ストアドプロシージャ呼び出し
 *   - ブランドタブ UpdateButtonShusei（MCM2005Uで別途対応）
 */
@Service
public class Mcm2003uService {

    @Autowired
    private Mcm2003uRepository repo;

    // ===================================================================
    // ロード
    // 【変換元】Mcm2003uScreen.vb — SearchRead()
    // ===================================================================

    public Mcm2003uForm load(BigDecimal umKihonMitsumoriId) {
        Mcm2003uForm form = repo.findKihonMitsumori(umKihonMitsumoriId);
        form.setMitsumoriRows(repo.findMitsumoriRows(umKihonMitsumoriId));
        form.setTenpuRows(repo.findTenpuRows(umKihonMitsumoriId));
        form.setBrandRows(repo.findBrandRows(umKihonMitsumoriId));
        form.setKoseiRows(repo.findKoseiRows(umKihonMitsumoriId));
        form.setMeisaiRows(repo.findMeisaiRows(umKihonMitsumoriId));
        return form;
    }

    // ===================================================================
    // 見積登録（保存）
    // 【変換元】MitsumoriButton_Click → UpdateButton(BUTTON_FLG_SAKUSEI) → SenteiNashiUpdate
    //   Web版: MCM_UM_KIHON_MITSUMORI のヘッダー項目のみ更新
    //          ブランドタブ更新(UpdateButtonShusei)は省略（TODO）
    // ===================================================================

    @Transactional
    public void save(Mcm2003uForm form, String loginUser) {
        requireEditable(form);
        repo.updateKihonMitsumori(form, loginUser);
    }

    // ===================================================================
    // 見積発行
    // 【変換元】HakkoButton_Click → UpdateButton(BUTTON_FLG_HAKKO) → SenteiNashiUpdate
    //   状態が作成中(9)の場合のみ見積(1)に変更。
    //   Excel帳票出力(setChohyo)は省略。
    // ===================================================================

    @Transactional
    public void publish(Mcm2003uForm form, String loginUser) {
        requireEditable(form);
        repo.updateKihonMitsumori(form, loginUser);
        if (McmConstants.JOTAI_SAKUSEICHU.equals(form.getJotai())) {
            repo.updateJotai(form.getUmKihonMitsumoriId(), McmConstants.JOTAI_MITSUMORI, loginUser);
        }
    }

    // ===================================================================
    // 申請
    // 【変換元】SinseiButton_Click
    //   添付ファイル必須チェック → SYOUNIN_JOTAI を審査中(1)へ変更
    //   メール送信は省略。
    // ===================================================================

    @Transactional
    public void sinsei(BigDecimal umKihonMitsumoriId, String loginUser) {
        requireEditable(repo.findKihonMitsumori(umKihonMitsumoriId));
        if (repo.countTenpu(umKihonMitsumoriId) == 0) {
            throw new IllegalStateException("見積資料を添付してください。");
        }
        repo.updateSyouninJotai(umKihonMitsumoriId, McmConstants.SHONINJOTAI_SHINSACHU_CD, loginUser);
    }

    /** 入力の保存と申請を同じトランザクションで行い、未保存内容の欠落を防ぐ。 */
    @Transactional
    public void saveAndApply(Mcm2003uForm form, String user) {
        requireEditable(form);
        if (repo.countTenpu(form.getUmKihonMitsumoriId()) == 0) throw new IllegalStateException("見積資料を添付してください。");
        repo.updateKihonMitsumori(form,user);
        repo.updateSyouninJotai(form.getUmKihonMitsumoriId(),McmConstants.SHONINJOTAI_SHINSACHU_CD,user);
    }
    public void requireEditable(Mcm2003uForm form) {
        if (form.getUmKihonMitsumoriId()==null || form.getUmKihonMitsumoriId().signum()<=0) throw new IllegalStateException("見積を選択してください。");
        repo.lock(form.getUmKihonMitsumoriId());
        var current=repo.findKihonMitsumori(form.getUmKihonMitsumoriId());
        if (isReadOnly(current.getSyouninJotai(),form.getSeniMotoKbn()) || "3".equals(current.getJotai()) || "4".equals(current.getJotai()))
            throw new IllegalStateException("見積の状態が変更されています。再読み込みしてください。");
        if (!java.util.Objects.equals(current.getLastupdateDt(),form.getLastupdateDt())) throw new IllegalStateException("他のユーザによって更新されています。再読み込みしてください。");
        if (form.getMitsumoriLevel()!=null && !form.getMitsumoriLevel().isBlank() && !form.getMitsumoriLevel().matches("[0-9]{1,2}"))
            throw new IllegalStateException("見積レベルは数値で入力してください。");
    }

    // ===================================================================
    // ReadOnly判定
    // 【変換元】View() — 承認状態が審査中/承認中/承認済、または遷移元区分=4の場合ReadOnly
    // ===================================================================

    public boolean isReadOnly(String syouninJotai, int seniMotoKbn) {
        if (seniMotoKbn == Mcm2003uConstants.SENIMOTO_KBN_SHONIN) {
            return true;
        }
        return McmConstants.SHONINJOTAI_SHINSACHU_CD.equals(syouninJotai)
            || McmConstants.SHONINJOTAI_SHONINCHU_CD.equals(syouninJotai)
            || McmConstants.SHONINJOTAI_SHONINZUMI_CD.equals(syouninJotai);
    }
}
