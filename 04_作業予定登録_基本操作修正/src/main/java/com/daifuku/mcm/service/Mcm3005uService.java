package com.daifuku.mcm.service;

import com.daifuku.mcm.form.Mcm3005uForm;
import com.daifuku.mcm.form.Mcm3005uForm.TenpuRowForm;
import com.daifuku.mcm.repository.Mcm3005uRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.math.BigDecimal;
import java.util.List;

/**
 * 【変換元】Mcm3005uScreen.vb
 * MCM3005U 作業予定登録サービス
 */
@Service
public class Mcm3005uService {

    @Autowired
    private Mcm3005uRepository repository;

    /**
     * 画面初期表示ロード。
     * 【変換元】Mcm3005uScreen_Load
     *   - torokuId == 0 → 新規フォーム（TOROKU_IDをMAX+1で発番）
     *   - torokuId > 0  → 既存データをロード
     */
    public Mcm3005uForm load(BigDecimal torokuId) {
        if (torokuId == null || torokuId.compareTo(BigDecimal.ZERO) == 0) {
            // 新規
            Mcm3005uForm form = new Mcm3005uForm();
            form.setTorokuId(repository.getNextTorokuId());
            form.setNewRecord(true);
            return form;
        }
        // 既存レコードをロード
        Mcm3005uForm form = repository.findByTorokuId(torokuId);
        if (form == null) {
            throw new IllegalStateException("検索結果が1件も存在しません。");
        }
        List<TenpuRowForm> tenpuList = repository.findTenpuByTorokuId(torokuId);
        form.setTenpuList(tenpuList);
        return form;
    }

    /**
     * 保存（INSERT or UPDATE）。
     * 【変換元】Mcm3005uScreen.vb UpdateButton_Click（deleteFlg=false 時）
     *
     * @param form      フォームデータ
     * @param loginUser ログインユーザ
     * @param isNew     true=INSERT、false=UPDATE
     */
    @Transactional
    public void save(Mcm3005uForm form, String loginUser, boolean isNew) {
        validate(form);
        if (isNew) {
            repository.insert(form, loginUser);
        } else {
            int cnt = repository.update(form, loginUser);
            if (cnt == 0) {
                throw new IllegalStateException("他のユーザがデータを変更した可能性があります。処理をやり直してください。");
            }
        }
    }

    /** 確認ダイアログの前と保存直前で同じ入力チェックを行う。 */
    public void validate(Mcm3005uForm form) {
        if (form.getYoteiDt() == null || form.getYoteiDt().isBlank()) {
            throw new IllegalArgumentException("予定日は必ず入力してください。");
        }
        try {
            java.time.LocalDate.parse(form.getYoteiDt(), java.time.format.DateTimeFormatter
                    .ofPattern("uuuu/MM/dd").withResolverStyle(java.time.format.ResolverStyle.STRICT));
        } catch (java.time.format.DateTimeParseException ex) {
            throw new IllegalArgumentException("予定日はyyyy/MM/ddの書式で入力してください。");
        }
        if (form.getKanryoFlg() != 0 && form.getKanryoFlg() != 1)
            throw new IllegalArgumentException("完了は0または1で入力してください。");
        if (form.getNaiyo() != null && form.getNaiyo().length() > 4000)
            throw new IllegalArgumentException("内容は4000桁以下で入力してください。");
        if (form.getBiko() != null && form.getBiko().length() > 4000)
            throw new IllegalArgumentException("備考は4000桁以下で入力してください。");
    }

    /**
     * 削除。
     * 【変換元】Mcm3005uScreen.vb UpdateButton_Click（deleteFlg=true 時）
     *   添付ファイルが存在する場合は MSG_0139 で中断。
     */
    @Transactional
    public void delete(BigDecimal torokuId) {
        validateDelete(torokuId);
        if (repository.delete(torokuId) == 0)
            throw new IllegalStateException("他のユーザがデータを変更した可能性があります。処理をやり直してください。");
    }

    public void validateDelete(BigDecimal torokuId) {
        if (torokuId == null || torokuId.compareTo(BigDecimal.ZERO) == 0) {
            throw new IllegalArgumentException("削除対象のレコードが不正です。");
        }
        // 【変換元】VB版: Count(MCM_MO_TENPU) > 0 → MSG_0139「添付ファイルが存在するため削除できません」
        int tenpuCount = repository.countTenpu(torokuId);
        if (tenpuCount > 0) {
            throw new IllegalStateException(
                "添付ファイルを削除してください。");
        }
    }
}
