package com.daifuku.mcm.service;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.dto.UserInfo;
import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.form.Mcm1011uForm;
import com.daifuku.mcm.form.Mcm1011uForm.KeiyakuRowForm;
import com.daifuku.mcm.repository.Mcm1011uRepository;

/**
 * 【変換元】Mcm1011uScreen.vb
 * MCM1011U 審査・承認（取引先契約）サービス
 */
@Service
public class Mcm1011uService {

    @Autowired
    private Mcm1011uRepository repository;

    @Autowired
    private Mcm2004uService permissions;

    /**
     * 【移植】元VB: Mcm1011uScreen.SYONINButton/SASHIMODOSHIButton.AuthorityIsThrough=False
     *   「取引先契約関連」（機能ID=MCM1011U）の実効権限が作成（"2"）の場合のみ
     *   承認・差戻しを許可する。参照専用ユーザーは画面上のボタン非活性に加え、
     *   サーバー側でも必ず検証する。
     */
    public boolean canUpdate(String user) {
        return "2".equals(permissions.getAuthority(user, "MCM1011U"));
    }

    /** MCM1005U attachment service requires read (1) or update (2) authority. */
    public boolean canReadAttachment(String user) {
        return user != null && java.util.Set.of("1", "2").contains(
            String.valueOf(permissions.getAuthority(user, "MCM1005U")));
    }

    @Autowired private Mcm1011uNatsuinService stamps;
    @Autowired private Mcm1011uMailService mail;

    // ===================================================================
    // 申請一覧検索
    // 【変換元】Mcm1011uScreen.vb KensakuButton_Click → Search()
    // ===================================================================

    @Transactional(readOnly = true)
    public List<KeiyakuRowForm> search(Mcm1011uForm form, UserInfo userInfo) {
        List<String> shoninjotaiList = buildShoninjotaiList(form);
        if (shoninjotaiList.isEmpty()) return List.of();
        List<KeiyakuRowForm> rows = repository.search(
            form.getNonyusakiCd(), form.getNonyusakiNk(),
            form.getSupportId(), form.getPlantNk(),
            shoninjotaiList);

        var periods = new java.util.HashMap<BigDecimal, List<Mcm1011uForm.TenpuRowForm>>();
        for (KeiyakuRowForm row : rows) {
            var files = periods.computeIfAbsent(row.getTkKikanId(), repository::findTenpuByKikanId);
            row.setTenpuList(files.stream().filter(t -> java.util.Objects.equals(row.getHistoryKey(), t.getHistoryKey())
                    && shoninjotaiList.contains(t.getShoninjotai())).toList());
            row.setPendingAttachments(files.stream().filter(t -> !"3".equals(t.getShoninjotai())).toList());
        }
        return rows;
    }

    // ===================================================================
    // 承認処理
    // 【変換元】Mcm1011uScreen.vb ShoninButton_Click
    //   審査中(1) → 承認中(2)
    //   承認中(2) → 承認済(3) + MCM_TK_KIKAN.YUKO_FLG=有効
    //
    //   金額チェック:
    //   - 審査中行: SIKIRIGOKEI_KIN > KeiyakuShinsaKin → エラー
    //   - 承認中行: SIKIRIGOKEI_KIN > KeiyakuShoninKin → エラー
    // ===================================================================

    @Transactional(rollbackFor = Exception.class)
    public List<String> approve(List<KeiyakuRowForm> checkedRows, UserInfo userInfo, String loginUser) {
        var rows = validateCurrent(checkedRows, userInfo, loginUser, true);
        for (var row : rows) if (!java.util.Objects.equals(row.getContractState(), row.getShoninjotai()))
            throw new McmBusinessException("契約と添付資料の承認状態が一致しません。再検索し、申請内容を確認してください。");
        String actor = userInfo.getUserName();
        if (actor == null || actor.isBlank()) throw new McmBusinessException("担当者名が未設定です。担当者マスタを確認してください。");
        var notices = mail.prepare(rows, loginUser);
        stamps.stamp(rows, actor);
        var changedContracts = new java.util.HashSet<BigDecimal>();
        for (var row : rows) {
            if (changedContracts.add(row.getTkKeiyakuId().stripTrailingZeros()))
                repository.advanceContract(row.getTkKeiyakuId(), row.getContractState(), loginUser, actor);
            repository.advancePeriod(row.getTkKeiyakuId(), row.getTkKikanId(), row.getContractState(), loginUser, actor);
        }
        var warnings = new ArrayList<String>();
        mail.afterCommit(notices, warnings);
        return warnings;
    }

    @Transactional(rollbackFor = Exception.class)
    public void reject(List<KeiyakuRowForm> checkedRows, UserInfo userInfo, String loginUser) {
        var rows = validateCurrent(checkedRows, userInfo, loginUser, false);
        var changedContracts = new java.util.HashSet<BigDecimal>();
        for (var row : rows) {
            if (changedContracts.add(row.getTkKeiyakuId().stripTrailingZeros()))
                repository.rejectContract(row.getTkKeiyakuId(), loginUser);
            repository.rejectPeriod(row.getTkKeiyakuId(), row.getTkKikanId(), loginUser);
        }
    }

    private List<KeiyakuRowForm> validateCurrent(List<KeiyakuRowForm> requested, UserInfo userInfo, String user, boolean approving) {
        if (!canUpdate(user) || userInfo == null) throw new McmBusinessException("権限がないため実行できません。");
        if (requested == null || requested.isEmpty()) throw new McmBusinessException("対象の申請を選択してください。");
        if (requested.stream().anyMatch(r -> r == null || r.getTkKeiyakuId() == null || r.getTkKikanId() == null))
            throw new McmBusinessException("選択した申請が無効です。再検索してください。");
        var rows = new ArrayList<KeiyakuRowForm>();
        var seen = new java.util.HashSet<String>();
        for (var previous : requested.stream().sorted(java.util.Comparator.comparing(KeiyakuRowForm::getTkKeiyakuId).thenComparing(KeiyakuRowForm::getTkKikanId)).toList()) {
            if (!seen.add(previous.getRowKey())) throw new McmBusinessException("申請が重複しています。再検索してください。");
            var current = repository.lockCurrent(previous.getTkKeiyakuId(), previous.getTkKikanId());
            if (current == null || !"PENDING".equals(previous.getHistoryKey())
                    || !java.util.Objects.equals(current.getContractState(), previous.getContractState())
                    || !java.util.Objects.equals(current.getPeriodVersion(), previous.getPeriodVersion())
                    || !sameAttachments(current.getPendingAttachments(), previous.getPendingAttachments())
                    || !java.util.Set.of("1", "2").contains(String.valueOf(current.getContractState()))
                    || current.getTkKikanId() == null || previous.getTkKikanId() == null
                    || current.getTkKikanId().compareTo(previous.getTkKikanId()) != 0
                    || !java.util.Objects.equals(current.getShoninjotai(), previous.getShoninjotai())
                    || !java.util.Objects.equals(current.getJotai(), previous.getJotai())
                    || !java.util.Objects.equals(current.getLastupdateDt(), previous.getLastupdateDt())
                    || current.getKingaku() == null || previous.getKingaku() == null
                    || current.getKingaku().compareTo(previous.getKingaku()) != 0
                    || !("1".equals(current.getShoninjotai()) || "2".equals(current.getShoninjotai())))
                throw new McmBusinessException("申請が更新されたか、処理できない状態です。再検索してください。");
            boolean shinsa = "1".equals(current.getShoninjotai());
            if (!(shinsa ? userInfo.isShinsa() : userInfo.isShonin()))
                throw new McmBusinessException("担当外の承認状態です。再検索してください。");
            Integer limit = shinsa ? userInfo.getKeiyakuShinsaKin() : userInfo.getKeiyakuShoninKin();
            if (limit == null || current.getKingaku().compareTo(BigDecimal.valueOf(limit)) > 0)
                throw new McmBusinessException("審査・承認限度額を確認してください。");
            // VBは承認時のみ破棄・解約行のチェックを外す。差戻しの状態条件は維持する。
            if (approving && ("3".equals(current.getJotai()) || "4".equals(current.getJotai())))
                throw new McmBusinessException("破棄・解約済の契約は承認できません。再検索してください。");
            rows.add(current);
        }
        return rows;
    }

    // ===================================================================
    // private: 承認状態フィルターリスト構築
    // ===================================================================

    private boolean sameAttachments(List<Mcm1011uForm.TenpuRowForm> current, List<Mcm1011uForm.TenpuRowForm> previous) {
        if (previous == null || current.size() != previous.size()) return false;
        for (int i = 0; i < current.size(); i++) if (!current.get(i).sameVersion(previous.get(i))) return false;
        return true;
    }

    private List<String> buildShoninjotaiList(Mcm1011uForm form) {
        List<String> list = new ArrayList<>();
        if (form.isShinsachu()) list.add("1");
        if (form.isShoninchu()) list.add("2");
        if (form.isShoninzumi()) list.add("3");
        if (form.isSashimodoshi()) list.add("4");
        return list;
    }
}
