package com.daifuku.mcm.service;

import com.daifuku.mcm.dto.Mcm2004uDeliveryDto;
import com.daifuku.mcm.dto.Mcm2004uRowDto;
import com.daifuku.mcm.form.Mcm2004uForm;
import com.daifuku.mcm.repository.Mcm2004uJdbcRepository;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.List;

/**
 * 【変換元】Mcm2004uScreen.vb（26,158 bytes / 6メソッド）
 * MCM2004U（カスタマー見積・契約一覧）Service
 *
 * <p>検索（動的11条件）、見積破棄、契約破棄、行削除、権限制御を担当。</p>
 *
 * @author MCM Migration Tool
 */
@Service
public class Mcm2004uService {

    private static final Logger log = LoggerFactory.getLogger(Mcm2004uService.class);

    private final Mcm2004uJdbcRepository repo;

    private static final DateTimeFormatter FMT_YYYYMMDD =
        DateTimeFormatter.ofPattern("uuuu/MM/dd").withResolverStyle(java.time.format.ResolverStyle.STRICT);

    // VB版 McmConstant.vb の状態コード。
    private static final String JOTAI_MITSUMORI_CD = "1";
    private static final String JOTAI_KEIYAKU_CD   = "2";
    private static final String JOTAI_KAIYAKU_CD   = "4";
    private static final String JOTAI_HAKI_CD      = "3";

    public Mcm2004uService(Mcm2004uJdbcRepository repo) {
        this.repo = repo;
    }

    public String getAuthority(String loginId,String functionId){return repo.findAuthority(loginId,functionId);}
    public boolean isMaintenance(String loginId){return repo.findMaintenance(loginId);}

    // ================================================================
    // 内部クラス: 検索結果 + 権限フラグ
    // ================================================================

    /**
     * 画面ロード／検索の結果を保持する。
     */
    public static class SearchResult {
        private List<Mcm2004uRowDto> rows = new ArrayList<>();
        /** 見積複製リンク表示 */
        private boolean showMitsumoriCopy;
        /** 見積破棄リンク表示 */
        private boolean showMitsumoriDel;
        /** 契約作成リンク表示 */
        private boolean showKeiyakuKeiyaku;
        /** 契約破棄リンク表示 */
        private boolean showKeiyakuDel;
        /** 新規ボタン有効 */
        private boolean newButtonEnabled;
        /** 行削除ボタン表示 */
        private boolean deleteButtonVisible;
        /** エラーメッセージ */
        private List<String> errors = new ArrayList<>();

        public List<Mcm2004uRowDto> getRows() { return rows; }
        public void setRows(List<Mcm2004uRowDto> v) { this.rows = v; }
        public boolean isShowMitsumoriCopy() { return showMitsumoriCopy; }
        public void setShowMitsumoriCopy(boolean v) { this.showMitsumoriCopy = v; }
        public boolean isShowMitsumoriDel() { return showMitsumoriDel; }
        public void setShowMitsumoriDel(boolean v) { this.showMitsumoriDel = v; }
        public boolean isShowKeiyakuKeiyaku() { return showKeiyakuKeiyaku; }
        public void setShowKeiyakuKeiyaku(boolean v) { this.showKeiyakuKeiyaku = v; }
        public boolean isShowKeiyakuDel() { return showKeiyakuDel; }
        public void setShowKeiyakuDel(boolean v) { this.showKeiyakuDel = v; }
        public boolean isNewButtonEnabled() { return newButtonEnabled; }
        public void setNewButtonEnabled(boolean v) { this.newButtonEnabled = v; }
        public boolean isDeleteButtonVisible() { return deleteButtonVisible; }
        public void setDeleteButtonVisible(boolean v) { this.deleteButtonVisible = v; }
        public List<String> getErrors() { return errors; }
        public void setErrors(List<String> v) { this.errors = v; }
    }

    // ================================================================
    // 1. 画面ロード
    // 【変換元】Mcm2004uScreen_Load()（56行）
    //   元コード:
    //     If IsNotNull(delivery) Then
    //       Search(umKihonMitsumoriId)
    //     End If
    //     '' 権限制御
    //     mitsumoriMenuAuthorityDivision = GetAuthorityDivision("MCM2003U")
    //     keiyakuMenuAuthorityDivision = GetAuthorityDivision("MCM2006U")
    // ================================================================
    @Transactional(readOnly = true)
    public SearchResult loadScreen(Mcm2004uDeliveryDto delivery,
                                    String mitsumoriAuthority,
                                    String keiyakuAuthority,
                                    boolean isMaintenance) {

        SearchResult result = new SearchResult();

        // --- 権限制御 ---
        /**
         * 【変換元】Mcm2004uScreen_Load() - 権限制御
         *   元コード:
         *     If mitsumoriMenuAuthorityDivision.Equals(AUTHORITY_DIVISION_UPDATE) Then
         *       MITSUMORI_COPY → Visible = True
         *       MITSUMORI_DEL → Visible = True
         *       NEWROWButton.Enabled = True
         *     Else ...
         */
        boolean hasMitsumoriUpdate = "UPDATE".equals(mitsumoriAuthority);
        result.setShowMitsumoriCopy(hasMitsumoriUpdate);
        result.setShowMitsumoriDel(hasMitsumoriUpdate);
        result.setNewButtonEnabled(hasMitsumoriUpdate);

        /**
         * 【変換元】Mcm2004uScreen_Load() - 契約権限制御
         *   元コード:
         *     If keiyakuMenuAuthorityDivision.Equals(AUTHORITY_DIVISION_UPDATE) Then
         *       KEIYAKU_KEIYAKU → Visible = True
         *       KEIYAKU_DEL → Visible = True
         */
        boolean hasKeiyakuUpdate = "UPDATE".equals(keiyakuAuthority);
        result.setShowKeiyakuKeiyaku(hasKeiyakuUpdate);
        result.setShowKeiyakuDel(hasKeiyakuUpdate);

        /**
         * 【変換元】Mcm2004uScreen_Load() - メンテナンス権限
         *   元コード:
         *     If CPSettingInfo.GetUserInfo.Maintenance = False Then
         *       RowDeleteButton.Visible = False
         */
        result.setDeleteButtonVisible(isMaintenance);

        // --- デリバリーからの初期検索 ---
        if (delivery != null && delivery.getUmKihonMitsumoriId() != null
            && delivery.getUmKihonMitsumoriId().compareTo(BigDecimal.ZERO) > 0) {
            /**
             * 【変換元】Mcm2004uScreen_Load()
             *   元コード: Search(umKihonMitsumoriId)
             */
            List<Mcm2004uRowDto> rows =
                repo.searchByUmKihonMitsumoriId(delivery.getUmKihonMitsumoriId());
            result.setRows(rows);
        }

        return result;
    }

    // ================================================================
    // 2. 検索
    // 【変換元】SEARCHButton_Click()（110行）
    //   元コード:
    //     (1) 全条件空白チェック → MSG_0001
    //     (2) 開始日付フォーマットチェック → MSG_0058
    //     (3) 終了日付フォーマットチェック → MSG_0059
    //     (4) 状態チェックボックス → jotaiCheck 文字列構築
    //     (5) 承認状態チェックボックス → syouninCheck 文字列構築
    //     (6) Fill(DataSet, params, XML)
    // ================================================================
    @Transactional(readOnly = true)
    public SearchResult search(Mcm2004uForm form) {

        SearchResult result = new SearchResult();

        // --- (1)(2)(3) バリデーション ---
        List<String> errors = validateSearch(form);
        if (!errors.isEmpty()) {
            result.setErrors(errors);
            return result;
        }

        // --- (6) 検索実行 ---
        List<Mcm2004uRowDto> rows = repo.searchByConditions(form);

        /*
         * 【変換元】SEARCHButton_Click()
         *   元コード:
         *     '' 検索結果が1件も存在しない場合
         *     If mcm2004VCount = 0 Then
         *       MyBase.DisplayMessage(CPMessageConstant.MSG_0002)
         *     End If
         */
        if (rows.isEmpty()) {
            List<String> noDataErrors = new ArrayList<>();
            noDataErrors.add("検索結果が1件も存在しません。");
            result.setErrors(noDataErrors);
        }

        result.setRows(rows);

        return result;
    }

    // ================================================================
    // 3. 検索バリデーション
    // 【変換元】SEARCHButton_Click() 先頭部分
    //   元コード:
    //     If IsNull(全テキスト) And 全CheckBox.Checked = False Then
    //       DisplayMessage(MSG_0001)  '' 「検索条件を入力してください」
    //     End If
    //     If IsNotNull(MITSUMORI_DT_START) And Not IsCheckYmd(...) Then
    //       DisplayMessage(MSG_0058)  '' 「開始日が不正です」
    //     End If
    // ================================================================
    public List<String> validateSearch(Mcm2004uForm form) {

        List<String> errors = new ArrayList<>();

        // --- 全条件空白チェック ---
        boolean allEmpty =
            isEmpty(form.getUmMitsumoriNo()) &&
            isEmpty(form.getKeiyakuNo()) &&
            isEmpty(form.getMitsumoriDtStart()) &&
            isEmpty(form.getMitsumoriDtEnd()) &&
            isEmpty(form.getSofutenpoId()) &&
            isEmpty(form.getNonyusakiCd()) &&
            isEmpty(form.getNonyusakiNk()) &&
            isEmpty(form.getSupportId()) &&
            isEmpty(form.getPlantNk()) &&
            !form.isJotaiMitsumori() &&
            !form.isJotaiKeiyaku() &&
            !form.isJotaiKaiyaku() &&
            !form.isJotaiHaki() &&
            !form.isSyouninJotaiSakuseichu() &&
            !form.isSyouninJotaiSinsachu() &&
            !form.isSyouninJotaiShoninchu() &&
            !form.isSyouninJotaiShoninzumi() &&
            !form.isSyouninJotaiSashimodoshichu();

        if (allEmpty) {
            errors.add("検索条件は、1項目以上選択して下さい。");
            return errors;
        }

        if (isNotEmpty(form.getSofutenpoId())) {
            try { if (new BigDecimal(form.getSofutenpoId()).signum() <= 0) throw new NumberFormatException(); }
            catch (NumberFormatException e) { return List.of("送付先事業所を選択し直してください。"); }
        }
        // --- 開始日付フォーマットチェック ---
        if (isNotEmpty(form.getMitsumoriDtStart())) {
            try {
                LocalDate.parse(form.getMitsumoriDtStart(), FMT_YYYYMMDD);
            } catch (DateTimeParseException e) {
                errors.add("開始日付の書式を指定して下さい。(YYYY/MM/DD)"); return errors;
            }
        }

        // --- 終了日付フォーマットチェック ---
        if (isNotEmpty(form.getMitsumoriDtEnd())) {
            try {
                LocalDate.parse(form.getMitsumoriDtEnd(), FMT_YYYYMMDD);
            } catch (DateTimeParseException e) {
                errors.add("終了日付の書式を指定して下さい。(YYYY/MM/DD)");
            }
        }

        return errors;
    }

    // ================================================================
    // 4. 見積破棄
    // 【変換元】CellContentClick() - 見積破棄リンク分岐（約60行）
    //   元コード:
    //     '' ユーザ契約IDが存在する場合はエラー
    //     If IsNotNull(ukKeiyakuId) And ukKeiyakuId <> 0 Then
    //       DisplayMessage(MSG_0025)  '' 「契約が存在するため破棄できません」
    //       Return
    //     End If
    //     '' ConfirmMessage → If Yes Then
    //     MCM_UM_KIHON_MITSUMORI(0).JOTAI = JOTAI_HAKI_CD
    //     DaoContainer.Update(DataSet, "MCM_UM_KIHON_MITSUMORI")
    //     Search(umKihonMitsumoriId)
    // ================================================================
    @Transactional(isolation=org.springframework.transaction.annotation.Isolation.SERIALIZABLE)
    public List<String> discardMitsumori(BigDecimal umKihonMitsumoriId,
                                          BigDecimal ukKeiyakuId,
                                          String loginUserId) {

        List<String> errors = new ArrayList<>();

        requireRow(umKihonMitsumoriId, ukKeiyakuId, "discard-mitsumori");
        // --- ユーザ契約存在チェック ---
        /**
         * 【変換元】CellContentClick() - 見積破棄 ユーザ契約チェック
         *   元コード: If IsNotNull(ukKeiyakuId) And ukKeiyakuId <> 0
         */
        if (ukKeiyakuId != null && ukKeiyakuId.compareTo(BigDecimal.ZERO) != 0) {
            int count = repo.countUkKeiyaku(ukKeiyakuId);
            if (count > 0) {
                errors.add("店舗への見積依頼 又は 取引先との契約として使用されている為、破棄する事が出来ません。");
                return errors;
            }
        }

        // --- JOTAI更新 ---
        requireOne(repo.updateMitsumoriJotai(umKihonMitsumoriId, JOTAI_HAKI_CD, loginUserId));

        return errors;
    }

    // ================================================================
    // 5. 契約破棄
    // 【変換元】CellContentClick() - 契約破棄リンク分岐（約50行）
    //   元コード:
    //     '' ConfirmMessage → If Yes Then
    //     '' 見積のJOTAIを「見積」に戻す
    //     MCM_UM_KIHON_MITSUMORI(0).JOTAI = JOTAI_MITSUMORI_CD
    //     DaoContainer.Update(DataSet, "MCM_UM_KIHON_MITSUMORI")
    //     '' 契約のJOTAIを「破棄」にする
    //     MCM_UK_KEIYAKU(0).JOTAI = JOTAI_HAKI_CD
    //     DaoContainer.Update(DataSet, "MCM_UK_KEIYAKU")
    //     Search(umKihonMitsumoriId)
    // ================================================================
    @Transactional(isolation=org.springframework.transaction.annotation.Isolation.SERIALIZABLE)
    public void discardKeiyaku(BigDecimal umKihonMitsumoriId,
                                BigDecimal ukKeiyakuId,
                                String loginUserId) {

        /**
         * 【変換元】CellContentClick() - 契約破棄
         *   (1) 見積の状態を「見積」に戻す
         *   (2) 契約の状態を「破棄」にする
         */
        requireRow(umKihonMitsumoriId, ukKeiyakuId, "discard-keiyaku");
        requireOne(repo.updateMitsumoriJotai(umKihonMitsumoriId, JOTAI_MITSUMORI_CD, loginUserId));
        requireOne(repo.updateUkKeiyakuJotai(ukKeiyakuId, JOTAI_HAKI_CD, loginUserId));
    }

    // ================================================================
    // 6. 行削除
    // 【変換元】RowDeleteButton_Click()（30行）
    //   VB版の選択行削除は McmDBUtility.DeleteUser → MCM_DEL_PAC.SP_UM。
    //   複数選択時は deleteRows() で全対象を一括してトランザクション処理する。
    // ================================================================
    @Transactional(isolation=org.springframework.transaction.annotation.Isolation.SERIALIZABLE)
    public void deleteRow(BigDecimal umKihonMitsumoriId, String loginUserId) {

        /**
         * 【変換元】RowDeleteButton_Click()
         *   元コード: McmDBUtility.DeleteUser(umKihonMitsumoriId)
         */
        requireRow(umKihonMitsumoriId, null, "delete-row");
        requireOne(repo.deleteUser(umKihonMitsumoriId, loginUserId));
    }

    // ================================================================
    // 7. 送付先事業所一覧取得（コンボボックス用）
    // 【変換元】McmSohumeisyo4NkDataTable + McmSohumeisyo4NkDataTable_Sql.xml
    //   元コード: SOFUMEISHO4_NKCombobox のデータソース
    //   DisplayMember = "VALUE" (MEISHO4_NK), ValueMember = "KEY" (SOFUTENPO_ID)
    // ================================================================
    @Transactional(readOnly = true)
    public List<java.util.Map<String, Object>> getSofutenpoList() {
        try {
            return repo.findDistinctSofutenpo();
        } catch (Exception e) {
            log.warn("MCM2004U 送付先事業所一覧取得エラー（テーブル未作成の可能性）: {}", e.getMessage());
            return new ArrayList<>();
        }
    }

    // ================================================================
    // ユーティリティ
    // ================================================================

    public static final String CHANGED = "他のユーザがデータを変更した可能性があります。処理をやり直してください。";
    public Mcm2004uRowDto requireRow(BigDecimal id, BigDecimal contract, String operation) {
        if (id == null || id.signum() <= 0) throw new IllegalStateException(CHANGED);
        return repo.searchByUmKihonMitsumoriId(id).stream()
            .filter(r -> same(r.getUkKeiyakuId(), contract) || "delete-row".equals(operation))
            .filter(r -> switch(operation) {
                case "discard-mitsumori" -> isNotEmpty(r.getMitsumoriDel());
                case "discard-keiyaku" -> isNotEmpty(r.getKeiyakuDel()) && contract != null && contract.signum() > 0;
                default -> true;
            }).findFirst().orElseThrow(() -> new IllegalStateException(CHANGED));
    }
    public static boolean same(BigDecimal a,BigDecimal b) { return (a==null?BigDecimal.ZERO:a).compareTo(b==null?BigDecimal.ZERO:b)==0; }
    private static void requireOne(int count) { if(count != 1) throw new IllegalStateException(CHANGED); }
    @Transactional(isolation=org.springframework.transaction.annotation.Isolation.SERIALIZABLE)
    public void deleteRows(List<BigDecimal> ids,String user) {
        for (var id:ids) requireRow(id,null,"delete-row");
        for (var id:ids) requireOne(repo.deleteUser(id,user));
    }
    private boolean isEmpty(String s) {
        return s == null || s.isBlank();
    }

    private boolean isNotEmpty(String s) {
        return s != null && !s.isBlank();
    }
}
