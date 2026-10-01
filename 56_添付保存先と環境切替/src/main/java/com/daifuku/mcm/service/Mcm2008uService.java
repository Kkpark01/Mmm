package com.daifuku.mcm.service;

import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.dto.UserInfo;
import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.form.Mcm2008uForm;
import com.daifuku.mcm.form.Mcm2008uForm.MitsumoriRowForm;
import com.daifuku.mcm.repository.Mcm2008uRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * 【変換元】Mcm2008uScreen.vb
 * MCM2008U 審査・承認（店舗見積）サービス
 *
 * 承認時の添付Excel捺印（UpdateExcelReport）は Mcm2008uNatsuinService で行う。
 *
 * Web版省略:
 *   - メール送信（CPSendMailClass）
 */
@Service
public class Mcm2008uService {

    @Autowired
    private Mcm2008uRepository repository;

    @Autowired
    private Mcm2008uNatsuinService natsuinService;

    @Autowired
    private Mcm2004uService permissions;

    /** 添付ファイルの保存ルート。MCM2003U と同一の保管場所を参照する。 */
    @Autowired private com.daifuku.mcm.common.AttachmentStorage storage;

    /**
     * 【移植】元VB: CPCoreUserControl.GetAuthorityDivision() および
     * Mcm2008uScreen.SYONINButton/SASHIMODOSHIButton.AuthorityIsThrough=False 相当。
     * 「店舗見積関連」（機能ID=MCM2008U）の実効権限が作成（"2"）の場合のみ、
     * 承認・差戻しを許可する。参照専用ユーザーは画面上のボタン非活性に加え、
     * サーバー側でも必ず検証する。
     */
    public boolean canUpdate(String user) {
        return "2".equals(permissions.getAuthority(user, "MCM2008U"));
    }

    /**
     * 添付ファイルの物理パスを解決する。
     * 【移植】MCM2003U の添付資料ダウンロードと同一の保管場所・同一の検証ロジックを使用する。
     *   保存ルート配下のファイルであること、通常ファイルであることを検証し、
     *   ディレクトリトラバーサルを防止する。
     *
     * @param umKihonMitsumoriId 店舗基本見積ID（対象行の特定に使用）
     * @param attachmentId       添付ID（MCM_UM_TENPU.UM_TENPU_ID）
     * @return 添付ファイルの物理パス
     * @throws java.io.IOException 保存先が不正な場合、またはファイルが存在しない場合
     */
    @Transactional(readOnly = true)
    public java.nio.file.Path resolveAttachment(BigDecimal umKihonMitsumoriId, BigDecimal attachmentId)
            throws java.io.IOException {
        Mcm2008uForm.TenpuRowForm target = repository.findTenpuByKihonMitsumoriId(umKihonMitsumoriId).stream()
            .filter(r -> r.getUmTenpuId() != null && r.getUmTenpuId().compareTo(attachmentId) == 0)
            .findFirst()
            .orElseThrow(() -> new IllegalStateException("添付資料が見つかりません。"));

        return storage.resolve(target.getDirectory(), target.getTenpufileNk());
    }

    /**
     * 添付ファイル行を取得する（ダウンロード時のファイル名決定用）。
     *
     * @param umKihonMitsumoriId 店舗基本見積ID
     * @param attachmentId       添付ID
     * @return 添付ファイル行
     */
    @Transactional(readOnly = true)
    public Mcm2008uForm.TenpuRowForm findAttachmentRow(BigDecimal umKihonMitsumoriId, BigDecimal attachmentId) {
        return repository.findTenpuByKihonMitsumoriId(umKihonMitsumoriId).stream()
            .filter(r -> r.getUmTenpuId() != null && r.getUmTenpuId().compareTo(attachmentId) == 0)
            .findFirst()
            .orElseThrow(() -> new IllegalStateException("添付資料が見つかりません。"));
    }

    // ===================================================================
    // 申請一覧検索
    // 【変換元】Mcm2008uScreen.vb SearchButton_Click → Search()
    // ===================================================================

    @Transactional(readOnly = true)
    public List<MitsumoriRowForm> search(Mcm2008uForm form) {
        List<String> syouninJotaiList = buildSyouninJotaiList(form);
        List<MitsumoriRowForm> rows = repository.search(
            form.getNonyusakiCd(), form.getNonyusakiNk(),
            form.getSupportId(), form.getPlantNk(),
            syouninJotaiList);

        for (MitsumoriRowForm row : rows) {
            if (row.getUmKihonMitsumoriId() != null) {
                row.setTenpuList(repository.findTenpuByKihonMitsumoriId(row.getUmKihonMitsumoriId()));
                row.setTvaList(repository.findTvaByKihonMitsumoriId(row.getUmKihonMitsumoriId()));
            }
        }
        return rows;
    }

    /**
     * 行ごとのチェックボックス選択可否を算出して設定する。
     * 【変換元】Mcm2008uScreen.vb Search() の CHECK_FLG_MITSUMORI_TextBox.ReadOnly 制御。
     *   ・審査中(1): 金額 &gt; 審査限度額（MitsumoriShinsaKin） → 選択不可
     *   ・承認中(2): 金額 &gt; 承認限度額（MitsumoriShoninKin） → 選択不可
     *   ・上記以外（承認済(3)/差戻し中(4)/審査中・承認中で限度額内） → VB版どおり
     *     ※承認済・差戻し中はそもそも承認/差戻し対象外のため選択不可とする。
     * サーバー側で確定した結果を Form に保持し、画面の disabled 制御に用いる。
     *
     * @param rows     検索結果行
     * @param userInfo ログインユーザー情報（限度額の取得元）。null の場合は限度額チェックを行わない。
     */
    public void applySelectable(List<MitsumoriRowForm> rows, UserInfo userInfo) {
        if (rows == null) {
            return;
        }
        Integer shinsaKin = (userInfo != null) ? userInfo.getMitsumoriShinsaKin() : null;
        Integer shoninKin = (userInfo != null) ? userInfo.getMitsumoriShoninKin() : null;

        for (MitsumoriRowForm row : rows) {
            String syouninJotai = row.getSyouninJotai();
            BigDecimal kingaku = row.getKingaku();

            boolean selectable;
            if (McmConstants.SHONINJOTAI_SHINSACHU_CD.equals(syouninJotai)) {
                // 審査中: 金額が審査限度額を超える場合は選択不可
                selectable = !(shinsaKin != null && kingaku != null
                    && kingaku.compareTo(BigDecimal.valueOf(shinsaKin)) > 0);
            } else if (McmConstants.SHONINJOTAI_SHONINCHU_CD.equals(syouninJotai)) {
                // 承認中: 金額が承認限度額を超える場合は選択不可（VB版準拠）
                selectable = !(shoninKin != null && kingaku != null
                    && kingaku.compareTo(BigDecimal.valueOf(shoninKin)) > 0);
            } else {
                // 承認済(3)/差戻し中(4) など、審査中・承認中以外は選択不可
                selectable = false;
            }
            row.setSelectable(selectable);
        }
    }

    // ===================================================================
    // 承認処理
    // 【変換元】Mcm2008uScreen.vb SYONINButton_Click
    //   審査中(1) → 承認中(2)
    //   承認中(2) → 承認済(3)
    //
    //   金額チェック:
    //   - 審査中行: KINGAKU > mitsumoriShinsaKin → エラー
    //   - 承認中行: KINGAKU > mitsumoriShoninKin → エラー
    // ===================================================================

    @Transactional(rollbackFor = Exception.class)
    public int approve(List<MitsumoriRowForm> checkedRows, UserInfo userInfo, String loginUser) {
        Integer shinsaKin = (userInfo != null) ? userInfo.getMitsumoriShinsaKin() : null;
        Integer shoninKin = (userInfo != null) ? userInfo.getMitsumoriShoninKin() : null;
        // 捺印の名前欄に使う氏名（MCM_MO_TANTO.TANTO_NK）。姓の抽出は捺印サービス側で行う。
        String userName = (userInfo != null) ? userInfo.getUserName() : null;

        // 実際に承認処理を行った件数。
        //   破棄・解約の除外後、審査中／承認中の有効な行が0件なら呼び出し元で MSG_0088G を表示する
        //   （VB版 SYONINButton_Click: 破棄・解約のチェックを外した後の件数チェック相当）。
        int processed = 0;

        for (MitsumoriRowForm row : checkedRows) {
            // 破棄・解約は承認対象外（VB版 SYONINButton_Click 冒頭「破棄・解約のチェックをはずす」相当）。
            //   破棄・解約の行はチェック可能だが、承認実行時に処理対象から除外する。
            //   不正リクエストでチェックが付いていてもサーバー側で確実に除外する。
            //   JOTAI の前後空白はトリムして判定する（表示側は #strings.trim で吸収済みのため揃える）。
            String jotai = (row.getJotai() != null) ? row.getJotai().trim() : null;
            if (McmConstants.JOTAI_HAKI.equals(jotai) || McmConstants.JOTAI_KAIYAKU.equals(jotai)) {
                continue;
            }

            String syouninJotai = row.getSyouninJotai();
            BigDecimal kingaku = row.getKingaku();

            if (McmConstants.SHONINJOTAI_SHINSACHU_CD.equals(syouninJotai)) {
                if (shinsaKin != null && kingaku != null
                        && kingaku.compareTo(BigDecimal.valueOf(shinsaKin)) > 0) {
                    throw new McmBusinessException(
                        "書類NO[" + row.getShoruiNo() + "]の金額（"
                        + kingaku.toPlainString() + "円）が審査限度額（"
                        + shinsaKin + "円）を超えています。");
                }
                // 【変換元】SYONINButton_Click: 捺印押下（審査印）→ 成功後に承認状態を更新
                //   SHINSA_BY は VB版どおり担当者名（UserName）、LASTUPDATE_BY はログインID
                natsuinService.stamp(row.getUmKihonMitsumoriId(), syouninJotai, userName);
                repository.approveShinsachu(row.getUmKihonMitsumoriId(), userName, loginUser);
                processed++;

            } else if (McmConstants.SHONINJOTAI_SHONINCHU_CD.equals(syouninJotai)) {
                if (shoninKin != null && kingaku != null
                        && kingaku.compareTo(BigDecimal.valueOf(shoninKin)) > 0) {
                    throw new McmBusinessException(
                        "書類NO[" + row.getShoruiNo() + "]の金額（"
                        + kingaku.toPlainString() + "円）が承認限度額（"
                        + shoninKin + "円）を超えています。");
                }
                // 【変換元】SYONINButton_Click: 捺印押下（承認印）→ 成功後に承認状態を更新
                //   SYOUNIN_BY は VB版どおり担当者名（UserName）、LASTUPDATE_BY はログインID
                natsuinService.stamp(row.getUmKihonMitsumoriId(), syouninJotai, userName);
                repository.approveShoninchu(row.getUmKihonMitsumoriId(), userName, loginUser);
                processed++;
            }
        }
        return processed;
    }

    // ===================================================================
    // 差戻し処理
    // 【変換元】Mcm2008uScreen.vb SASHIMODOSHIButton_Click
    //   審査中(1) / 承認中(2) のみ差戻し可
    // ===================================================================

    @Transactional(rollbackFor = Exception.class)
    public void reject(List<MitsumoriRowForm> checkedRows, String loginUser) {
        for (MitsumoriRowForm row : checkedRows) {
            String syouninJotai = row.getSyouninJotai();
            if (McmConstants.SHONINJOTAI_SHINSACHU_CD.equals(syouninJotai) ||
                McmConstants.SHONINJOTAI_SHONINCHU_CD.equals(syouninJotai)) {
                repository.sashimodoshi(row.getUmKihonMitsumoriId(), loginUser);
            }
        }
    }

    // ===================================================================
    // private
    // ===================================================================

    private List<String> buildSyouninJotaiList(Mcm2008uForm form) {
        List<String> list = new ArrayList<>();
        if (form.isShinsachu())    list.add(McmConstants.SHONINJOTAI_SHINSACHU_CD);
        if (form.isShoninchu())    list.add(McmConstants.SHONINJOTAI_SHONINCHU_CD);
        if (form.isShoninzumi())   list.add(McmConstants.SHONINJOTAI_SHONINZUMI_CD);
        if (form.isSashimodoshi()) list.add(McmConstants.SHONINJOTAI_SASHIMODOSHI_CD);
        return list;
    }
}
