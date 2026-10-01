package com.daifuku.mcm.service;

import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.dto.UserInfo;
import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.form.Mcm3001uForm;
import com.daifuku.mcm.form.Mcm3001uForm.RowForm;
import com.daifuku.mcm.form.Mcm3001uForm.TenpuRowForm;
import com.daifuku.mcm.repository.Mcm3001uRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * 【変換元】Mcm3001uScreen.vb
 * MCM3001U 承認申請一覧サービス
 *
 * 取引先契約（TK: kensakuKbn=1）と店舗見積（UM: kensakuKbn=2）を
 * ラジオボタン切り替えで検索・承認・差戻しする。
 *
 * 承認時の添付Excel捺印:
 *   - 店舗見積（UM）: Mcm2008uNatsuinService で捺印する（MCM2008Uと同一仕様）
 *
 * Web版省略:
 *   - 取引先契約（TK）のExcel捺印更新
 *   - メール送信
 */
@Service
public class Mcm3001uService {

    @Autowired
    private Mcm3001uRepository repository;

    /**
     * 【移植】元VB: Mcm3001uExcel.UpdateExcelReport()（KENSAKU_KBN_TENPO）
     *   店舗見積の捺印はMCM2008Uと同じ仕様（対象シート「見積総括表」「保守見積総括表(原価）」、
     *   図形 shinsa_* / shonin_*、見本 stp_*）のため、MCM2008Uの捺印サービスを再利用する。
     */
    @Autowired
    private Mcm2008uNatsuinService natsuinService;

    private static final int KENSAKU_KBN_TK = 1;

    /** 添付ファイルの保存ルート。MCM2003U/MCM2008U と同一の保管場所を参照する。 */
    @Autowired private com.daifuku.mcm.common.FileStorageService storage;

    // ===================================================================
    // 検索
    // 【変換元】Mcm3001uScreen.vb KensakuButton_Click → Search()
    //   kensakuKbn=1: MCM_TK_KEIYAKU / MCM_TK_KIKAN
    //   kensakuKbn=2: MCM_UM_KIHON_MITSUMORI
    //   各行に添付ファイル一覧をセット（N+1: VB版と同様の単一IDクエリ）
    // ===================================================================

    @Transactional(readOnly = true)
    public List<RowForm> search(Mcm3001uForm form, UserInfo userInfo) {
        List<String> syouninJotaiList = buildSyouninJotaiList(form);
        List<String> jotaiList = buildJotaiList(form);

        int kensakuKbn = form.getKensakuKbn();
        List<RowForm> rows;
        if (kensakuKbn == KENSAKU_KBN_TK) {
            rows = repository.searchTorihikisaki(
                form.getNonyusakiCd(), form.getNonyusakiNk(),
                form.getSupportId(), form.getPlantNk(),
                syouninJotaiList, jotaiList);
            for (RowForm row : rows) {
                if (row.getRelationId() != null) {
                    // TK: TK_KIKAN_ID = relationId
                    row.setTenpuList(repository.findTenpuByKikanId(row.getRelationId()));
                }
            }
        } else {
            rows = repository.searchTenpo(
                form.getNonyusakiCd(), form.getNonyusakiNk(),
                form.getSupportId(), form.getPlantNk(),
                syouninJotaiList, jotaiList);
            for (RowForm row : rows) {
                if (row.getRelationId() != null) {
                    // UM: UM_KIHON_MITSUMORI_ID = relationId
                    row.setTenpuList(repository.findTenpuByKihonMitsumoriId(row.getRelationId()));
                }
            }
        }

        // 【移植】元VB Mcm3001uScreen 検索後のグリッド制御（金額チェック）。
        //   各行の承認状態・金額・限度額から CHECK_FLG の ReadOnly 相当を決定する。
        for (RowForm row : rows) {
            row.setCheckDisabled(isCheckDisabled(row, kensakuKbn, userInfo));
        }
        return rows;
    }

    /**
     * チェックボックスを選択不可（VB版 CHECK_FLG.ReadOnly=True 相当）とすべきか判定する。
     * 【移植】元VB Mcm3001uScreen 検索後グリッド制御:
     *   ・審査中(1): 金額 > 審査限度額（TK=KeiyakuShinsaKin / UM=MitsumoriShinsaKin）→ 不可
     *   ・承認中(2): 金額 > 承認限度額（TK=KeiyakuShoninKin / UM=MitsumoriShoninKin）→ 不可
     *   ・上記以外（承認済(3)/差戻中(4)等）: 無条件で不可
     * 限度額判定は承認処理（approve）と同一条件を用い、判定ロジックの重複を避ける。
     *
     * @param row        対象行
     * @param kensakuKbn 申請内容区分（1=TK, 2=UM）
     * @param userInfo   ログインユーザー情報（限度額の取得元。null時は限度額判定なし）
     * @return 選択不可とする場合 true
     */
    private boolean isCheckDisabled(RowForm row, int kensakuKbn, UserInfo userInfo) {
        String syouninJotai = row.getSyouninJotai();
        BigDecimal kingaku = row.getKingaku();

        if (McmConstants.SHONINJOTAI_SHINSACHU_CD.equals(syouninJotai)) {
            // 審査中: 金額が審査限度額を超過している場合は選択不可
            Integer shinsaKin = (userInfo == null) ? null
                : (kensakuKbn == KENSAKU_KBN_TK
                    ? userInfo.getKeiyakuShinsaKin() : userInfo.getMitsumoriShinsaKin());
            return isOverLimit(kingaku, shinsaKin);
        }
        if (McmConstants.SHONINJOTAI_SHONINCHU_CD.equals(syouninJotai)) {
            // 承認中: 金額が承認限度額を超過している場合は選択不可
            Integer shoninKin = (userInfo == null) ? null
                : (kensakuKbn == KENSAKU_KBN_TK
                    ? userInfo.getKeiyakuShoninKin() : userInfo.getMitsumoriShoninKin());
            return isOverLimit(kingaku, shoninKin);
        }
        // 審査中/承認中以外（承認済・差戻中など）は選択不可
        return true;
    }

    /** 金額が限度額を超過しているか（approve と同一の比較条件）。 */
    private boolean isOverLimit(BigDecimal kingaku, Integer limit) {
        return limit != null && kingaku != null
            && kingaku.compareTo(BigDecimal.valueOf(limit)) > 0;
    }

    // ===================================================================
    // 承認処理
    // 【変換元】Mcm3001uScreen.vb SYONINButton_Click
    //   TK: 審査中(1)→承認中(2)  approveShinsachu(tkKeiyakuId, tkKikanId)
    //       承認中(2)→承認済(3)  approveShoninchu(tkKeiyakuId, tkKikanId)
    //   UM: 審査中(1)→承認中(2)  approveShinsachu(umKihonMitsumoriId)
    //       承認中(2)→承認済(3)  approveShoninchu(umKihonMitsumoriId)
    //
    //   金額チェック（超過時は McmBusinessException）:
    //     TK審査中: kingaku > keiyakuShinsaKin
    //     TK承認中: kingaku > keiyakuShoninKin
    //     UM審査中: kingaku > mitsumoriShinsaKin
    //     UM承認中: kingaku > mitsumoriShoninKin
    // ===================================================================

    @Transactional(rollbackFor = Exception.class)
    public void approve(List<RowForm> checkedRows, int kensakuKbn, UserInfo userInfo, String loginUser) {
        if (kensakuKbn == KENSAKU_KBN_TK) {
            Integer shinsaKin = (userInfo != null) ? userInfo.getKeiyakuShinsaKin() : null;
            Integer shoninKin = (userInfo != null) ? userInfo.getKeiyakuShoninKin() : null;
            for (RowForm row : checkedRows) {
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
                    // relationId = TK_KIKAN_ID, tkKeiyakuId = TK_KEIYAKU_ID
                    repository.approveTkShinsachu(row.getTkKeiyakuId(), row.getRelationId(), loginUser);
                } else if (McmConstants.SHONINJOTAI_SHONINCHU_CD.equals(syouninJotai)) {
                    if (shoninKin != null && kingaku != null
                            && kingaku.compareTo(BigDecimal.valueOf(shoninKin)) > 0) {
                        throw new McmBusinessException(
                            "書類NO[" + row.getShoruiNo() + "]の金額（"
                            + kingaku.toPlainString() + "円）が承認限度額（"
                            + shoninKin + "円）を超えています。");
                    }
                    repository.approveTkShoninchu(row.getTkKeiyakuId(), row.getRelationId(), loginUser);
                }
            }
        } else {
            Integer shinsaKin = (userInfo != null) ? userInfo.getMitsumoriShinsaKin() : null;
            Integer shoninKin = (userInfo != null) ? userInfo.getMitsumoriShoninKin() : null;
            // 捺印の名前欄に使う氏名（MCM_MO_TANTO.TANTO_NK）。姓の抽出は捺印サービス側で行う。
            String userName = (userInfo != null) ? userInfo.getUserName() : null;
            for (RowForm row : checkedRows) {
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
                    // 【変換元】SYONINButton_Click(KENSAKU_KBN_TENPO): 捺印押下（審査印）→ 成功後に承認状態を更新
                    //   relationId = UM_KIHON_MITSUMORI_ID
                    natsuinService.stamp(row.getRelationId(), syouninJotai, userName);
                    repository.approveUmShinsachu(row.getRelationId(), loginUser);
                } else if (McmConstants.SHONINJOTAI_SHONINCHU_CD.equals(syouninJotai)) {
                    if (shoninKin != null && kingaku != null
                            && kingaku.compareTo(BigDecimal.valueOf(shoninKin)) > 0) {
                        throw new McmBusinessException(
                            "書類NO[" + row.getShoruiNo() + "]の金額（"
                            + kingaku.toPlainString() + "円）が承認限度額（"
                            + shoninKin + "円）を超えています。");
                    }
                    // 【変換元】SYONINButton_Click(KENSAKU_KBN_TENPO): 捺印押下（承認印）→ 成功後に承認状態を更新
                    natsuinService.stamp(row.getRelationId(), syouninJotai, userName);
                    repository.approveUmShoninchu(row.getRelationId(), loginUser);
                }
            }
        }
    }

    // ===================================================================
    // 差戻し処理
    // 【変換元】Mcm3001uScreen.vb SASHIMODOSHIButton_Click
    //   審査中(1) / 承認中(2) のみ差戻し可
    // ===================================================================

    @Transactional(rollbackFor = Exception.class)
    public void reject(List<RowForm> checkedRows, int kensakuKbn, String loginUser) {
        if (kensakuKbn == KENSAKU_KBN_TK) {
            for (RowForm row : checkedRows) {
                String syouninJotai = row.getSyouninJotai();
                if (McmConstants.SHONINJOTAI_SHINSACHU_CD.equals(syouninJotai) ||
                    McmConstants.SHONINJOTAI_SHONINCHU_CD.equals(syouninJotai)) {
                    repository.rejectTk(row.getTkKeiyakuId(), row.getRelationId(), loginUser);
                }
            }
        } else {
            for (RowForm row : checkedRows) {
                String syouninJotai = row.getSyouninJotai();
                if (McmConstants.SHONINJOTAI_SHINSACHU_CD.equals(syouninJotai) ||
                    McmConstants.SHONINJOTAI_SHONINCHU_CD.equals(syouninJotai)) {
                    repository.rejectUm(row.getRelationId(), loginUser);
                }
            }
        }
    }

    // ===================================================================
    // 添付ファイルダウンロード
    // 【移植】MCM2008U(Mcm2008uService)の添付資料ダウンロードと同一の
    //   保管場所・検証ロジックを使用する。TK/UM で添付検索元テーブルが異なる。
    //     kensakuKbn=1(TK): MCM_TK_TENPU（relationId = TK_KIKAN_ID）
    //     kensakuKbn=2(UM): MCM_UM_TENPU（relationId = UM_KIHON_MITSUMORI_ID）
    // ===================================================================

    /**
     * 添付ファイル行を取得する（ダウンロード時のファイル名決定用）。
     *
     * @param kensakuKbn 申請内容区分（1=TK, 2=UM）
     * @param relationId TK時=TK_KIKAN_ID / UM時=UM_KIHON_MITSUMORI_ID
     * @param tenpuId    添付ID（MCM_TK_TENPU.TK_TENPU_ID / MCM_UM_TENPU.UM_TENPU_ID）
     * @return 添付ファイル行
     */
    @Transactional(readOnly = true)
    public TenpuRowForm findAttachmentRow(int kensakuKbn, BigDecimal relationId, BigDecimal tenpuId) {
        return findTenpuList(kensakuKbn, relationId).stream()
            .filter(t -> t.getTenpuId() != null && t.getTenpuId().compareTo(tenpuId) == 0)
            .findFirst()
            .orElseThrow(() -> new IllegalStateException("添付資料が見つかりません。"));
    }

    /**
     * 添付ファイルの物理パスを解決する。
     * 保存ルート配下の通常ファイルであることを検証し、ディレクトリトラバーサルを防止する。
     *
     * @param kensakuKbn 申請内容区分（1=TK, 2=UM）
     * @param relationId TK時=TK_KIKAN_ID / UM時=UM_KIHON_MITSUMORI_ID
     * @param tenpuId    添付ID
     * @return 添付ファイルの物理パス
     * @throws java.io.IOException 保存先が不正な場合、またはファイルが存在しない場合
     */
    @Transactional(readOnly = true)
    public java.nio.file.Path resolveAttachment(int kensakuKbn, BigDecimal relationId, BigDecimal tenpuId)
            throws java.io.IOException {

        TenpuRowForm target = findAttachmentRow(kensakuKbn, relationId, tenpuId);
        return storage.resolve(target.getDirectory(), target.getTenpufileNk());
    }

    private List<TenpuRowForm> findTenpuList(int kensakuKbn, BigDecimal relationId) {
        if (kensakuKbn == KENSAKU_KBN_TK) {
            return repository.findTenpuByKikanId(relationId);
        }
        return repository.findTenpuByKihonMitsumoriId(relationId);
    }

    // ===================================================================
    // private
    // ===================================================================

    private List<String> buildSyouninJotaiList(Mcm3001uForm form) {
        List<String> list = new ArrayList<>();
        if (form.isShinsachu())    list.add(McmConstants.SHONINJOTAI_SHINSACHU_CD);
        if (form.isShoninchu())    list.add(McmConstants.SHONINJOTAI_SHONINCHU_CD);
        if (form.isShoninzumi())   list.add(McmConstants.SHONINJOTAI_SHONINZUMI_CD);
        if (form.isSashimodoshi()) list.add(McmConstants.SHONINJOTAI_SASHIMODOSHI_CD);
        return list;
    }

    private List<String> buildJotaiList(Mcm3001uForm form) {
        List<String> list = new ArrayList<>();
        if (form.isJotaiMitsumori()) list.add(McmConstants.JOTAI_MITSUMORI);
        if (form.isJotaiKeiyaku())   list.add(McmConstants.JOTAI_KEIYAKU);
        if (form.isJotaiHaki())      list.add(McmConstants.JOTAI_HAKI);
        if (form.isJotaiKaiyaku())   list.add(McmConstants.JOTAI_KAIYAKU);
        return list;
    }
}
