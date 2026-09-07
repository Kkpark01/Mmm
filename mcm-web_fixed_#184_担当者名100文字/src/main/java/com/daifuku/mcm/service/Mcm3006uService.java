/**
 * 【変換元】Mcm3006uScreen.vb
 * 【説明】担当者マスタ（MCM3006U）サービスクラス
 *
 * 元イベント対応:
 *   Mcm3006uScreen_Load / Search()  → loadData()
 *   RowInsertButton_Click            → prepareNewRow() ※画面側処理のため実際の行追加はJS
 *   UpdateButton_Click               → save()
 *
 * 【不具合修正 #119】
 *   NG①: 登録日時/登録者がデータ更新時に更新されていた
 *         → 権限テーブルを「全削除→全INSERT」していたため CREATED_DT/CREATED_BY が上書きされていた。
 *            既存レコードがあれば UPDATE（CREATED_DT/BY は変更しない SQL）、なければ INSERT に変更。
 *   NG②: 更新対象外のログインIDの権限まで全件更新されていた
 *         → フォームから全担当者の権限が送信されるため updateTantoIds = 全担当者ID となり、
 *            全員分が削除→再INSERT されていた。
 *            DBの業務項目と比較し、担当者情報・権限それぞれの変更行だけを更新する。
 *   NG③: 登録者・更新者にログインIDが設定されていた
 *         → loginUser は Controller から担当者名（TANTO_NK）を受け取るよう変更。
 *            Controller 側で UserInfo.getUserName() を渡す。
 */
package com.daifuku.mcm.service;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.entity.TantoEntity;
import com.daifuku.mcm.entity.TantoKengenEntity;
import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.repository.KengenBunruiRepository;
import com.daifuku.mcm.repository.Mcm3006uRepository;

@Service
public class Mcm3006uService {

    @Autowired
    private Mcm3006uRepository repository;

    @Autowired
    private KengenBunruiRepository kengenBunruiRepository;

    /**
     * 画面初期表示データ取得
     * 元VB: Mcm3006uScreen_Load → Search()
     *   MCM_MO_TANTOTableAdapter.Fill()
     *   MCM_MO_TANTOKENGENTableAdapter.Fill()
     *
     * @return 担当者一覧
     */
    @Transactional(readOnly = true)
    public List<TantoEntity> loadTantoList() {
        return repository.findAll();
    }

    /**
     * 担当者権限一覧取得
     * 元VB: MCM_MO_TANTOKENGENTableAdapter.Fill()
     *
     * @return 担当者権限一覧
     */
    @Transactional(readOnly = true)
    public List<TantoKengenEntity> loadTantoKengenList() {
        // 【修正】findAllKengen() は MCM_MO_TANTOKENGEN の実データ行しか返さないため、
        //         権限が1件も登録されていない担当者を選択すると利用権限グリッドが空になっていた。
        //         元VB の DataGridView は権限分類マスタ全件（納入機器関連ほか）を表示するため、
        //         担当者×権限分類マスタ全件を返す findAllKengenCross() に切り替える。
        //         findAllKengen() は削除せず Repository に残してある。
        return repository.findAllKengenCross();
    }

    /**
     * 登録・更新・削除処理
     * 元VB: UpdateButton_Click
     *   1. ログインID重複チェック → MSG_0033
     *   2. 上位者ID存在チェック   → MSG_0039
     *   3. UpdateAll(MCM_MO_TANTO, MCM_MO_TANTOKENGEN, "TANTO_ID")
     *
     * 【不具合修正 #119】
     *   loginUser は担当者名（TANTO_NK）を受け取ること。
     *   Controller で UserInfo.getUserName() を渡すよう変更済み。
     *
     * @param updateList   画面から送信された担当者リスト（未変更行を含む）
     * @param deleteIds    削除対象の担当者IDリスト
     * @param kengenList   更新対象の担当者権限リスト（全担当者分が渡される）
     * @param loginUser    ログインユーザーの担当者名（TANTO_NK）
     */
    @Transactional(rollbackFor = Exception.class)
    public void save(
            List<TantoEntity> updateList,
            List<BigDecimal> deleteIds,
            List<TantoKengenEntity> kengenList,
            String loginUser) {

        // #184: ブラウザーの検証を経由しない送信も、DB更新前に拒否する。
        for (TantoEntity tanto : updateList) {
            if (deleteIds.stream().anyMatch(id -> id.compareTo(tanto.getTantoId()) == 0)) continue;
            String name = tanto.getTantoNk();
            if (name != null && name.length() > 100) {
                throw new McmBusinessException("担当者は100桁以下で入力してください。");
            }
        }

        // ── 1. ログインID重複チェック ──
        // 元VB: FOR i=0 TO rowCount-2, FOR j=i+1 TO rowCount-1
        //       If loginId.Equals(...) Then DisplayMessage(MSG_0033)
        for (int i = 0; i < updateList.size(); i++) {
            String loginId = updateList.get(i).getLoginId();
            if (loginId == null || loginId.trim().isEmpty()) {
                continue;
            }
            for (int j = i + 1; j < updateList.size(); j++) {
                if (loginId.equals(updateList.get(j).getLoginId())) {
                    throw new McmBusinessException("ログインIDが重複しています。");
                }
            }
        }

        // ── 2. 上位者ID存在チェック ──
        // 元VB: 上位者IDが設定されている場合、担当者一覧内に存在するか確認
        //       存在しない場合 DisplayMessage(MSG_0039)
        List<String> allLoginIds = updateList.stream()
                .map(TantoEntity::getLoginId)
                .filter(id -> id != null && !id.trim().isEmpty())
                .collect(Collectors.toList());

        for (TantoEntity tanto : updateList) {
            String oyaLoginId = tanto.getOyaloginId();
            if (oyaLoginId != null && !oyaLoginId.trim().isEmpty()) {
                boolean exists = allLoginIds.stream().anyMatch(id -> id.equals(oyaLoginId));
                if (!exists) {
                    throw new McmBusinessException("上位者IDが担当者一覧に存在しません。");
                }
            }
        }

        // ── 3. DB値と比較し、変更した行だけを更新 ──
        // IDの桁数表現（1 / 1.0）で別レコードと誤判定しない。
        Set<BigDecimal> deletedIds = deleteIds.stream()
                .map(Mcm3006uService::idKey).collect(Collectors.toSet());
        Map<BigDecimal, TantoEntity> existingTanto = new HashMap<>();
        for (TantoEntity e : repository.findAll()) {
            existingTanto.put(idKey(e.getTantoId()), e);
        }

        // 削除する担当者は、送信データに残っていても再登録しない。
        for (BigDecimal id : deletedIds) {
            repository.deleteTantoKengenByTantoId(id);
            repository.deleteTanto(id);
        }

        Set<BigDecimal> submittedIds = new java.util.HashSet<>();
        for (TantoEntity tanto : updateList) {
            BigDecimal id = idKey(tanto.getTantoId());
            if (deletedIds.contains(id)) continue;
            submittedIds.add(id);
            TantoEntity old = existingTanto.get(id);
            if (old == null) {
                repository.insertTanto(tanto, loginUser);
            } else {
                // この画面にパスワード編集欄はない。hiddenの既定値や古い値で
                // パスワード変更・初期化画面の結果を上書きしない。
                tanto.setPassword(old.getPassword());
                if (!sameTanto(old, tanto)) {
                    repository.updateTanto(tanto, loginUser);
                }
            }
        }

        // 担当者情報を変更せず権限だけ変更した場合も、独立して判定する。
        Map<BigDecimal, Map<BigDecimal, TantoKengenEntity>> existingKengen = new HashMap<>();
        for (BigDecimal id : submittedIds) {
            Map<BigDecimal, TantoKengenEntity> byBunrui = new HashMap<>();
            if (existingTanto.containsKey(id)) {
                for (TantoKengenEntity e : repository.findKengenByTantoId(id)) {
                    byBunrui.put(idKey(e.getKengenbunruiId()), e);
                }
            }
            existingKengen.put(id, byBunrui);
        }
        for (TantoKengenEntity kengen : kengenList) {
            BigDecimal id = idKey(kengen.getTantoId());
            if (!submittedIds.contains(id)) continue;
            Map<BigDecimal, TantoKengenEntity> byBunrui = existingKengen.get(id);
            BigDecimal bunruiId = idKey(kengen.getKengenbunruiId());
            TantoKengenEntity old = byBunrui.get(bunruiId);
            if (old != null) {
                if (!sameKengen(old, kengen)) {
                    // UPDATE SQLは登録日時・登録者を変更しない。
                    repository.updateTantoKengen(kengen, loginUser);
                }
            } else if (!existingTanto.containsKey(id) || !isDefaultKengen(kengen)) {
                // CROSS JOINが表示用に補完した「なし・備考空」の行は
                // 既存担当者では未変更として扱う。新規担当者の初期権限は登録する。
                repository.insertTantoKengen(kengen, loginUser);
            }
            byBunrui.put(bunruiId, kengen);
        }
    }

    private static BigDecimal idKey(BigDecimal id) {
        return id == null ? null : id.stripTrailingZeros();
    }

    // HTML入力欄ではDBのNULLも空文字で戻るため、空欄同士は未変更とする。
    private static String text(String value) {
        return value == null ? "" : value;
    }

    private static String flag(String value) {
        return value == null || value.isEmpty() ? "0" : value;
    }

    private static int flag(Integer value) {
        return value == null ? 0 : value;
    }

    private static boolean sameNumber(BigDecimal a, BigDecimal b) {
        return a == null ? b == null : b != null && a.compareTo(b) == 0;
    }

    private static boolean sameTanto(TantoEntity a, TantoEntity b) {
        return text(a.getLoginId()).equals(text(b.getLoginId()))
                && text(a.getTantoNk()).equals(text(b.getTantoNk()))
                && text(a.getJigyosyoNk()).equals(text(b.getJigyosyoNk()))
                && text(a.getMailaddress()).equals(text(b.getMailaddress()))
                && text(a.getOyaloginId()).equals(text(b.getOyaloginId()))
                && flag(a.getShinsaFlg()) == flag(b.getShinsaFlg())
                && sameNumber(a.getMitsumoriKin(), b.getMitsumoriKin())
                && sameNumber(a.getKeiyakuKin(), b.getKeiyakuKin())
                && flag(a.getShoninFlg()) == flag(b.getShoninFlg())
                && sameNumber(a.getMitsumorishoninKin(), b.getMitsumorishoninKin())
                && sameNumber(a.getKeiyakushoninKin(), b.getKeiyakushoninKin())
                && flag(a.getSystemriyoFlg()).equals(flag(b.getSystemriyoFlg()))
                && flag(a.getMaintenanceFlg()).equals(flag(b.getMaintenanceFlg()))
                && text(a.getBiko()).equals(text(b.getBiko()));
    }

    private static boolean sameKengen(TantoKengenEntity a, TantoKengenEntity b) {
        return flag(a.getRiyokengenKbn()).equals(flag(b.getRiyokengenKbn()))
                && text(a.getBiko()).equals(text(b.getBiko()));
    }

    private static boolean isDefaultKengen(TantoKengenEntity e) {
        return McmConstants.RIYOKENGEN_KBN_NASI_CD.equals(flag(e.getRiyokengenKbn()))
                && text(e.getBiko()).isEmpty();
    }

    /**
     * 新規行の担当者IDを採番する
     * 元VB: tantoMax = tantoMax - 1（新規行は負数で仮IDとして管理）
     * Java版では正数で採番する（MAX+1）
     *
     * @return 新規担当者ID
     */
    @Transactional(readOnly = true)
    public BigDecimal getNextTantoId() {
        return repository.getMaxTantoId().add(BigDecimal.ONE);
    }

    /**
     * 権限分類マスタに基づく初期権限行を生成する
     * 元VB: RowInsertButton_Click内 — 権限分類テーブル全件に対して担当者権限行を追加
     *
     * @param tantoId 新規担当者ID
     * @return 初期担当者権限リスト
     */
    @Transactional(readOnly = true)
    public List<TantoKengenEntity> buildInitialKengenRows(BigDecimal tantoId) {
        List<com.daifuku.mcm.entity.KengenBunruiEntity> kengenBunruiList =
                repository.findAllKengenBunrui();

        List<TantoKengenEntity> list = new ArrayList<>();
        for (com.daifuku.mcm.entity.KengenBunruiEntity bunrui : kengenBunruiList) {
            TantoKengenEntity kengen = new TantoKengenEntity();
            kengen.setTantoId(tantoId);
            kengen.setKengenbunruiId(bunrui.getKengenbunruiId());
            kengen.setKengenbunruiNk(bunrui.getKengenbunruiNk());
            // 元VB: RIYOKENGEN_KBN = McmConstant.RIYOKENGEN_KBN_NASI_CD
            kengen.setRiyokengenKbn(McmConstants.RIYOKENGEN_KBN_NASI_CD);
            list.add(kengen);
        }
        return list;
    }

    /**
     * 権限分類マスタ一覧取得（利用権限グリッドの補完表示用）
     * 元VB: McmKengenBunruiDataSet TableAdapter.Fill()
     * ※ 既存メソッドは一切削除せず、Repository の findAllKengenBunrui() を利用する。
     */
    @Transactional(readOnly = true)
    public List<com.daifuku.mcm.entity.KengenBunruiEntity> loadKengenBunruiList() {
        return repository.findAllKengenBunrui();
    }
}
