/**
 * 【変換元】Mcm0011uScreen.vb
 * 【説明】納入先マスタ（単票編集画面）のサービスクラス
 * 元メソッド対応:
 *   Search()         → findByNonyusakiId() / initNewForm()
 *   UpdateProcess()  → save()
 *   EnterUpdate()    → save() 内部
 *   削除処理         → delete()
 *   CountNonyusakiId → countPlantByNonyusakiId()
 */
package com.daifuku.mcm.service;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.Optional;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.daifuku.mcm.common.ExclusiveLockKey;
import com.daifuku.mcm.common.MessageService;
import com.daifuku.mcm.entity.NonyusakiEntity;
import com.daifuku.mcm.exception.ExclusiveControlException;
import com.daifuku.mcm.form.Mcm0011uForm;
import com.daifuku.mcm.repository.NonyusakiRepository;

@Service
public class Mcm0011uService {

    /** 【追加 #296】排他判定の調査ログ出力用 */
    private static final Logger log = LoggerFactory.getLogger(Mcm0011uService.class);

    @Autowired
    private NonyusakiRepository nonyusakiRepository;

    @Autowired
    private JdbcTemplate jdbcTemplate;

    /** 【不具合修正 #296】排他エラーメッセージ（mcm0011u.msg.exclusive）の外部化に使用 */
    @Autowired
    private MessageService messageService;

    /**
     * 納入先IDで検索（編集モード）
     * 元VB: Me.MCM_MA_NONYUSAKITableAdapter.Fill(... , nonyusakiId)
     */
    public Optional<NonyusakiEntity> findByNonyusakiId(BigDecimal nonyusakiId) {
        return nonyusakiRepository.findById(nonyusakiId);
    }

    public String currentLockKey(BigDecimal id) {
        return ExclusiveLockKey.format(jdbcTemplate.queryForObject(
                "SELECT LASTUPDATE_DT FROM MCM.MCM_MA_NONYUSAKI WHERE NONYUSAKI_ID=?",
                (rs, row) -> rs.getTimestamp(1) == null ? null : rs.getTimestamp(1).toLocalDateTime(), id));
    }

    /**
     * 新規モード用のフォーム初期化
     * 元VB: nonyusakiNewRow.DTSRENKEI_FLG = 0, nonyusakiNewRow.NONYUSAKI_ID = 1
     */
    public Mcm0011uForm initNewForm() {
        Mcm0011uForm form = new Mcm0011uForm();
        form.setNonyusakiId(BigDecimal.ZERO);
        form.setDtsrenkeiFlg(BigDecimal.ZERO);
        form.setNonyusakiSakujo(false);
        form.setNewMode(true);
        // 【不具合修正 #296】新規モードは排他対象レコードが存在しないため空
        form.setLockLastupdateDt("");
        return form;
    }

    /**
     * エンティティからフォームに変換（編集モード）
     */
    public Mcm0011uForm entityToForm(NonyusakiEntity entity) {
        Mcm0011uForm form = new Mcm0011uForm();
        form.setNonyusakiId(entity.getNonyusakiId());
        form.setNonyusakiCd(entity.getNonyusakiCd());
        form.setNonyusakiNk(entity.getNonyusakiNk());
        form.setKyunonyusakiNk(entity.getKyunonyusakiNk());
        form.setNonyusakikanaKn(entity.getNonyusakikanaKn());
        form.setNonyusakieimeiEn(entity.getNonyusakieimeiEn());
        form.setNonyusakikojoNk(entity.getNonyusakikojoNk());
        form.setYubinNo(entity.getYubinNo());
        form.setKuniNk(entity.getKuniNk());
        form.setJusyo1Nk(entity.getJusyo1Nk());
        form.setJusyo2Nk(entity.getJusyo2Nk());
        form.setTelNo(entity.getTelNo());
        form.setFaxNo(entity.getFaxNo());
        form.setBiko(entity.getBiko());
        form.setDtsrenkeiFlg(entity.getDtsrenkeiFlg() != null
            ? entity.getDtsrenkeiFlg() : BigDecimal.ZERO);
        form.setNonyusakiSakujo(false);
        form.setNewMode(false);
        // 【不具合修正 #296】表示時点の LASTUPDATE_DT を排他キーとして画面へ渡す
        form.setLockLastupdateDt(ExclusiveLockKey.format(entity.getLastupdateDt()));
        return form;
    }

    /**
     * 保存処理（新規INSERT / 既存UPDATE）
     * 元VB: UpdateProcess() → EnterUpdate() → UpdateAll()
     *
     * ※ 元VBのUPDATE文ではWHERE句で全カラムのOriginal値比較による楽観的排他制御を
     *    実施していた（MCM_COMMON_PAC.FUN_NEC_CONV()によるNEC特殊文字変換付き）。
     *    SQL Serverではタイムスタンプ（LASTUPDATE_DT）比較による簡易排他制御とする。
     *    MCM_COMMON_PAC.FUN_NEC_CONV → SQL Server版ファンクション作成後に対応予定。
     *
     * 【不具合修正 #296】上記コメントの「LASTUPDATE_DT比較による簡易排他制御」が
     *   未実装のままだったため、同一レコードを2クライアントで同時編集すると
     *   後勝ちで上書きされ、先行ユーザーの更新内容が失われていた
     *   （試験項目 INT07-059 NG）。MCM0022U / MCM0025U / MCM3006U と同一方式で実装する。
     *
     *   判定は2段構え。
     *   (1) 画面が保持していた LASTUPDATE_DT と DB の現在値を Java 側で比較する。
     *   (2) UPDATE の WHERE 句でも LASTUPDATE_DT を照合し、更新件数0件を競合とする。
     *       (1) から実際の更新までのわずかな時間差で割り込まれるケースを塞ぐ。
     *
     *   ※ 新規登録（INSERT）は排他対象レコードが存在しないためチェック対象外。
     *
     * @param form フォーム入力値
     * @param userId ログインユーザID
     * @return 保存後の納入先ID
     * @throws ExclusiveControlException 排他競合を検知した場合（DBは更新しない）
     */
    @Transactional
    public BigDecimal save(Mcm0011uForm form, String userId) {
        LocalDateTime now = LocalDateTime.now();

        if (form.getNonyusakiId() == null
            || form.getNonyusakiId().compareTo(BigDecimal.ZERO) == 0) {
            // ===== 新規INSERT =====
            // 元VB: INSERT INTO MCM_MA_NONYUSAKI (19カラム)
            BigDecimal newId = jdbcTemplate.queryForObject(
                "SELECT ISNULL(MAX(NONYUSAKI_ID), 0) + 1 FROM MCM.MCM_MA_NONYUSAKI",
                BigDecimal.class);

            jdbcTemplate.update(
            	    "INSERT INTO MCM.MCM_MA_NONYUSAKI (" +
            	    "NONYUSAKI_ID, NONYUSAKI_CD, NONYUSAKI_NK, KYUNONYUSAKI_NK, " +
            	    "NONYUSAKIKOJO_NK, NONYUSAKIKANA_KN, NONYUSAKIEIMEI_EN, " +
            	    "YUBIN_NO, JUSYO1_NK, JUSYO2_NK, KUNI_NK, TEL_NO, FAX_NO, " +
            	    "BIKO, DTSRENKEI_FLG, CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY" +
            	    ") VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                newId,
                form.getNonyusakiCd(), form.getNonyusakiNk(), form.getKyunonyusakiNk(),
                form.getNonyusakikojoNk(), form.getNonyusakikanaKn(), form.getNonyusakieimeiEn(),
                form.getYubinNo(), form.getJusyo1Nk(), form.getJusyo2Nk(),
                form.getKuniNk(), form.getTelNo(), form.getFaxNo(),
                form.getBiko(),
                form.getDtsrenkeiFlg() != null ? form.getDtsrenkeiFlg().toString() : "0",
                now, userId, now, userId
            );
            return newId;
        } else {
            // ===== 既存UPDATE =====
            // 元VB: UPDATE MCM_MA_NONYUSAKI SET ... WHERE NONYUSAKI_ID = :Original_NONYUSAKI_ID

            // 【不具合修正 #296】排他チェック（1段目）
            //   入力内容の妥当性より先に判定する。画面が古い時点の情報を前提に
            //   しているため、入力値が何であれ処理を続行してはならない。
            LocalDateTime dbLastupdateDt = readLastupdateDtForUpdate(
                    form.getNonyusakiId(), form.getLockLastupdateDt());
            now = ExclusiveLockKey.nextVersion(dbLastupdateDt);

            // 更新にはDBから読み取った値を渡す。画面から往復した値ではなくDB由来の値を
            // 使うことで、LASTUPDATE_DT の型精度（Oracle DATE=秒精度 /
            // SQL Server DATETIME=約3.33ms丸め）による誤判定を回避する。
            // 画面保持値との突き合わせは上記1段目で完了している。
        	int updated = jdbcTemplate.update(
        		    "UPDATE MCM.MCM_MA_NONYUSAKI SET " +
        		    "NONYUSAKI_CD = ?, NONYUSAKI_NK = ?, KYUNONYUSAKI_NK = ?, " +
        		    "NONYUSAKIKOJO_NK = ?, NONYUSAKIKANA_KN = ?, NONYUSAKIEIMEI_EN = ?, " +
        		    "YUBIN_NO = ?, JUSYO1_NK = ?, JUSYO2_NK = ?, KUNI_NK = ?, " +
        		    "TEL_NO = ?, FAX_NO = ?, BIKO = ?, DTSRENKEI_FLG = ?, " +
        		    "LASTUPDATE_DT = ?, LASTUPDATE_BY = ? " +
        		    "WHERE NONYUSAKI_ID = ? " +
        		    // 【不具合修正 #296】排他チェック（2段目）
        		    "AND ((? IS NULL AND LASTUPDATE_DT IS NULL) OR LASTUPDATE_DT = ?)",
                form.getNonyusakiCd(), form.getNonyusakiNk(), form.getKyunonyusakiNk(),
                form.getNonyusakikojoNk(), form.getNonyusakikanaKn(), form.getNonyusakieimeiEn(),
                form.getYubinNo(), form.getJusyo1Nk(), form.getJusyo2Nk(),
                form.getKuniNk(), form.getTelNo(), form.getFaxNo(),
                form.getBiko(),
                form.getDtsrenkeiFlg() != null ? form.getDtsrenkeiFlg().toString() : "0",
                now, userId,
                form.getNonyusakiId(),
                dbLastupdateDt, dbLastupdateDt
            );

            if (updated == 0) {
                throw exclusiveError();
            }
            return form.getNonyusakiId();
        }
    }

    /**
     * 【不具合修正 #296】更新／削除の直前に、DB上の LASTUPDATE_DT を読み取り、
     * 画面が保持していた排他キーと一致することを確認する。
     *
     * @param nonyusakiId 対象の納入先ID
     * @param screenLockKey 画面が hidden で保持していた排他キー文字列
     * @return DBから読み取った LASTUPDATE_DT（UPDATE / DELETE の WHERE 句へ渡す）
     * @throws ExclusiveControlException レコードが既に削除済み、または他ユーザーが更新済みの場合
     */
    private LocalDateTime readLastupdateDtForUpdate(BigDecimal nonyusakiId, String screenLockKey) {
        NonyusakiEntity current = nonyusakiRepository.findById(nonyusakiId).orElse(null);
        // 【#296 調査用ログ】排他判定に使った実値を必ず残す。
        //   screenLockKey が空 → 画面(hidden)で往復していない
        //   parsed が null かつ screenLockKey が非空 → 書式不正
        //   parsed == dbValue → 競合なしと判定される
        log.info("MCM0011U 排他判定 nonyusakiId={} screenLockKey='{}' parsed={} dbValue={}",
                 nonyusakiId, screenLockKey, ExclusiveLockKey.parse(screenLockKey),
                 current == null ? null : current.getLastupdateDt());
        if (current == null) {
            // 他ユーザーが削除済み。VBでは更新件数0件 → DBConcurrencyException となるケース。
            throw exclusiveError();
        }
        if (!ExclusiveLockKey.matches(current.getLastupdateDt(),
                                      ExclusiveLockKey.parse(screenLockKey))) {
            throw exclusiveError();
        }
        return current.getLastupdateDt();
    }

    /** 【不具合修正 #296】排他エラー（FWM_0009 相当）を生成する */
    private ExclusiveControlException exclusiveError() {
        return new ExclusiveControlException(
                messageService.getMessage("mcm0011u.msg.exclusive"));
    }

    /**
     * 納入先削除
     * 元VB: 削除チェックON → CountNonyusakiId → 0なら削除
     *       → Rows(0).Delete → UpdateAll → データクリア → 新規モードへ
     *
     * 【不具合修正 #296】削除も更新と同じ方式で排他判定する。
     *   元VBの DELETE 文も WHERE 句に Original_* 値を持つ楽観的同時実行制御が有効だった。
     *
     * @param nonyusakiId 削除対象の納入先ID
     * @param screenLockKey 画面が保持していた排他キー（LASTUPDATE_DT）
     * @return true: 削除成功, false: プラント使用中で削除不可(MSG_0006)
     * @throws ExclusiveControlException 排他競合を検知した場合（DBは更新しない）
     */
    @Transactional
    public boolean delete(BigDecimal nonyusakiId, String screenLockKey) {
        // 【不具合修正 #296】排他チェック（1段目）。使用チェックより先に判定する。
        LocalDateTime dbLastupdateDt = readLastupdateDtForUpdate(nonyusakiId, screenLockKey);

        // プラント使用チェック
        // 元VB: SELECT COUNT(PLANT_ID) FROM MCM_MA_PLANT
        //       WHERE NONYUSAKI_ID = :nonyusakiId AND ROWNUM <= 1
        // → SQL Server: SELECT COUNT(PLANT_ID) FROM MCM_MA_PLANT WHERE NONYUSAKI_ID = ?
        int plantCount = countPlantByNonyusakiId(nonyusakiId);
        if (plantCount > 0) {
            // MSG_0006: プラントマスタで使用されているため削除不可
            return false;
        }
        // 削除実行
        // 元VB: DELETE FROM MCM_MA_NONYUSAKI WHERE NONYUSAKI_ID = :Original_NONYUSAKI_ID
        // 【不具合修正 #296】排他チェック（2段目）
        int deleted = jdbcTemplate.update(
            "DELETE FROM MCM.MCM_MA_NONYUSAKI WHERE NONYUSAKI_ID = ?" +
            " AND ((? IS NULL AND LASTUPDATE_DT IS NULL) OR LASTUPDATE_DT = ?)",
            nonyusakiId, dbLastupdateDt, dbLastupdateDt
        );
        if (deleted == 0) {
            throw exclusiveError();
        }
        return true;
    }

    /**
     * プラント使用件数チェック
     * 元VB: CommandCollection(1) → SELECT COUNT(PLANT_ID)
     *       FROM MCM_MA_PLANT WHERE NONYUSAKI_ID = :nonyusakiId AND ROWNUM <= 1
     * → SQL Server: ROWNUMは不要（COUNTで十分）
     */
    public int countPlantByNonyusakiId(BigDecimal nonyusakiId) {
        Integer count = jdbcTemplate.queryForObject(
            "SELECT COUNT(PLANT_ID) FROM MCM.MCM_MA_PLANT WHERE NONYUSAKI_ID = ?",
            Integer.class, nonyusakiId);
        return count != null ? count : 0;
    }
}
