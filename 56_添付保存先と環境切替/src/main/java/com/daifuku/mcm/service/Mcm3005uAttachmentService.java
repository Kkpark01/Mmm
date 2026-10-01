package com.daifuku.mcm.service;

import com.daifuku.mcm.exception.ExclusiveControlException;
import com.daifuku.mcm.exception.InputCheckException;
import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.form.Mcm3005uForm.TenpuRowForm;
import com.daifuku.mcm.repository.Mcm3005uRepository;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.transaction.support.TransactionSynchronization;
import org.springframework.transaction.support.TransactionSynchronizationManager;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.UUID;

/**
 * 【変換元】Mcm3005uScreen.vb FileUpdateButton_Click / FileDelButton_Click / CellContentClick
 * MCM3005U 作業予定の添付ファイル（MCM_MO_TENPU）の追加・削除・参照を行うサービス。
 *
 * VB版はクライアントPCのファイルを共有フォルダ「作業予定\登録ID」へコピーし、
 * DIRECTORY 列にファイルのフルパスを保持していた。Web版はアップロードされたファイルを
 * mcm.file.upload-path 配下の 作業予定\登録ID\(UUID)\ファイル名 に保存し、相対パスを保持する
 * （MCM1005U と同じ保存方式。UUIDフォルダはロールバック時に今回分だけを確実に消すため）。
 */
@Service
public class Mcm3005uAttachmentService {

    private static final Logger log = LoggerFactory.getLogger(Mcm3005uAttachmentService.class);

    /** MSG_0138E */
    public static final String MSG_DUPLICATE = "同一名のファイルが既にアップロードされている為、ファイル追加することは出来ません。";
    /** MSG_0109E */
    public static final String MSG_NOT_FOUND = "該当ファイルが存在しません。";
    /** アップロードファイル未選択 */
    public static final String MSG_NOT_SELECTED = "添付ファイルを選択してください。";
    /** ファイル名がWindows上で保存できない */
    public static final String MSG_BAD_NAME = "ファイル名を変更して添付してください。";
    /** 保存失敗 */
    public static final String MSG_SAVE_FAILED = "添付ファイルを保存できませんでした。保存先を確認してください。";
    /** 保存先がアップロードルート外（VB版の共有フォルダパス等） */
    public static final String MSG_OUTSIDE_ROOT = "添付資料の保存先を確認してください。";

    /** 【変換元】app.config FileUploadFolder3 = 作業予定 */
    private static final String FOLDER = "作業予定";
    /** TENPUFILE_NK / DIRECTORY は VARCHAR(2000)。保存先パスも収まるようファイル名は180文字までとする。 */
    private static final int MAX_NAME_LENGTH = 180;

    @Autowired
    private Mcm3005uRepository repository;

    @Autowired private com.daifuku.mcm.common.FileStorageService storage;

    /**
     * 添付ファイルを追加する。
     * 【変換元】FileUpdateButton_Click（同名チェック MSG_0138E → アップロード → MCM_MO_TENPU 追加 → MSG_0072）
     *
     * @param torokuId  登録ID
     * @param file      アップロードファイル
     * @param loginUser ログインユーザ
     * @return 追加した添付行
     */
    @Transactional(rollbackFor = Exception.class)
    public TenpuRowForm add(BigDecimal torokuId, MultipartFile file, String loginUser) {
        // 親の作業予定が他ユーザに削除されていないことをロック下で確認する。
        if (repository.lockAndGetLastupdate(torokuId).isEmpty()) {
            throw new ExclusiveControlException(Mcm3005uService.MSG_STALE);
        }
        if (file == null || file.isEmpty()) {
            throw new InputCheckException(MSG_NOT_SELECTED);
        }
        String name = baseName(file.getOriginalFilename());
        if (name.isBlank() || name.length() > MAX_NAME_LENGTH
                || name.matches(".*[\\\\/:*?\"<>|\\p{Cntrl}].*") || name.endsWith(".") || name.endsWith(" ")
                || name.matches("(?i)(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\\..*)?")) {
            throw new InputCheckException(MSG_BAD_NAME);
        }
        if (repository.existsTenpuName(torokuId, name)) {
            throw new InputCheckException(MSG_DUPLICATE);
        }
        Path directory = storage.createDirectory(FOLDER + "/" + torokuId.toPlainString() + "/" + UUID.randomUUID());
        Path target = directory.resolve(name);
        try {
            try (InputStream in = file.getInputStream()) {
                Files.copy(in, target);
            }
            registerRollbackCleanup(target, directory);
            TenpuRowForm row = new TenpuRowForm();
            row.setTenpuId(repository.nextTenpuId());
            row.setTorokuId(torokuId);
            row.setTenpufileNk(name);
            row.setDirectory(storage.relative(target));
            row.setBiko("");
            repository.insertTenpu(row, loginUser);
            // 画面側で後続の備考更新を行う際の排他判定用に、登録直後の値を取得し直す。
            repository.findTenpu(torokuId, row.getTenpuId())
                .ifPresent(saved -> row.setLastupdateDtValue(saved.getLastupdateDtValue()));
            return row;
        } catch (IOException ex) {
            discard(target, directory);
            log.error("作業予定添付ファイルの保存に失敗しました: torokuId={}", torokuId, ex);
            throw new McmBusinessException(MSG_SAVE_FAILED, ex);
        } catch (RuntimeException ex) {
            discard(target, directory);
            throw ex;
        }
    }

    /**
     * 添付ファイルを削除する。
     * 【変換元】FileDelButton_Click（MCM_MO_TENPU 行削除 + McmFileUploadUtility.DeleteFile → MSG_0073）
     * 物理ファイルはDB削除のコミット後に削除する（ロールバック時にファイルだけ消えるのを防ぐため）。
     *
     * @param torokuId 登録ID
     * @param tenpuId  添付ID
     */
    @Transactional(rollbackFor = Exception.class)
    public void delete(BigDecimal torokuId, BigDecimal tenpuId) {
        if (repository.lockTenpuAndGetLastupdate(torokuId, tenpuId).isEmpty()) {
            throw new ExclusiveControlException(Mcm3005uService.MSG_STALE);
        }
        TenpuRowForm row = repository.findTenpu(torokuId, tenpuId)
            .orElseThrow(() -> new ExclusiveControlException(Mcm3005uService.MSG_STALE));
        if (repository.deleteTenpu(torokuId, tenpuId) != 1) {
            throw new ExclusiveControlException(Mcm3005uService.MSG_STALE);
        }
        Path file = insideRoot(row.getDirectory(), row.getTenpufileNk());
        if (file == null) {
            // VB版の共有フォルダ等、アップロードルート外のパスはWebサーバから削除しない（任意パス削除の防止）。
            log.warn("アップロードルート外の添付のため物理ファイルは削除しません: tenpuId={}", tenpuId);
            return;
        }
        Runnable cleanup = () -> {
            try {
                Files.deleteIfExists(file);
                Path parent = file.getParent();
                if (parent != null) {
                    try {
                        Files.deleteIfExists(parent);
                    } catch (IOException ex) {
                        log.debug("添付フォルダが空でないため残します: {}", parent);
                    }
                }
            } catch (IOException ex) {
                log.error("添付削除後の物理ファイル削除に失敗しました: tenpuId={}", tenpuId, ex);
            }
        };
        if (TransactionSynchronizationManager.isSynchronizationActive()) {
            TransactionSynchronizationManager.registerSynchronization(new TransactionSynchronization() {
                @Override
                public void afterCommit() {
                    cleanup.run();
                }
            });
        } else {
            cleanup.run();
        }
    }

    /**
     * 添付ファイルの実体を取得する（ファイル名リンク押下時）。
     * 【変換元】MCM_MO_TENPUCPDataGridView_CellContentClick（存在しない場合 MSG_0109E）
     *
     * @param torokuId 登録ID
     * @param tenpuId  添付ID
     * @return 添付行と実ファイルパス
     * @throws McmBusinessException ファイルが存在しない・保存先がアップロードルート外の場合
     */
    @Transactional(readOnly = true)
    public AttachmentFile resolve(BigDecimal torokuId, BigDecimal tenpuId) {
        TenpuRowForm row = repository.findTenpu(torokuId, tenpuId)
            .orElseThrow(() -> new McmBusinessException(MSG_NOT_FOUND));
        return new AttachmentFile(row, storage.resolve(row.getDirectory(), row.getTenpufileNk()));
    }

    /** 添付行と実ファイルの組。 */
    public record AttachmentFile(TenpuRowForm row, Path path) { }

    /** DIRECTORY（ファイルのフルパス）がアップロードルート配下なら正規化したパスを返す。 */
    private Path insideRoot(String directory, String fileName) {
        try { return storage.resolve(directory, fileName); }
        catch (McmBusinessException ex) {
            log.warn("添付削除の実ファイルを確認できないため物理ファイルは保管します: {}", ex.getMessage());
            return null;
        }
    }

    /** ブラウザによってはフルパスが送られるため、ファイル名部分だけを取り出す。 */
    private static String baseName(String original) {
        if (original == null) {
            return "";
        }
        String name = original.replace('\\', '/');
        int slash = name.lastIndexOf('/');
        return (slash >= 0 ? name.substring(slash + 1) : name).trim();
    }

    private static void registerRollbackCleanup(Path target, Path directory) {
        if (!TransactionSynchronizationManager.isSynchronizationActive()) {
            return;
        }
        TransactionSynchronizationManager.registerSynchronization(new TransactionSynchronization() {
            @Override
            public void afterCompletion(int status) {
                if (status != STATUS_COMMITTED) {
                    discard(target, directory);
                }
            }
        });
    }

    private static void discard(Path target, Path directory) {
        try {
            Files.deleteIfExists(target);
            Files.deleteIfExists(directory);
        } catch (IOException ex) {
            log.error("アップロード取消時のファイル削除に失敗しました: {}", target, ex);
        }
    }
}
