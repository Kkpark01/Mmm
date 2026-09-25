package com.daifuku.mcm.service;

import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.entity.McmTkTenpuEntity;
import com.daifuku.mcm.exception.InputCheckException;
import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.repository.Mcm1005uRepository;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.*;
import java.util.UUID;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.transaction.support.TransactionSynchronization;
import org.springframework.transaction.support.TransactionSynchronizationManager;
import org.springframework.web.multipart.MultipartFile;

/** VB FileUpButton_Click: attachment persistence is independent of contract form edits. */
@Service
public class Mcm1005uAttachmentService {
    @Autowired private Mcm1005uRepository repository;
    @Autowired private Mcm1005uService contracts;
    @Autowired private Mcm2004uService permissions;
    @Value("${mcm.file.upload-path:/opt/mcm/upload}") private String uploadPath;
    private static final org.slf4j.Logger log = org.slf4j.LoggerFactory.getLogger(Mcm1005uAttachmentService.class);

    @Transactional(rollbackFor = Exception.class)
    public void add(BigDecimal contractId, BigDecimal periodId, MultipartFile file, String user) {
        if (!contracts.canUpdate(user)) throw new McmBusinessException("権限がないため実行できません。");
        var contract = repository.lockKeiyaku(contractId);
        if (!contracts.canEditDetails(contract)) throw new InputCheckException("現在の状態では追加できません。");
        if (repository.findKikanByKeiyakuId(contractId).stream().noneMatch(p -> p.getTkKikanId().compareTo(periodId) == 0))
            throw new InputCheckException("対象期間が見つかりません。画面を開き直してください。");
        if (file == null || file.isEmpty()) throw new InputCheckException("添付ファイルを選択してください。");
        String name = file.getOriginalFilename();
        if (name == null || name.isBlank() || name.length() > 180
                || name.matches(".*[\\\\/:*?\"<>|\\p{Cntrl}].*") || name.endsWith(".") || name.endsWith(" ")
                || name.matches("(?i)(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\\..*)?"))
            throw new InputCheckException("ファイル名を変更して添付してください。");
        if (repository.findTenpuByKikanId(periodId).stream().anyMatch(a -> name.equalsIgnoreCase(a.getTenpufileNk())))
            throw new InputCheckException("同一名のファイルが既にアップロードされている為、ファイル追加することは出来ません。");
        Path root = Paths.get(uploadPath).toAbsolutePath().normalize();
        Path directory = root.resolve("torihikisaki").resolve(contractId.toPlainString()).resolve(periodId.toPlainString()).resolve(UUID.randomUUID().toString());
        Path target = directory.resolve(name);
        try {
            Files.createDirectories(directory);
            if (!directory.toRealPath().startsWith(root.toRealPath())) throw new IOException("Attachment directory escapes upload root");
            try (var in = file.getInputStream()) { Files.copy(in, target); }
            if (TransactionSynchronizationManager.isSynchronizationActive()) {
                TransactionSynchronizationManager.registerSynchronization(new TransactionSynchronization() {
                    @Override public void afterCompletion(int status) {
                        if (status != STATUS_COMMITTED) discard(target, directory);
                    }
                });
            }
            var row = new McmTkTenpuEntity();
            row.setTkTenpuId(repository.nextTenpuId()); row.setTkKikanId(periodId);
            row.setTenpufileNk(name); row.setDirectory(target.toString());
            row.setShoninjotai(McmConstants.SHONINJOTAI_SAKUSEICHU_CD);
            repository.insertTenpu(row, user);
        } catch (IOException ex) {
            discard(target, directory);
            log.error("添付ファイル保存失敗", ex);
            throw new McmBusinessException("添付ファイルを保存できませんでした。保存先を確認してください。");
        } catch (RuntimeException ex) {
            discard(target, directory);
            throw ex;
        }
    }

    @Transactional(readOnly = true)
    public McmTkTenpuEntity find(BigDecimal contractId, BigDecimal attachmentId, String user) {
        if (!java.util.Set.of("1", "2").contains(String.valueOf(permissions.getAuthority(user, "MCM1005U"))))
            throw new McmBusinessException("権限がないため実行できません。");
        return repository.findKikanByKeiyakuId(contractId).stream()
                .flatMap(p -> repository.findTenpuByKikanId(p.getTkKikanId()).stream())
                .filter(a -> a.getTkTenpuId().compareTo(attachmentId) == 0).findFirst()
                .orElseThrow(() -> new InputCheckException("添付資料が見つかりません。"));
    }

    public Path file(McmTkTenpuEntity row) throws IOException {
        Path root = Paths.get(uploadPath).toAbsolutePath().normalize().toRealPath();
        // VB DIRECTORY contains the full file path, not just its parent directory.
        Path path = Paths.get(row.getDirectory()).toAbsolutePath().normalize();
        if (!path.startsWith(root) || !Files.isRegularFile(path) || !path.toRealPath().startsWith(root))
            throw new InputCheckException("添付資料の保存先を確認してください。");
        return path;
    }

    private static void discard(Path target, Path directory) {
        try { Files.deleteIfExists(target); Files.deleteIfExists(directory); }
        catch (IOException ex) { log.error("アップロード取消時のファイル削除失敗: {}", target, ex); }
    }
}
