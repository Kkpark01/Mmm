package com.daifuku.mcm.service;

import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.entity.McmTkTenpuEntity;
import com.daifuku.mcm.exception.InputCheckException;
import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.repository.Mcm1005uRepository;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.*;
import java.time.format.DateTimeFormatter;
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
    public void delete(BigDecimal contractId, BigDecimal attachmentId, String user) {
        var row = find(contractId, attachmentId, user);
        Path target;
        try { target = file(row); }
        catch (IOException ex) { throw new McmBusinessException("添付ファイルを確認できません。管理者に確認してください。"); }
        contracts.deleteTenpu(attachmentId, contractId, user);
        Runnable cleanup = () -> {
            try { Files.deleteIfExists(target); }
            catch (IOException ex) { log.error("添付削除後のファイル削除に失敗: {}", attachmentId, ex); }
        };
        if (TransactionSynchronizationManager.isSynchronizationActive()) {
            TransactionSynchronizationManager.registerSynchronization(new TransactionSynchronization() {
                @Override public void afterCommit() { cleanup.run(); }
            });
        } else cleanup.run();
    }

    @Transactional(rollbackFor = Exception.class)
    public void add(BigDecimal contractId, BigDecimal periodId, MultipartFile file, String user) {
        if (!contracts.canUpdate(user)) throw new McmBusinessException("権限がないため実行できません。");
        var contract = repository.lockKeiyaku(contractId);
        if (!contracts.canEditWithUnlock(contract)) throw new InputCheckException("現在の状態では追加できません。");
        var period = repository.findKikanByKeiyakuId(contractId).stream()
                .filter(p -> p.getTkKikanId().compareTo(periodId) == 0).findFirst()
                .orElseThrow(() -> new InputCheckException("対象期間が見つかりません。画面を開き直してください。"));
        contracts.checkAttachmentEditable(contract, periodId);
        if (file == null || file.isEmpty()) throw new InputCheckException("添付ファイルを選択してください。");
        String name = file.getOriginalFilename();
        if (name == null || name.isBlank() || name.length() > 180
                || name.matches(".*[\\\\/:*?\"<>|\\p{Cntrl}].*") || name.endsWith(".") || name.endsWith(" ")
                || name.matches("(?i)(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\\..*)?"))
            throw new InputCheckException("ファイル名を変更して添付してください。");
        if (repository.findTenpuByKikanId(periodId).stream().anyMatch(a -> name.equalsIgnoreCase(a.getTenpufileNk())))
            throw new InputCheckException("同一名のファイルが既にアップロードされている為、ファイル追加することは出来ません。");
        Path root = Paths.get(uploadPath).toAbsolutePath().normalize();
        // VB FileUpButton_Click: 契約ID_取引先名 / 期間ID_期間開始日 / 契約No_元ファイル名。
        // 名称はVBの1003U引継ぎ元ビュー、開始日は対象契約に所属する指定期間から取得する。
        if (period.getKaisiDt() == null || period.getKaisiDt().getYear() < 1
                || period.getKaisiDt().getYear() > 9999)
            throw new InputCheckException("契約期間の開始日を確認してください。");
        String supplier = pathElement(repository.findAttachmentSupplierName(contractId), "取引先名");
        String contractNo = pathElement(contract.getKeiyakuNo(), "契約No");
        String contractFolder = pathElement(idText(contractId, "取引先契約ID") + "_" + supplier, "契約フォルダ名");
        String periodFolder = idText(periodId, "取引先契約期間ID") + "_"
                + period.getKaisiDt().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        String storedName = pathElement(contractNo + "_" + name, "保存ファイル名");
        Path directory = root.resolve("取引先").resolve(contractFolder).resolve(periodFolder);
        Path target = directory.resolve(storedName);
        boolean created = false;
        try {
            Files.createDirectories(directory);
            if (!directory.toRealPath().startsWith(root.toRealPath())) throw new IOException("Attachment directory escapes upload root");
            // 同名の実ファイルも上書きしない。作成に成功したファイルだけを取消対象にする。
            try (var output = Files.newOutputStream(target, StandardOpenOption.CREATE_NEW, StandardOpenOption.WRITE)) {
                created = true;
                try (var input = file.getInputStream()) { input.transferTo(output); }
            }
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
        } catch (FileAlreadyExistsException ex) {
            throw new InputCheckException("同一名のファイルが保存先に存在する為、ファイル追加することは出来ません。");
        } catch (IOException ex) {
            if (created) discard(target, directory);
            log.error("添付ファイル保存失敗", ex);
            throw new McmBusinessException("添付ファイルを保存できませんでした。保存先を確認してください。");
        } catch (RuntimeException ex) {
            if (created) discard(target, directory);
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

    private static String idText(BigDecimal value, String label) {
        try {
            if (value == null || value.signum() <= 0) throw new ArithmeticException();
            return value.toBigIntegerExact().toString();
        } catch (ArithmeticException ex) {
            throw new InputCheckException(label + "を確認してください。");
        }
    }

    private static String pathElement(String value, String label) {
        // 使用不可文字を勝手に置換するとVBの名称と変わるため、案内して中止する。
        if (value == null || value.isBlank() || value.length() > 255
                || value.matches("(?s).*[\\\\/:*?\"<>|\\p{Cntrl}].*") || value.endsWith(".") || value.endsWith(" ")
                || value.matches("(?i)(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\\..*)?"))
            throw new InputCheckException(label + "を保存先に使用できません。登録内容を確認してください。");
        return value;
    }

    private static void discard(Path target, Path directory) {
        try {
            Files.deleteIfExists(target);
            // 複数添付が同じ期間フォルダを使用する。別の添付があるフォルダは残す。
            try { Files.deleteIfExists(directory); }
            catch (DirectoryNotEmptyException ignored) { /* 他の添付を維持する。 */ }
        } catch (IOException ex) { log.error("アップロード取消時のファイル削除失敗: {}", target, ex); }
    }
}
