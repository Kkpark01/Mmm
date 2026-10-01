package com.daifuku.mcm.service;

import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.*;
import java.util.UUID;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.transaction.support.TransactionSynchronization;
import org.springframework.transaction.support.TransactionSynchronizationManager;
import org.springframework.web.multipart.MultipartFile;
import com.daifuku.mcm.form.Mcm2003uForm;
import com.daifuku.mcm.repository.Mcm2003uRepository;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;

/** VBの見積添付をWebのアップロード／ダウンロードへ置き換える。 */
@Service
public class Mcm2003uAttachmentService {
    private static final org.slf4j.Logger log = org.slf4j.LoggerFactory.getLogger(Mcm2003uAttachmentService.class);
    @Autowired private Mcm2003uRepository repo;
    @Autowired private Mcm2003uService estimates;
    @Autowired private com.daifuku.mcm.common.FileStorageService storage;
    @Transactional(rollbackFor=Exception.class)
    public void add(Mcm2003uForm form, MultipartFile file, String user) throws IOException {
        estimates.requireEditable(form);
        if(file==null || file.isEmpty())throw new IllegalStateException("添付ファイルを選択してください。");
        String name=file.getOriginalFilename();
        if(name==null || name.isBlank() || name.length()>180 || name.matches(".*[\\\\/:*?\"<>|\\p{Cntrl}].*") || name.endsWith(".") || name.endsWith(" ")
                || name.matches("(?i)(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\\..*)?"))
            throw new IllegalStateException("ファイル名を変更して添付してください。");
        // requireEditableで取得した見積ロック内で、保存前に同じ見積の登録済み添付を確認する。
        if(repo.findTenpuRows(form.getUmKihonMitsumoriId()).stream()
                .anyMatch(row -> name.equalsIgnoreCase(row.getTenpufileNk())))
            throw new IllegalStateException("同一名のファイルが既にアップロードされている為、ファイル追加することは出来ません。"); // MSG_0138E
        Path dir=storage.createDirectory("店舗/" + form.getUmKihonMitsumoriId().stripTrailingZeros().toPlainString() + "/" + UUID.randomUUID());
        Path target=dir.resolve(name);
        try {
            try(var in=file.getInputStream()) { Files.copy(in,target); }
            if(TransactionSynchronizationManager.isSynchronizationActive()) TransactionSynchronizationManager.registerSynchronization(new TransactionSynchronization(){
                @Override public void afterCompletion(int status){
                    if(status!=STATUS_COMMITTED)try{discardUpload(target,dir);}
                    catch(IOException | RuntimeException ex){log.error("添付ロールバック後のファイル削除に失敗しました: {}",target,ex);}
                }
            });
            repo.insertAttachment(form.getUmKihonMitsumoriId(),name,storage.relative(dir),user);
        } catch(IOException | RuntimeException ex) {
            try { discardUpload(target,dir); }
            catch(IOException | RuntimeException cleanupError) { ex.addSuppressed(cleanupError); }
            if (ex instanceof IOException) throw new com.daifuku.mcm.exception.McmBusinessException(
                "添付ファイルを保存できませんでした。管理者に確認してください。", ex);
            throw ex;
        }
    }
    /** 今回のアップロードと専用フォルダだけを削除する。既存添付には触れない。 */
    private static void discardUpload(Path target,Path dir) throws IOException {
        Files.deleteIfExists(target);
        Files.deleteIfExists(dir);
    }
    @Transactional
    public void remove(Mcm2003uForm form,BigDecimal attachment,String user) {
        estimates.requireEditable(form);
        row(form.getUmKihonMitsumoriId(),attachment);
        repo.deleteAttachment(form.getUmKihonMitsumoriId(),attachment);
        // 物理ファイルは保管し、誤削除時に復元できるようにする。
    }
    public Mcm2003uForm.TenpuRowForm row(BigDecimal estimate,BigDecimal attachment) {
        return repo.findTenpuRows(estimate).stream().filter(r->same(r.getUmTenpuId(),attachment)).findFirst()
            .orElseThrow(()->new IllegalStateException("添付資料が見つかりません。"));
    }
    public Path file(Mcm2003uForm.TenpuRowForm row) throws IOException {
        return storage.resolve(row.getDirectory(), row.getTenpufileNk());
    }
}
