package com.daifuku.mcm.service;

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
import com.daifuku.mcm.form.Mcm2003uForm;
import com.daifuku.mcm.repository.Mcm2003uRepository;
import static com.daifuku.mcm.common.CustomerScreenSupport.*;

/** VBの見積添付をWebのアップロード／ダウンロードへ置き換える。 */
@Service
public class Mcm2003uAttachmentService {
    @Autowired private Mcm2003uRepository repo;
    @Autowired private Mcm2003uService estimates;
    @Value("${mcm.file.upload-path:/opt/mcm/upload}") private String uploadPath;
    @Transactional(rollbackFor=Exception.class)
    public void add(Mcm2003uForm form, MultipartFile file, String user) throws IOException {
        estimates.requireEditable(form);
        if(file==null || file.isEmpty())throw new IllegalStateException("添付ファイルを選択してください。");
        String name=file.getOriginalFilename();
        if(name==null || name.isBlank() || name.length()>180 || name.matches(".*[\\\\/:*?\"<>|\\p{Cntrl}].*") || name.endsWith(".") || name.endsWith(" ")
                || name.matches("(?i)(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\\..*)?"))
            throw new IllegalStateException("ファイル名を変更して添付してください。");
        Path root=Paths.get(uploadPath).toAbsolutePath().normalize();
        Path dir=root.resolve("estimates").resolve(form.getUmKihonMitsumoriId().stripTrailingZeros().toPlainString()).resolve(UUID.randomUUID().toString());
        Files.createDirectories(dir);
        if(!dir.toRealPath().startsWith(root.toRealPath()))throw new IllegalStateException("添付資料の保存先を確認してください。");
        Path target=dir.resolve(name);
        try(var in=file.getInputStream()) { Files.copy(in,target); }
        if(TransactionSynchronizationManager.isSynchronizationActive()) TransactionSynchronizationManager.registerSynchronization(new TransactionSynchronization(){
            @Override public void afterCompletion(int status){if(status!=STATUS_COMMITTED)try{Files.deleteIfExists(target);}catch(IOException ignored){}}
        });
        try { repo.insertAttachment(form.getUmKihonMitsumoriId(),name,dir.toString(),user); }
        catch(RuntimeException ex) {Files.deleteIfExists(target);throw ex;}
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
        Path root=Paths.get(uploadPath).toAbsolutePath().normalize().toRealPath();
        Path file=Paths.get(row.getDirectory()).resolve(row.getTenpufileNk()).toAbsolutePath().normalize();
        if(!file.startsWith(root) || !Files.isRegularFile(file) || !file.toRealPath().startsWith(root))
            throw new IllegalStateException("添付資料の保存先を確認してください。");
        return file;
    }
}
