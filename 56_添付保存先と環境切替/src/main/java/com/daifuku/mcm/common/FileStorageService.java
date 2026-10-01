package com.daifuku.mcm.common;

/**
 * 【変換元】CPFileManager.vb (FRAMEWORK/FILE/CPFileManager.vb 約250行)
 *
 * ファイルストレージサービス
 * VB.NETの CPFileManager（ファイル操作ユーティリティ）を
 * Spring サービスに変換。
 *
 * 変換マッピング:
 *   CPFileManager.CopyFile()   → store() / copy()
 *   CPFileManager.DeleteFile() → delete()
 *   CPFileManager.MoveFile()   → move()
 *   CPFileManager.ExistsFile() → exists()
 */

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@Service
public class FileStorageService {

    private static final Logger logger = LoggerFactory.getLogger(FileStorageService.class);

    @org.springframework.beans.factory.annotation.Autowired private com.daifuku.mcm.common.AttachmentStorage storage;

    @org.springframework.beans.factory.annotation.Value("${mcm.file.output-path:/opt/mcm/output}")
    private String outputPath;

    /**
     * 【変換元】CPFileManager.vb - CopyFile() / アップロード処理
     *   Web版ではMultipartFileのアップロード保存に変換
     */
    public Path store(MultipartFile file, String subDir) throws IOException {
        Path dir = storage.createDirectory(subDir);
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"));
        String fileName = timestamp + "_" + file.getOriginalFilename();
        if (fileName.matches(".*[\\\\/:*?\"<>|\\p{Cntrl}].*"))
            throw new com.daifuku.mcm.exception.InputCheckException("ファイル名を変更して添付してください。");
        Path target = dir.resolve(fileName);
        try (var in = file.getInputStream()) {
            Files.copy(in, target, StandardCopyOption.REPLACE_EXISTING);
        }
        logger.info("[ファイル保存] {}", target);
        return target;
    }

    /**
     * 【変換元】CPFileManager.vb - DeleteFile()
     */
    public boolean delete(String filePath) {
        try {
            boolean deleted = Files.deleteIfExists(Paths.get(filePath));
            logger.info("[ファイル削除] {} - {}", filePath, deleted ? "成功" : "ファイルなし");
            return deleted;
        } catch (IOException e) {
            logger.error("[ファイル削除エラー] {}", e.getMessage());
            return false;
        }
    }

    /**
     * 【変換元】CPFileManager.vb - ExistsFile()
     */
    public boolean exists(String filePath) {
        return Files.exists(Paths.get(filePath));
    }

    /**
     * 【変換元】CPFileManager.vb - MoveFile()
     */
    public Path move(String sourcePath, String targetDir) throws IOException {
        Path source = Paths.get(sourcePath);
        Path dir = Paths.get(targetDir);
        Files.createDirectories(dir);
        Path target = dir.resolve(source.getFileName());
        return Files.move(source, target, StandardCopyOption.REPLACE_EXISTING);
    }
}
