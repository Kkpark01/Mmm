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
import org.springframework.core.env.Environment;
import org.springframework.boot.context.properties.bind.Bindable;
import org.springframework.boot.context.properties.bind.Binder;
import com.daifuku.mcm.exception.McmBusinessException;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@Service
public class FileStorageService {

    private static final Logger logger = LoggerFactory.getLogger(FileStorageService.class);

    private final Environment environment;

    public FileStorageService(Environment environment) {
        this.environment = environment;
    }

    @org.springframework.beans.factory.annotation.Value("${mcm.file.output-path:/opt/mcm/output}")
    private String outputPath;

    /**
     * 【変換元】CPFileManager.vb - CopyFile() / アップロード処理
     *   Web版ではMultipartFileのアップロード保存に変換
     */
    public Path store(MultipartFile file, String subDir) throws IOException {
        Path dir = createDirectory(subDir);
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

    // 添付の保存ルート・相対パス・旧絶対パスの読み替えを共通で扱う。
    private Path root() {
        String value = environment.getProperty("mcm.file.upload-path", "");
        if (value.isBlank()) throw failure(StorageFailure.UNCONFIGURED, null, null);
        try {
            Path path = Path.of(value.replace('\\', '/'));
            if (!path.isAbsolute()) throw failure(StorageFailure.INVALID_PATH, value, null);
            return path.normalize();
        } catch (InvalidPathException e) { throw failure(StorageFailure.INVALID_PATH, value, e); }
    }

    private Path realRoot() {
        Path root = root();
        try {
            Path real = root.toRealPath();
            if (!Files.isDirectory(real)) throw failure(StorageFailure.ROOT_UNAVAILABLE, root.toString(), null);
            return real;
        } catch (AccessDeniedException e) { throw failure(StorageFailure.ACCESS_DENIED, root.toString(), e); }
        catch (IOException | SecurityException e) { throw failure(StorageFailure.ROOT_UNAVAILABLE, root.toString(), e); }
    }

    /** 呼出元のチェック済みID＋UUIDによる専用フォルダを作る。ルート自体は配備時に用意する。 */
    public Path createDirectory(String relative) {
        Path realRoot = realRoot();
        Path directory = relativePath(relative);
        Path parent = directory;
        try {
            while (!Files.exists(parent, LinkOption.NOFOLLOW_LINKS)) parent = parent.getParent();
            if (!parent.toRealPath().startsWith(realRoot)) throw failure(StorageFailure.INVALID_PATH, directory.toString(), null);
            Files.createDirectories(directory);
            if (!directory.toRealPath().startsWith(realRoot)) throw failure(StorageFailure.INVALID_PATH, directory.toString(), null);
            return directory;
        } catch (AccessDeniedException e) { throw failure(StorageFailure.ACCESS_DENIED, directory.toString(), e); }
        catch (IOException | SecurityException e) { throw failure(StorageFailure.SAVE_FAILED, directory.toString(), e); }
    }

    /** DIRECTORYは画面によりフォルダまたはフルファイル。ファイル名を二重連結しない。 */
    public Path resolve(String stored, String fileName) {
        Path realRoot = realRoot();
        if (!validName(fileName) || stored == null || stored.isBlank()) throw failure(StorageFailure.INVALID_PATH, stored, null);
        String portable = stored.replace('\\', '/');
        boolean absolute = portable.startsWith("/") || portable.matches("^[A-Za-z]:/.*");
        if (absolute) {
            List<String> roots = new ArrayList<>();
            roots.add(root().toString());
            roots.addAll(Binder.get(environment).bind("mcm.file.legacy-roots", Bindable.listOf(String.class)).orElse(List.of()));
            roots.sort(Comparator.comparingInt(String::length).reversed());
            String suffix = null;
            for (String alias : roots) {
                if (alias == null || alias.isBlank()) continue;
                String prefix = alias.replace('\\', '/').replaceAll("/+$", "");
                boolean windows = prefix.matches("^[A-Za-z]:/.*") || prefix.startsWith("//");
                String p = windows ? portable.toLowerCase(Locale.ROOT) : portable;
                String a = windows ? prefix.toLowerCase(Locale.ROOT) : prefix;
                if (p.startsWith(a + "/")) { suffix = portable.substring(prefix.length() + 1); break; }
            }
            if (suffix == null) throw failure(StorageFailure.INVALID_PATH, stored, null);
            portable = suffix;
        }
        if (portable.endsWith("/")) portable = portable.substring(0, portable.length() - 1);
        int slash = portable.lastIndexOf('/');
        if (!portable.substring(slash + 1).equalsIgnoreCase(fileName)) portable += "/" + fileName;
        Path path = relativePath(portable);
        try {
            Path real = path.toRealPath();
            if (!real.startsWith(realRoot)) throw failure(StorageFailure.INVALID_PATH, path.toString(), null);
            BasicFileAttributes attributes = Files.readAttributes(real, BasicFileAttributes.class);
            if (!attributes.isRegularFile()) throw failure(StorageFailure.NOT_FOUND, path.toString(), null);
            // Files.isReadableによる事前判定だけでは共有フォルダのACLを正確に判断できない。
            try (var ignored = Files.newInputStream(real)) { /* 実際に読めることを確認 */ }
            return real;
        } catch (NoSuchFileException e) { throw failure(StorageFailure.NOT_FOUND, path.toString(), e); }
        catch (AccessDeniedException e) { throw failure(StorageFailure.ACCESS_DENIED, path.toString(), e); }
        catch (IOException | SecurityException e) { throw failure(StorageFailure.ROOT_UNAVAILABLE, path.toString(), e); }
    }

    public String relative(Path path) {
        Path root = root();
        Path normalized = path.toAbsolutePath().normalize();
        if (!normalized.startsWith(root)) throw failure(StorageFailure.INVALID_PATH, path.toString(), null);
        return root.relativize(normalized).toString().replace('\\', '/');
    }

    private Path relativePath(String relative) {
        if (relative == null || relative.isBlank() || relative.indexOf(':') >= 0 || relative.indexOf('\\') >= 0
                || relative.startsWith("/") || relative.matches(".*[\\p{Cntrl}].*"))
            throw failure(StorageFailure.INVALID_PATH, relative, null);
        for (String part : relative.split("/"))
            if (part.equals("..") || part.equals(".")) throw failure(StorageFailure.INVALID_PATH, relative, null);
        try {
            Path result = root().resolve(relative).normalize();
            if (!result.startsWith(root()) || result.equals(root())) throw failure(StorageFailure.INVALID_PATH, relative, null);
            return result;
        } catch (InvalidPathException e) { throw failure(StorageFailure.INVALID_PATH, relative, e); }
    }

    private boolean validName(String name) {
        return name != null && !name.isBlank() && !name.equals(".") && !name.equals("..")
                && !name.matches(".*[\\\\/:*?\"<>|\\p{Cntrl}].*") && !name.endsWith(".") && !name.endsWith(" ");
    }

    private McmBusinessException failure(StorageFailure reason, String path, Throwable cause) {
        logger.warn("添付保存先エラー reason={} path={}", reason, path, cause);
        return new McmBusinessException(reason.message, cause);
    }

    /** 利用者向け文言とログ用の原因区分。専用の例外クラスは作らない。 */
    private enum StorageFailure {
        UNCONFIGURED("添付資料の保存先が未設定です。管理者に確認してください。"),
        ROOT_UNAVAILABLE("添付資料の保存先に接続できません。管理者に確認してください。"),
        ACCESS_DENIED("添付資料の保存先にアクセスできません。管理者に確認してください。"),
        NOT_FOUND("添付資料の実ファイルが見つかりません。管理者に確認してください。"),
        INVALID_PATH("添付資料の保存先を確認してください。"),
        SAVE_FAILED("添付ファイルを保存できませんでした。管理者に確認してください。");

        private final String message;
        StorageFailure(String message) { this.message = message; }
    }
}
