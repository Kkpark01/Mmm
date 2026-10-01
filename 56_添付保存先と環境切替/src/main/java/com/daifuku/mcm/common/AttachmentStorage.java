package com.daifuku.mcm.common;

import com.daifuku.mcm.exception.AttachmentStorageException;
import static com.daifuku.mcm.exception.AttachmentStorageException.Reason.*;
import java.io.IOException;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.context.properties.bind.Bindable;
import org.springframework.boot.context.properties.bind.Binder;
import org.springframework.core.env.Environment;
import org.springframework.stereotype.Service;

/** 環境の保存ルートとDBの相対パス／旧絶対パスを解決する共通境界。 */
@Service
public class AttachmentStorage {
    private static final Logger log = LoggerFactory.getLogger(AttachmentStorage.class);
    private final Environment environment;
    public AttachmentStorage(Environment environment) { this.environment = environment; }

    private Path root() {
        String value = environment.getProperty("mcm.file.upload-path", "");
        if (value.isBlank()) throw failure(UNCONFIGURED, null, null);
        try {
            Path path = Path.of(value.replace('\\', '/'));
            if (!path.isAbsolute()) throw failure(INVALID_PATH, value, null);
            return path.normalize();
        } catch (InvalidPathException e) { throw failure(INVALID_PATH, value, e); }
    }

    private Path realRoot() {
        Path root = root();
        try {
            Path real = root.toRealPath();
            if (!Files.isDirectory(real)) throw failure(ROOT_UNAVAILABLE, root.toString(), null);
            return real;
        } catch (AccessDeniedException e) { throw failure(ACCESS_DENIED, root.toString(), e); }
        catch (IOException | SecurityException e) { throw failure(ROOT_UNAVAILABLE, root.toString(), e); }
    }

    /** 呼出元のチェック済みID＋UUIDによる専用フォルダを作る。ルート自体は配備時に用意する。 */
    public Path createDirectory(String relative) {
        Path realRoot = realRoot();
        Path directory = relativePath(relative);
        Path parent = directory;
        try {
            while (!Files.exists(parent, LinkOption.NOFOLLOW_LINKS)) parent = parent.getParent();
            if (!parent.toRealPath().startsWith(realRoot)) throw failure(INVALID_PATH, directory.toString(), null);
            Files.createDirectories(directory);
            if (!directory.toRealPath().startsWith(realRoot)) throw failure(INVALID_PATH, directory.toString(), null);
            return directory;
        } catch (AccessDeniedException e) { throw failure(ACCESS_DENIED, directory.toString(), e); }
        catch (IOException | SecurityException e) { throw failure(SAVE_FAILED, directory.toString(), e); }
    }

    /** DIRECTORYは画面によりフォルダまたはフルファイル。ファイル名を二重連結しない。 */
    public Path resolve(String stored, String fileName) {
        Path realRoot = realRoot();
        if (!validName(fileName) || stored == null || stored.isBlank()) throw failure(INVALID_PATH, stored, null);
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
            if (suffix == null) throw failure(INVALID_PATH, stored, null);
            portable = suffix;
        }
        if (portable.endsWith("/")) portable = portable.substring(0, portable.length() - 1);
        int slash = portable.lastIndexOf('/');
        if (!portable.substring(slash + 1).equalsIgnoreCase(fileName)) portable += "/" + fileName;
        Path path = relativePath(portable);
        try {
            Path real = path.toRealPath();
            if (!real.startsWith(realRoot)) throw failure(INVALID_PATH, path.toString(), null);
            BasicFileAttributes attributes = Files.readAttributes(real, BasicFileAttributes.class);
            if (!attributes.isRegularFile()) throw failure(NOT_FOUND, path.toString(), null);
            // Files.isReadableによる事前判定だけでは共有フォルダのACLを正確に判断できない。
            try (var ignored = Files.newInputStream(real)) { /* 実際に読めることを確認 */ }
            return real;
        } catch (NoSuchFileException e) { throw failure(NOT_FOUND, path.toString(), e); }
        catch (AccessDeniedException e) { throw failure(ACCESS_DENIED, path.toString(), e); }
        catch (IOException | SecurityException e) { throw failure(ROOT_UNAVAILABLE, path.toString(), e); }
    }

    public String relative(Path path) {
        Path root = root();
        Path normalized = path.toAbsolutePath().normalize();
        if (!normalized.startsWith(root)) throw failure(INVALID_PATH, path.toString(), null);
        return root.relativize(normalized).toString().replace('\\', '/');
    }

    private Path relativePath(String relative) {
        if (relative == null || relative.isBlank() || relative.indexOf(':') >= 0 || relative.indexOf('\\') >= 0
                || relative.startsWith("/") || relative.matches(".*[\\p{Cntrl}].*"))
            throw failure(INVALID_PATH, relative, null);
        for (String part : relative.split("/"))
            if (part.equals("..") || part.equals(".")) throw failure(INVALID_PATH, relative, null);
        try {
            Path result = root().resolve(relative).normalize();
            if (!result.startsWith(root()) || result.equals(root())) throw failure(INVALID_PATH, relative, null);
            return result;
        } catch (InvalidPathException e) { throw failure(INVALID_PATH, relative, e); }
    }

    private boolean validName(String name) {
        return name != null && !name.isBlank() && !name.equals(".") && !name.equals("..")
                && !name.matches(".*[\\\\/:*?\"<>|\\p{Cntrl}].*") && !name.endsWith(".") && !name.endsWith(" ");
    }

    private AttachmentStorageException failure(AttachmentStorageException.Reason reason, String path, Throwable cause) {
        log.warn("添付保存先エラー reason={} path={}", reason, path, cause);
        return new AttachmentStorageException(reason, cause);
    }
}
