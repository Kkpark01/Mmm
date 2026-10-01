package com.daifuku.mcm.config;

import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.context.config.ConfigDataEnvironmentPostProcessor;
import org.springframework.boot.env.EnvironmentPostProcessor;
import org.springframework.core.Ordered;
import org.springframework.core.env.ConfigurableEnvironment;

/** 本番で検証用の接続設定が暗黙に継承されないよう、Bean/DB初期化前に検査する。 */
public class McmEnvironmentGuard implements EnvironmentPostProcessor, Ordered {
    @Override public int getOrder() { return ConfigDataEnvironmentPostProcessor.ORDER + 1; }

    @Override public void postProcessEnvironment(ConfigurableEnvironment environment, SpringApplication application) {
        var profiles = Arrays.stream(environment.getActiveProfiles())
                .filter(List.of("local", "test", "prod")::contains).toList();
        if (profiles.size() != 1)
            throw new IllegalStateException("local、test、prodの環境プロファイルを1つだけ指定してください。");
        String kind = profiles.get(0);
        if (!kind.equals(environment.getProperty("mcm.runtime.environment")))
            throw new IllegalStateException("選択した環境のapplication.yml設定が読み込まれていません。");
        if (kind.equals("prod")) {
            for (String key : List.of("spring.datasource.url", "spring.datasource.username",
                    "spring.datasource.password", "mcm.file.upload-path")) {
                if (environment.getProperty(key, "").isBlank())
                    throw new IllegalStateException("本番環境の必須設定が未設定です: " + key);
            }
        }
        String root = environment.getProperty("mcm.file.upload-path", "");
        if (!root.isBlank() && !Path.of(root).isAbsolute())
            throw new IllegalStateException("添付保存先は絶対パスまたはUNCで設定してください。");
    }
}
