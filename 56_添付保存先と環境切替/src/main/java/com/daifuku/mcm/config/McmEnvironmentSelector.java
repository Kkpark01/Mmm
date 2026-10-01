package com.daifuku.mcm.config;

import java.util.List;
import java.util.Map;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.context.config.ConfigDataEnvironmentPostProcessor;
import org.springframework.boot.env.EnvironmentPostProcessor;
import org.springframework.core.Ordered;
import org.springframework.core.env.ConfigurableEnvironment;
import org.springframework.core.env.MapPropertySource;
import org.springframework.core.env.StandardEnvironment;

/** ConfigDataより先に選択し、標準のSpringプロファイルでDBと保存先を同時に読む。 */
public class McmEnvironmentSelector implements EnvironmentPostProcessor, Ordered {
    @Override public int getOrder() { return ConfigDataEnvironmentPostProcessor.ORDER - 1; }

    @Override public void postProcessEnvironment(ConfigurableEnvironment environment, SpringApplication application) {
        // CLI / SPRING_PROFILES_ACTIVE / JVM指定を最優先する。
        if (!environment.getProperty("spring.profiles.active", "").isBlank()
                || environment.getActiveProfiles().length != 0 || !application.getAdditionalProfiles().isEmpty()) return;
        // 別のEnvironmentで標準の外部設定読込だけを行う。実環境のプロファイル／Beanはまだ確定しない。
        // ./config/application.yml、spring.config.additional-location等の優先順位を独自実装しない。
        var preview = new StandardEnvironment();
        for (var source : java.util.stream.StreamSupport.stream(preview.getPropertySources().spliterator(), false).toList())
            preview.getPropertySources().remove(source.getName());
        for (var source : environment.getPropertySources()) preview.getPropertySources().addLast(source);
        preview.setDefaultProfiles(environment.getDefaultProfiles());
        ConfigDataEnvironmentPostProcessor.applyTo(preview, application.getResourceLoader(), null, List.of());
        // 外部YAMLで明示されたプロファイルも自動判定より優先する。
        if (!preview.getProperty("spring.profiles.active", "").isBlank()) return;
        String selected = environment.getProperty("mcm.environment", environment.getProperty("MCM_ENV", ""));
        if (selected.isBlank()) {
            String test = preview.getProperty("mcm.runtime.test-server", "").trim();
            String prod = preview.getProperty("mcm.runtime.prod-server", "").trim();
            if (!test.isBlank() && test.equalsIgnoreCase(prod))
                throw new IllegalStateException("検証・本番のサーバー名を別々に設定してください。");
            // DNS名・URLからは判定しない。COMPUTERNAMEはJavaが動くWindowsホスト名。
            String host = environment.getProperty("COMPUTERNAME", "");
            selected = !prod.isBlank() && prod.equalsIgnoreCase(host) ? "prod"
                    : !test.isBlank() && test.equalsIgnoreCase(host) ? "test" : "local";
        }
        selected = selected.trim().toLowerCase(java.util.Locale.ROOT);
        if (!List.of("local", "test", "prod").contains(selected))
            throw new IllegalStateException("MCM_ENVはlocal、test、prodのいずれかを指定してください。");
        environment.getPropertySources().addLast(new MapPropertySource("mcmEnvironmentSelection",
                Map.of("spring.profiles.active", selected)));
    }

}
