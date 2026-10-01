package com.daifuku.mcm;

import java.nio.file.Path;
import java.util.List;
import org.springframework.beans.factory.config.BeanFactoryPostProcessor;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.core.env.Environment;

@SpringBootApplication
public class McmApplication {

    public static void main(String[] args) {
        SpringApplication.run(McmApplication.class, args);
    }

    /** 外部YAMLも読み込んだ後、DBを含む通常Beanの生成前に配備設定を検査する。 */
    @Bean
    public static BeanFactoryPostProcessor deploymentSettingsValidator(Environment environment) {
        return beanFactory -> {
            String kind = environment.getProperty("mcm.runtime.environment", "");
            if (!List.of("test", "prod").contains(kind)) {
                throw new IllegalStateException("application.ymlの使用環境をtestまたはprodに設定してください。");
            }
            if (kind.equals("prod")) {
                for (String key : List.of("spring.datasource.url", "spring.datasource.username",
                        "spring.datasource.password", "mcm.file.upload-path")) {
                    if (environment.getProperty(key, "").isBlank()) {
                        throw new IllegalStateException("本番環境の必須設定が未設定です: " + key);
                    }
                }
            }
            String root = environment.getProperty("mcm.file.upload-path", "");
            if (!root.isBlank() && !Path.of(root).isAbsolute()) {
                throw new IllegalStateException("添付保存先は絶対パスまたはUNCで設定してください。");
            }
        };
    }
}
