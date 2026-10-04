package com.daifuku.mcm;

import java.sql.DriverManager;
import java.sql.DriverPropertyInfo;
import java.sql.SQLException;
import java.util.Arrays;
import java.util.Properties;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.core.env.Environment;

@SpringBootApplication
public class McmApplication {

    private static final Logger log = LoggerFactory.getLogger(McmApplication.class);

    public static void main(String[] args) {
        SpringApplication.run(McmApplication.class, args);
    }

    /** 設定上の接続先を起動時に表示する。DB接続や添付保存先へのアクセスは行わない。 */
    @Bean
    public ApplicationRunner mcmEnvironmentInfo(Environment environment) {
        return args -> {
            String[] profiles = environment.getActiveProfiles();
            if (profiles.length == 0) {
                profiles = environment.getDefaultProfiles();
            }
            String url = environment.getProperty("spring.datasource.hikari.jdbc-url",
                    environment.getProperty("spring.datasource.url"));
            log.info("[MCM起動設定] 使用環境={}", String.join(",", profiles));
            log.info("[MCM起動設定] DB接続先（設定）={}", databaseTarget(url));
            log.info("[MCM起動設定] 添付保存先={}",
                    logValue(environment.getProperty("mcm.file.upload-path")));
            if (Arrays.asList(profiles).contains("local")) {
                log.info("[MCM起動設定] localの添付はこのPCに保存されます。"
                        + "DB接続先は上記の設定を確認してください。共有添付の確認にはdevを指定してください。");
            }
        };
    }

    /** JDBCドライバーでURLを解析し、サーバー・DB名だけを表示する。認証情報やURL全体は出さない。 */
    private static String databaseTarget(String url) {
        if (url == null || url.isBlank()) {
            return "未設定";
        }
        try {
            DriverPropertyInfo[] properties = DriverManager.getDriver(url)
                    .getPropertyInfo(url, new Properties());
            return "サーバー=" + property(properties, "serverName")
                    + " / インスタンス=" + property(properties, "instanceName")
                    + " / ポート=" + property(properties, "portNumber")
                    + " / DB=" + property(properties, "databaseName");
        } catch (SQLException | RuntimeException ex) {
            // 解析例外のメッセージにはURL・認証情報が含まれる可能性があるためログへ出さない。
            return "表示できません（DB接続設定を確認してください）";
        }
    }

    private static String property(DriverPropertyInfo[] properties, String name) {
        for (DriverPropertyInfo property : properties) {
            if (name.equalsIgnoreCase(property.name)) {
                return logValue(property.value);
            }
        }
        return "未指定";
    }

    private static String logValue(String value) {
        return value == null || value.isBlank() ? "未指定"
                : value.replaceAll("[\\p{Cntrl}]", " ");
    }
}
