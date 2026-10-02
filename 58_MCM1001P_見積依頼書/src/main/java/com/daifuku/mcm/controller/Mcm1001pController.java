package com.daifuku.mcm.controller;

import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.dto.Mcm1001pExportState;
import com.daifuku.mcm.service.Mcm1001pExcelService;
import jakarta.servlet.http.HttpSession;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.HashSet;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/** MCM1002Uで発行した取引先別の見積依頼書を取得する。 */
@Controller
@RequestMapping("/mcm1001p")
@RequiredArgsConstructor
@Slf4j
public class Mcm1001pController extends BaseController {
    private final Mcm1001pExcelService excelService;

    @Override protected String getScreenTitle() { return "見積依頼書"; }
    @Override protected String getFunctionId() { return "MCM1001P"; }

    @GetMapping("/download")
    public ResponseEntity<?> download(HttpSession session,
            @RequestParam(name = "state", required = false) String state) {
        var deliveries = Mcm1001pExportState.find(session, state);
        if (deliveries == null) {
            return error(410, "帳票出力データが見つかりません。見積依頼機器選定画面で依頼書を再発行してください。");
        }
        try {
            byte[] data;
            String name;
            String mime;
            if (deliveries.size() == 1) {
                var delivery = deliveries.get(0);
                data = excelService.generateReport(delivery);
                name = excelService.getRawOutputFileName(delivery);
                mime = "application/vnd.ms-excel";
            } else {
                // ブラウザーの複数ダウンロード制限に依存せず、VBの全取引先を出力する。
                var names = new HashSet<String>();
                try (var bytes = new ByteArrayOutputStream();
                     var zip = new ZipOutputStream(bytes, StandardCharsets.UTF_8)) {
                    for (var delivery : deliveries) {
                        String base = excelService.getRawOutputFileName(delivery);
                        String entry = base;
                        for (int suffix = 2; !names.add(entry); suffix++) {
                            entry = base.substring(0, base.length() - 4) + "_" + suffix + ".xls";
                        }
                        zip.putNextEntry(new ZipEntry(entry));
                        zip.write(excelService.generateReport(delivery));
                        zip.closeEntry();
                    }
                    zip.finish();
                    data = bytes.toByteArray();
                }
                name = "見積依頼書.zip";
                mime = "application/zip";
            }
            // 全件の生成成功後にレスポンスを確定し、途中までのファイルを返さない。
            return ResponseEntity.ok()
                    .header(HttpHeaders.CACHE_CONTROL, "no-store")
                    .header(HttpHeaders.CONTENT_DISPOSITION, ContentDisposition.attachment()
                            .filename(name, StandardCharsets.UTF_8).build().toString())
                    .contentType(MediaType.parseMediaType(mime)).contentLength(data.length).body(data);
        } catch (IOException | RuntimeException e) {
            log.error("MCM1001P 見積依頼書の生成に失敗しました", e);
            return error(500, "Excel出力に失敗しました。もう一度「依頼書を取得」を押してください。");
        }
    }

    private ResponseEntity<?> error(int status, String message) {
        return ResponseEntity.status(status).header(HttpHeaders.CACHE_CONTROL, "no-store")
                .body(Map.of("message", message));
    }
}
