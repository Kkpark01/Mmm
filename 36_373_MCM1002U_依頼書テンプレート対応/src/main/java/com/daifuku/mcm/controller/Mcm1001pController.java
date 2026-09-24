/**
 * 【変換元】MCM1001P 見積依頼書（帳票出力）
 *   VB.NETでは画面から直接Excel Interopを呼び出していたが、
 *   Java版ではControllerからServiceを呼び出しHTTPレスポンスとして返却する。
 *
 *   元コード: McmXXXXxScreen.vb 内の帳票出力ボタンクリックイベント
 */
package com.daifuku.mcm.controller;

import com.daifuku.mcm.dto.Mcm1001pDeliveryDto;
import com.daifuku.mcm.service.Mcm1001pExcelService;
import jakarta.servlet.http.HttpServletResponse;
import jakarta.servlet.http.HttpSession;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.io.IOException;
import java.io.OutputStream;

/**
 * MCM1001P 見積依頼書 帳票出力コントローラー
 * <p>
 * セッションに格納された帳票出力用データ（Mcm1001pDeliveryDto）を取得し、
 * Excel帳票を生成してダウンロードレスポンスとして返却する。
 * </p>
 *
 * 【変換元】MCM1001P 帳票出力画面
 *   元コード: VB.NETでは呼び出し元画面のボタンイベントから
 *            Mcm1001pExcel.CreateExcel() を直接呼び出していた
 */
@Controller
@RequestMapping("/mcm1001p")
@RequiredArgsConstructor
@Slf4j
public class Mcm1001pController {

    /** Excel帳票出力サービス */
    private final Mcm1001pExcelService excelService;

    /** セッションキー：帳票出力データ */
    private static final String SESSION_KEY_DELIVERY = "mcm1001pDelivery";

    /**
     * 見積依頼書Excel帳票をダウンロードする。
     *
     * 【変換元】MCM1001P - 帳票出力処理
     *   元コード: Mcm1001pExcel.CreateExcel(delivery)
     *   変換: VB.NETのファイル保存ダイアログ → HTTPレスポンスでダウンロード
     *
     * @param session HttpSession（帳票出力データを取得）
     * @param response HttpServletResponse（Excelファイルをストリーム出力）
     * @param redirectAttributes リダイレクト時のメッセージ用
     * @return エラー時のリダイレクト先（正常時はvoid相当でnull返却）
     */
    @GetMapping("/download")
    public String download(HttpSession session,
                           HttpServletResponse response,
                           RedirectAttributes redirectAttributes) {
        log.info("見積依頼書帳票ダウンロード開始");

        // セッションから帳票出力データを取得
        Mcm1001pDeliveryDto delivery =
                (Mcm1001pDeliveryDto) session.getAttribute(SESSION_KEY_DELIVERY);

        if (delivery == null) {
            log.warn("セッションに帳票出力データが存在しません");
            redirectAttributes.addFlashAttribute("error",
                    "帳票出力データが見つかりません。見積画面から再度操作してください。");
            return "redirect:/mcm1002u";
        }

        try {
            // Excel帳票生成
            byte[] excelData = excelService.generateReport(delivery);

            // レスポンスヘッダー設定
            String fileName = excelService.getOutputFileName(delivery);
            response.setContentType(
                    "application/vnd.ms-excel");
            response.setHeader("Content-Disposition",
                    "attachment; filename*=UTF-8''" + fileName);
            response.setContentLength(excelData.length);

            // Excelファイルをレスポンスに書き込み
            try (OutputStream os = response.getOutputStream()) {
                os.write(excelData);
                os.flush();
            }

            log.info("見積依頼書帳票ダウンロード完了: {}", fileName);

            // セッションから帳票データをクリア（メモリ解放）
            session.removeAttribute(SESSION_KEY_DELIVERY);

            return null; // ストリーム出力済みのため遷移なし

        } catch (IOException e) {
            log.error("見積依頼書帳票ダウンロードでエラー発生", e);
            redirectAttributes.addFlashAttribute("error",
                    "帳票出力中にエラーが発生しました。");
            return "redirect:/mcm1002u";
        }
    }
}
