/**
 * 【変換元】Mcm1001pExcel.vb
 *   MCM1001P 見積依頼書 - Excel帳票出力サービス
 *   元コード: Public Class Mcm1001pExcel
 *   元ファイル行数: 約753行
 *
 * 変換内容:
 *   - Excel Interop (CPExcelManager) → Apache POI (XSSFWorkbook)
 *   - VB.NETの1ベースインデックス → POIの0ベースインデックスに変換（-1）
 *   - Range.Merge / Borders → CellRangeAddress + RegionUtil
 */
package com.daifuku.mcm.service;

import com.daifuku.mcm.constants.Mcm1001pConstants;
import com.daifuku.mcm.dto.Mcm1001pDeliveryDto;
import com.daifuku.mcm.dto.Mcm1001pHeaderDto;
import com.daifuku.mcm.dto.Mcm1001pViewDto;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.ss.util.WorkbookUtil;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.List;

/**
 * MCM1001P 見積依頼書 Excel帳票出力サービス
 * <p>
 * テンプレートExcelファイルを読み込み、ヘッダー情報と明細データを書き込んで
 * Excel帳票を生成する。複数シート対応（契約時間帯ごとにシートを分割）。
 * </p>
 *
 * 【変換元】Mcm1001pExcel.vb - 全メソッド
 *   元コード: CPExcelManager による Excel Interop 操作
 */
@Service
@Slf4j
public class Mcm1001pExcelService {

    /**
     * Excel帳票を生成し、バイト配列で返却する。
     *
     * 【変換元】Mcm1001pExcel.vb - CreateExcel()
     *   元コード: objExcel = New CPExcelManager() / objExcel.BookOpen(templatePath)
     *
     * @param delivery 帳票出力用データ（ヘッダー一覧＋明細一覧）
     * @return Excel帳票のバイト配列
     * @throws IOException テンプレート読込みまたはExcel生成でエラーが発生した場合
     */
    public byte[] generateReport(Mcm1001pDeliveryDto delivery) throws IOException {
        log.info("見積依頼書Excel帳票生成開始: 依頼NO={}, 取引先={}",
                delivery.getTmIraiNo(), delivery.getTorihikisakiCd());

        ClassPathResource templateResource = new ClassPathResource(
                "templates/excel/" + Mcm1001pConstants.TEMPLATE_FILE_NAME);

        try (InputStream is = templateResource.getInputStream();
             Workbook workbook = WorkbookFactory.create(is);
             ByteArrayOutputStream bos = new ByteArrayOutputStream()) {

            List<Mcm1001pHeaderDto> headerList = delivery.getHeaderList();
            if (headerList == null || headerList.isEmpty()) {
                throw new IOException("依頼書の出力対象がありません。");
            }
            // VBのSetTemplateSheetsと同様、書込み前の原紙から全シートを複製する。
            for (int i = 1; i < headerList.size(); i++) {
                workbook.cloneSheet(0);
            }

            for (int i = 0; i < headerList.size(); i++) {
                Mcm1001pHeaderDto header = headerList.get(i);

                /*
                 * 【変換元】Mcm1001pExcel.vb - シート複製処理
                 *   元コード: objExcel.SheetCopy(1, sheetName)
                 */
                Sheet sheet = workbook.getSheetAt(i);

                // シート名設定
                String sheetName = WorkbookUtil.createSafeSheetName(
                        Mcm1001pConstants.SHEET_NK + header.getKeiyakujikantai());
                String baseName = sheetName;
                for (int suffix = 2; workbook.getSheet(sheetName) != null
                        && workbook.getSheet(sheetName) != sheet; suffix++) {
                    String tail = "_" + suffix;
                    sheetName = baseName.substring(0, Math.min(baseName.length(), 31 - tail.length())) + tail;
                }
                int sheetIndex = i;
                workbook.setSheetName(sheetIndex, sheetName);

                // ヘッダー情報書き込み
                writeHeader(sheet, header);

                // 明細データ書き込み
                int lastDataRow = writeDetails(sheet, delivery.getMeisaiList(), header);

                // 罫線設定
                setBordersForAllColumns(sheet, lastDataRow);
            }

            workbook.write(bos);
            log.info("見積依頼書Excel帳票生成完了");
            return bos.toByteArray();
        }
    }

    /**
     * 出力ファイル名を生成する。
     *
     * 【変換元】Mcm1001pExcel.vb - ファイル名生成処理
     *   元コード: strCreateFileNK = tmIraiNo & "_" & torihikisakiCd & "_見積依頼書.xlsx"
     *
     * @param delivery 帳票出力用データ
     * @return URLエンコード済み出力ファイル名
     */
    public String getOutputFileName(Mcm1001pDeliveryDto delivery) {
        String rawName = delivery.getTmIraiNo() + "_"
                + delivery.getTorihikisakiCd() + "_"
                + Mcm1001pConstants.CREATE_FILE_NAME;
        return URLEncoder.encode(rawName, StandardCharsets.UTF_8)
                .replace("+", "%20");
    }

    // ========================================
    // ヘッダー書き込み
    // ========================================

    /**
     * ヘッダー情報をシートに書き込む。
     *
     * 【変換元】Mcm1001pExcel.vb - SetHeader()
     *   元コード: objExcel.SetCellValue(row, col, value)
     *
     * @param sheet 対象シート
     * @param header ヘッダー情報
     */
    private void writeHeader(Sheet sheet, Mcm1001pHeaderDto header) {
        // 見積依頼NO
        setCellValue(sheet,
                Mcm1001pConstants.HEADER_ROWINDEX_IRAI_NO - 1,
                Mcm1001pConstants.HEADER_COLINDEX_IRAI_NO - 1,
                header.getIraiNo());

        // 取引先名
        setCellValue(sheet,
                Mcm1001pConstants.HEADER_ROWINDEX_TORIHIKISAKI_NK - 1,
                Mcm1001pConstants.HEADER_COLINDEX_TORIHIKISAKI_NK - 1,
                header.getTorihikisakiNk());

        // 納入先名
        setCellValue(sheet,
                Mcm1001pConstants.HEADER_ROWINDEX_NONYUSAKI_NK - 1,
                Mcm1001pConstants.HEADER_COLINDEX_NONYUSAKI_NK - 1,
                header.getNonyusakiNk());

        // 納入先住所1
        setCellValue(sheet,
                Mcm1001pConstants.HEADER_ROWINDEX_NONYUSAKIJUSYO1_NK - 1,
                Mcm1001pConstants.HEADER_COLINDEX_NONYUSAKIJUSYO1_NK - 1,
                header.getNonyusakijusyo1Nk());

        // 納入先住所2
        setCellValue(sheet,
                Mcm1001pConstants.HEADER_ROWINDEX_NONYUSAKIJUSYO2_NK - 1,
                Mcm1001pConstants.HEADER_COLINDEX_NONYUSAKIJUSYO2_NK - 1,
                header.getNonyusakijusyo2Nk());

        // 回答希望日
        setCellValue(sheet,
                Mcm1001pConstants.HEADER_ROWINDEX_KAITOKIZITSU_DT - 1,
                Mcm1001pConstants.HEADER_COLINDEX_KAITOKIZITSU_DT - 1,
                header.getKaitokizitsuDt());

        // 依頼日（見積日）
        setCellValue(sheet,
                Mcm1001pConstants.HEADER_ROWINDEX_MITSUMORI_DT - 1,
                Mcm1001pConstants.HEADER_COLINDEX_MITSUMORI_DT - 1,
                header.getMitsumoriDt());

        // 契約時間帯
        setCellValue(sheet,
                Mcm1001pConstants.HEADER_ROWINDEX_KEIYAKUJIKANTAI - 1,
                Mcm1001pConstants.HEADER_COLINDEX_KEIYAKUJIKANTAI - 1,
                header.getKeiyakujikantai());

        // 保守方法
        setCellValue(sheet,
                Mcm1001pConstants.HEADER_ROWINDEX_HOSYUHOHO - 1,
                Mcm1001pConstants.HEADER_COLINDEX_HOSYUHOHO - 1,
                header.getHosyuhoho());

        // 保守点検（点検回数）
        setCellValue(sheet,
                Mcm1001pConstants.HEADER_ROWINDEX_HOSYUTENKEN - 1,
                Mcm1001pConstants.HEADER_COLINDEX_HOSYUTENKEN - 1,
                header.getTenkenumu());

        // 点検可能曜日
        setCellValue(sheet,
                Mcm1001pConstants.HEADER_ROWINDEX_TENKENKANOYOBI - 1,
                Mcm1001pConstants.HEADER_COLINDEX_TENKENKANOYOBI - 1,
                header.getTenkenkanobi());

        // 夜間点検
        setCellValue(sheet,
                Mcm1001pConstants.HEADER_ROWINDEX_YAKANTENKEN - 1,
                Mcm1001pConstants.HEADER_COLINDEX_YAKANTENKEN - 1,
                header.getYakanTenken());
    }

    // ========================================
    // 明細データ書き込み
    // ========================================

    /**
     * 明細データをシートに書き込む。
     *
     * 【変換元】Mcm1001pExcel.vb - SetIchiran()
     *   元コード: objExcel.SetCellValue(rowIdx, colIdx, meisaiData.Value)
     *
     * @param sheet 対象シート
     * @param allMeisaiList 全明細一覧
     * @param header ヘッダー情報（startIndex / endIndex で明細範囲を特定）
     * @return 最終データ行のインデックス（0ベース）
     */
    private int writeDetails(Sheet sheet, List<Mcm1001pViewDto> allMeisaiList,
                             Mcm1001pHeaderDto header) {
        int startRow = Mcm1001pConstants.ICHIRAN_ROWINDEX_DEFAULT - 1; // 0ベース変換
        int currentRow = startRow;

        // startIndex ～ endIndex の明細を書き込み
        int startIdx = header.getStartIndex();
        int endIdx = Math.min(header.getEndIndex(), allMeisaiList.size() - 1);

        for (int i = startIdx; i <= endIdx; i++) {
            Mcm1001pViewDto meisai = allMeisaiList.get(i);

            /*
             * 【変換元】Mcm1001pExcel.vb - SetIchiran() 内ループ
             *   元コード: objExcel.SetCellValue(intRowIdx, ICHIRAN_COLINDEX_SEIZOMAKER_NK, ...)
             */

            // 製造メーカー名
            setCellValue(sheet, currentRow,
                    Mcm1001pConstants.ICHIRAN_COLINDEX_SEIZOMAKER_NK - 1,
                    meisai.getSeizomakerNk());

            // 品名
            setCellValue(sheet, currentRow,
                    Mcm1001pConstants.ICHIRAN_COLINDEX_KIKIHINMEI_NK - 1,
                    meisai.getKikihinmeiNk());

            // 型式
            setCellValue(sheet, currentRow,
                    Mcm1001pConstants.ICHIRAN_COLINDEX_KIKIKATASHIKI - 1,
                    meisai.getKikikatashiki());

            // 数量
            if (meisai.getSuryoNm() != null) {
                setCellValue(sheet, currentRow,
                        Mcm1001pConstants.ICHIRAN_COLINDEX_SURYO_NM - 1,
                        meisai.getSuryoNm().doubleValue());
            }

            // 手配製番
            setCellValue(sheet, currentRow,
                    Mcm1001pConstants.ICHIRAN_COLINDEX_TEHAISEIBAN - 1,
                    meisai.getTehaiseiban());

            currentRow++;
        }

        return currentRow - 1; // 最終データ行を返す
    }

    // ========================================
    // 罫線設定
    // ========================================

    /**
     * 全カラムグループに対して罫線を設定する。
     *
     * 【変換元】Mcm1001pExcel.vb - SetBorders()
     *   元コード: objExcel.SetBorders(startRow, startCol, endRow, endCol, xlThin)
     *
     * @param sheet 対象シート
     * @param lastDataRow 最終データ行（0ベース）
     */
    private void setBordersForAllColumns(Sheet sheet, int lastDataRow) {
        int startRow = Mcm1001pConstants.ICHIRAN_ROWINDEX_DEFAULT - 1; // 0ベース変換

        if (lastDataRow < startRow) {
            return; // データなしの場合はスキップ
        }

        // 製造メーカー名
        setBorders(sheet, startRow,
                Mcm1001pConstants.ICHIRAN_COLINDEX_SEIZOMAKER_NK - 1,
                lastDataRow,
                Mcm1001pConstants.ICHIRAN_COLINDEX_SEIZOMAKER_NK_END - 1);

        // 品名
        setBorders(sheet, startRow,
                Mcm1001pConstants.ICHIRAN_COLINDEX_KIKIHINMEI_NK - 1,
                lastDataRow,
                Mcm1001pConstants.ICHIRAN_COLINDEX_KIKIHINMEI_NK_END - 1);

        // 型式
        setBorders(sheet, startRow,
                Mcm1001pConstants.ICHIRAN_COLINDEX_KIKIKATASHIKI - 1,
                lastDataRow,
                Mcm1001pConstants.ICHIRAN_COLINDEX_KIKIKATASHIKI_END - 1);

        // 数量
        setBorders(sheet, startRow,
                Mcm1001pConstants.ICHIRAN_COLINDEX_SURYO_NM_START - 1,
                lastDataRow,
                Mcm1001pConstants.ICHIRAN_COLINDEX_SURYO_NM - 1);

        // 手配製番
        setBorders(sheet, startRow,
                Mcm1001pConstants.ICHIRAN_COLINDEX_TEHAISEIBAN - 1,
                lastDataRow,
                Mcm1001pConstants.ICHIRAN_COLINDEX_TEHAISEIBAN_END - 1);

        // 備考
        setBorders(sheet, startRow,
                Mcm1001pConstants.ICHIRAN_COLINDEX_BIKO - 1,
                lastDataRow,
                Mcm1001pConstants.ICHIRAN_COLINDEX_BIKO_END - 1);
    }

    /**
     * 指定範囲に対して薄い罫線（THIN）を設定する。
     *
     * 【変換元】Mcm1001pExcel.vb - 罫線設定処理
     *   元コード: objExcel.SetBorders(sr, sc, er, ec, Excel.XlBorderWeight.xlThin)
     *   変換: Excel Interop Range.Borders → POI RegionUtil
     *
     * @param sheet 対象シート
     * @param startRow 開始行（0ベース）
     * @param startCol 開始列（0ベース）
     * @param endRow 終了行（0ベース）
     * @param endCol 終了列（0ベース）
     */
    private void setBorders(Sheet sheet, int startRow, int startCol,
                            int endRow, int endCol) {
        for (int r = startRow; r <= endRow; r++) {
            CellRangeAddress region = new CellRangeAddress(r, r, startCol, endCol);

            RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
            RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
        }
    }

    // ========================================
    // セル値設定ユーティリティ
    // ========================================

    /**
     * セルに文字列値を設定する。
     *
     * 【変換元】Mcm1001pExcel.vb - objExcel.SetCellValue()
     *
     * @param sheet 対象シート
     * @param rowIndex 行インデックス（0ベース）
     * @param colIndex 列インデックス（0ベース）
     * @param value 設定値
     */
    private void setCellValue(Sheet sheet, int rowIndex, int colIndex, String value) {
        if (value == null) {
            return;
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        cell.setCellValue(value);
    }

    /**
     * セルに数値を設定する。
     *
     * 【変換元】Mcm1001pExcel.vb - objExcel.SetCellValue()（数量用）
     *
     * @param sheet 対象シート
     * @param rowIndex 行インデックス（0ベース）
     * @param colIndex 列インデックス（0ベース）
     * @param value 設定値
     */
    private void setCellValue(Sheet sheet, int rowIndex, int colIndex, double value) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        cell.setCellValue(value);
    }
}
