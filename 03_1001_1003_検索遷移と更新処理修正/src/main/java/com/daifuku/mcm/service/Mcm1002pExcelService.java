package com.daifuku.mcm.service;

import com.daifuku.mcm.constants.Mcm1002pConstants;
import com.daifuku.mcm.dto.Mcm1002pDataViewDto;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * 【変換元】Mcm1002pExcel.vb - MCM1002P 見積依頼リスト Excel出力サービス
 * 【変換元行数】約200行
 * 【作成者】T.Kajimura (2009/03/11)
 * 【更新者】M.Ohsuka (2010/03/23)
 *
 * 元コード: CPExcelManager を継承し、Microsoft.Office.Interop.Excel で出力。
 *           Initialize() でテンプレートファイル名をセット → Execute() でExcel生成。
 * 変換後:   Apache POI (XSSFWorkbook) でExcel出力。テンプレートは classpath から読込。
 */
@Service
public class Mcm1002pExcelService {

    private static final Logger log = LoggerFactory.getLogger(Mcm1002pExcelService.class);

    /**
     * 見積依頼リストをExcel出力する。
     *
     * 【変換元】Mcm1002pExcel.vb - WriteDataInExcel()
     *   元コード: Dim sheets As Excel.Sheets = objExcelApp.Worksheets
     *             Dim sheet As Excel.Worksheet = CType(sheets(startPos), Excel.Worksheet)
     *             sheet.Name = Mcm1002pConstant.SHEET_NK
     *             For j = 0 To uvaList.Count - 1
     *               sheet.Cells(cellCount + ROWINDEX, COLINDEX).Value = ...
     *             Next
     *             setBorders(...)
     *
     * @param dataList 一覧データ（画面から受け渡し）
     * @return Excelファイルのバイト配列
     * @throws IOException テンプレート読込・書込エラー時
     */
    public byte[] generate(List<Mcm1002pDataViewDto> dataList) throws IOException {
        log.info("MCM1002P Excel生成開始: データ件数={}", dataList.size());

        /*
         * 【変換元】Mcm1002pExcel.vb - Initialize()
         *   元コード: MyBase.TemplateFileName = Mcm1002pConstant.TEMPLATE_FILE_NAME
         *             MyBase.CreateFileName   = Mcm1002pConstant.CREATE_FILE_NAME
         */
        ClassPathResource templateResource =
                new ClassPathResource("templates/excel/" + Mcm1002pConstants.TEMPLATE_FILE_NAME);

        try (Workbook workbook = openWorkbook(templateResource);
             ByteArrayOutputStream bos = new ByteArrayOutputStream()) {

            // シート取得（1シート目）
            Sheet sheet = workbook.getSheetAt(0);

            /*
             * 【変換元】Mcm1002pExcel.vb - WriteDataInExcel()
             *   元コード: sheet.Name = Mcm1002pConstant.SHEET_NK
             */
            workbook.setSheetName(0, Mcm1002pConstants.SHEET_NK);

            // ===================================================================
            // 罫線用スタイル
            // 【変換元】Mcm1002pExcel.vb - setBorders()
            //   元コード: setBorders(sheet, startRow, endRow, startCol, endCol,
            //              VerticalLineStyle_Nashi)
            // ===================================================================
            CellStyle borderStyle = workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);
            borderStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            // 数値用スタイル（時間帯列：右寄せ＋罫線）
            CellStyle numberBorderStyle = workbook.createCellStyle();
            numberBorderStyle.cloneStyleFrom(borderStyle);
            numberBorderStyle.setAlignment(HorizontalAlignment.RIGHT);

            /*
             * 【変換元】Mcm1002pExcel.vb - WriteDataInExcel()
             *   元コード:
             *     Dim uvaList As List(Of Mcm1002pData_View) = delivery.DgvArray
             *     Dim cellCount As Integer = 0
             *     For j As Integer = 0 To uvaList.Count - 1
             *       sheet.Cells(cellCount + ICHIRAN_ROWINDEX_DEFAULT,
             *                   ICHIRAN_COLINDEX_IRAI_NO).Value = uvaList(j).IRAI_NO
             *       ...
             *       cellCount += 1
             *     Next
             *
             * ※ VB.NETは1-based → POIは0-basedのため、行・列ともに -1 して変換
             */
            int startRowIndex = Mcm1002pConstants.ICHIRAN_ROWINDEX_DEFAULT - 1; // 0-based: row 1

            for (int j = 0; j < dataList.size(); j++) {
                Mcm1002pDataViewDto rowData = dataList.get(j);
                int rowIdx = startRowIndex + j;
                Row row = sheet.getRow(rowIdx);
                if (row == null) {
                    row = sheet.createRow(rowIdx);
                }

                // --- 依頼No（col 0） ---
                // 元コード: sheet.Cells(..., ICHIRAN_COLINDEX_IRAI_NO).Value = uvaList(j).IRAI_NO
                setCellValue(row,
                        Mcm1002pConstants.ICHIRAN_COLINDEX_IRAI_NO - 1,
                        rowData.getIraiNo(), borderStyle);

                // --- 契約時間帯（col 1）※数値 ---
                // 元コード: sheet.Cells(..., ICHIRAN_COLINDEX_JIKANTAI).Value = uvaList(j).JIKANTAI
                Cell jikantaiCell = getOrCreateCell(row,
                        Mcm1002pConstants.ICHIRAN_COLINDEX_JIKANTAI - 1);
                if (rowData.getJikantai() != null) {
                    jikantaiCell.setCellValue(rowData.getJikantai().doubleValue());
                }
                jikantaiCell.setCellStyle(numberBorderStyle);

                // --- 状態（col 2） ---
                // 元コード: sheet.Cells(..., ICHIRAN_COLINDEX_JOTAI).Value = uvaList(j).JOTAI
                setCellValue(row,
                        Mcm1002pConstants.ICHIRAN_COLINDEX_JOTAI - 1,
                        rowData.getJotai(), borderStyle);

                // --- 契約NO（col 3） ---
                // 元コード: sheet.Cells(..., ICHIRAN_COLINDEX_KEIYAKU_NO).Value = uvaList(j).KEIYAKU_NO
                // ※ 2010/03/23 M.Ohsuka 追加
                setCellValue(row,
                        Mcm1002pConstants.ICHIRAN_COLINDEX_KEIYAKU_NO - 1,
                        rowData.getKeiyakuNo(), borderStyle);

                // --- 取引先名（col 4） ---
                // 元コード: sheet.Cells(..., ICHIRAN_COLINDEX_TORIHIKISAKI_NK).Value = uvaList(j).TORIHIKISAKI_NK
                setCellValue(row,
                        Mcm1002pConstants.ICHIRAN_COLINDEX_TORIHIKISAKI_NK - 1,
                        rowData.getTorihikisakiNk(), borderStyle);

                // --- 納入先名（col 5） ---
                // 元コード: sheet.Cells(..., ICHIRAN_COLINDEX_NONYUSAKI_NK).Value = uvaList(j).NONYUSAKI_NK
                setCellValue(row,
                        Mcm1002pConstants.ICHIRAN_COLINDEX_NONYUSAKI_NK - 1,
                        rowData.getNonyusakiNk(), borderStyle);

                // --- サポートID（col 6） ---
                // 元コード: sheet.Cells(..., ICHIRAN_COLINDEX_SUPPORT_ID).Value = uvaList(j).SUPPORT_ID
                setCellValue(row,
                        Mcm1002pConstants.ICHIRAN_COLINDEX_SUPPORT_ID - 1,
                        rowData.getSupportId(), borderStyle);

                // --- プラント名（col 7） ---
                // 元コード: sheet.Cells(..., ICHIRAN_COLINDEX_PLANT_NK).Value = uvaList(j).PLANT_NK
                setCellValue(row,
                        Mcm1002pConstants.ICHIRAN_COLINDEX_PLANT_NK - 1,
                        rowData.getPlantNk(), borderStyle);

                // --- 回答期日（col 8） ---
                // 元コード: sheet.Cells(..., ICHIRAN_COLINDEX_KAITO_DT).Value = uvaList(j).KAITO_DT
                setCellValue(row,
                        Mcm1002pConstants.ICHIRAN_COLINDEX_KAITO_DT - 1,
                        rowData.getKaitoDt(), borderStyle);

                // --- 依頼日（col 9） ---
                // 元コード: sheet.Cells(..., ICHIRAN_COLINDEX_IRAI_DT).Value = uvaList(j).IRAI_DT
                setCellValue(row,
                        Mcm1002pConstants.ICHIRAN_COLINDEX_IRAI_DT - 1,
                        rowData.getIraiDt(), borderStyle);

                // --- 依頼担当者（col 10） ---
                // 元コード: sheet.Cells(..., ICHIRAN_COLINDEX_IRAISHA).Value = uvaList(j).IRAISYA
                setCellValue(row,
                        Mcm1002pConstants.ICHIRAN_COLINDEX_IRAISHA - 1,
                        rowData.getIraisya(), borderStyle);
            }

            /*
             * 【変換元】Mcm1002pExcel.vb - WriteDataInExcel()
             *   元コード: setBorders(sheet,
             *              ICHIRAN_ROWINDEX_DEFAULT,
             *              cellCount + ICHIRAN_ROWINDEX_DEFAULT - 1,
             *              ICHIRAN_COLINDEX_IRAI_NO,
             *              ICHIRAN_COLINDEX_IRAISHA,
             *              VerticalLineStyle_Nashi)
             *   → 上記ループ内で各セルにborderStyleを適用済み
             */

            workbook.write(bos);
            log.info("MCM1002P Excel生成完了");
            return bos.toByteArray();
        }
    }

    /**
     * セルに文字列値を設定しスタイルを適用する。
     *
     * @param row    対象の行
     * @param colIdx 列インデックス（0-based）
     * @param value  セルに設定する文字列値
     * @param style  適用するセルスタイル
     */
    private void setCellValue(Row row, int colIdx, String value, CellStyle style) {
        Cell cell = getOrCreateCell(row, colIdx);
        if (value != null) {
            cell.setCellValue(value);
        }
        cell.setCellStyle(style);
    }

    /**
     * セルを取得する。存在しない場合は新規作成する。
     *
     * @param row    対象の行
     * @param colIdx 列インデックス（0-based）
     * @return セル
     */
    private Cell getOrCreateCell(Row row, int colIdx) {
        Cell cell = row.getCell(colIdx);
        if (cell == null) {
            cell = row.createCell(colIdx);
        }
        return cell;
    }

    private Workbook openWorkbook(ClassPathResource template) throws IOException {
        if (template.exists()) {
            try (InputStream input = template.getInputStream()) {
                return new XSSFWorkbook(input);
            }
        }
        // 配布srcには帳票テンプレートがないため、既存の列定義に従った一覧を作成する。
        // テンプレートが配備されている環境の書式・印刷設定は変更しない。
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(Mcm1002pConstants.SHEET_NK);
        String[] headers = {"依頼NO", "契約時間帯", "状態", "契約NO", "取引先名", "納入先名",
                "サポートID", "プラント名", "回答期日", "依頼日", "依頼担当者"};
        int[] widths = {20, 12, 12, 20, 28, 28, 18, 36, 16, 16, 20};
        CellStyle headerStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        headerStyle.setFont(font);
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Row header = sheet.createRow(0);
        for (int column = 0; column < headers.length; column++) {
            Cell cell = header.createCell(column);
            cell.setCellValue(headers[column]);
            cell.setCellStyle(headerStyle);
            sheet.setColumnWidth(column, widths[column] * 256);
        }
        sheet.createFreezePane(0, 1);
        sheet.setRepeatingRows(new org.apache.poi.ss.util.CellRangeAddress(0, 0, -1, -1));
        sheet.getPrintSetup().setLandscape(true);
        sheet.getPrintSetup().setFitWidth((short) 1);
        sheet.getPrintSetup().setFitHeight((short) 0);
        sheet.setFitToPage(true);
        return workbook;
    }
}
