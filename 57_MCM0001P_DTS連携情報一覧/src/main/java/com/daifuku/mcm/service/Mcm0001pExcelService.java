package com.daifuku.mcm.service;

import com.daifuku.mcm.common.BaseExcelService;
import com.daifuku.mcm.dto.Mcm0001pData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.format.ResolverStyle;
import java.util.List;

/** Mcm0001pExcel.vb + CPExcelManagerの可視列・汎用Excel出力。 */
@Service
public class Mcm0001pExcelService extends BaseExcelService {
    public static final String FILE_NAME = "DTS連携情報一覧.xlsx";
    private static final String[] HEADERS = {
        "ログ出力日", "納入先コード", "納入先名", "サポートID", "プラント名", "内容"
    };
    private static final DateTimeFormatter LOG_DATE =
        DateTimeFormatter.ofPattern("uuuu/MM/dd").withResolverStyle(ResolverStyle.STRICT);

    public byte[] generate(List<Mcm0001pData> dataList) throws IOException {
        try (Workbook workbook = createNewWorkbook()) {
            Sheet sheet = workbook.createSheet("DTS連携情報一覧");
            CellStyle textStyle = borderedStyle(workbook, "@");
            CellStyle dateStyle = borderedStyle(workbook, "yyyy/mm/dd");
            XSSFCellStyle headerStyle = (XSSFCellStyle) borderedStyle(workbook, "@");
            // VBのExcel.ColorIndex=20（CPColorConstant）。POIのindexed番号とは異なる。
            headerStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 204, (byte) 255, (byte) 255}, null));
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            for (int col = 0; col < HEADERS.length; col++) {
                setCellValue(sheet, 0, col, HEADERS[col]);
                sheet.getRow(0).getCell(col).setCellStyle(headerStyle);
            }
            if (dataList != null) {
                for (int index = 0; index < dataList.size(); index++) {
                    Mcm0001pData data = dataList.get(index);
                    String[] values = {data.getLogDt(), data.getNonyusakiCd(), data.getNonyusakiNk(),
                        data.getSupportId(), data.getPlantNk(), data.getNaiyo()};
                    for (int col = 0; col < values.length; col++) {
                        setCellValue(sheet, index + 1, col, values[col]);
                        Cell cell = sheet.getRow(index + 1).getCell(col);
                        cell.setCellStyle(col == 0 ? dateStyle : textStyle);
                        if (col == 0 && values[col] != null && !values[col].isBlank()) {
                            try {
                                // 書式指定だけで文字列を書いてもExcelの日付にはならない。
                                cell.setCellValue(LocalDate.parse(values[col], LOG_DATE));
                            } catch (DateTimeParseException ex) {
                                // 不正な既存値も一覧の内容として保持する。
                                cell.setCellStyle(textStyle);
                            }
                        }
                    }
                }
            }
            for (int col = 0; col < HEADERS.length; col++) sheet.autoSizeColumn(col);
            return toByteArray(workbook);
        }
    }

    private CellStyle borderedStyle(Workbook workbook, String format) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat(format));
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }
}
