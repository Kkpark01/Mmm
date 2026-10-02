package com.daifuku.mcm.service;

import com.daifuku.mcm.common.BaseExcelService;
import com.daifuku.mcm.constants.Mcm1001pConstants;
import com.daifuku.mcm.dto.Mcm1001pDeliveryDto;
import com.daifuku.mcm.dto.Mcm1001pHeaderDto;
import com.daifuku.mcm.dto.Mcm1001pViewDto;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;
import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/** VB Mcm1001pExcel.vbの項目位置・取引先別／時間帯別出力を移植する。 */
@Service
public class Mcm1001pExcelService extends BaseExcelService {
    private static final int FIRST_DETAIL_ROW = Mcm1001pConstants.ICHIRAN_ROWINDEX_DEFAULT - 1;
    private static final int[][] COLUMN_GROUPS = {{0,5},{6,19},{20,31},{32,33},{34,37},{38,54}};

    public byte[] generateReport(Mcm1001pDeliveryDto delivery) throws IOException {
        if (delivery == null || delivery.getHeaderList() == null || delivery.getHeaderList().isEmpty()
                || delivery.getMeisaiList() == null) {
            throw new IOException("依頼書の出力対象がありません。");
        }
        var resource = new ClassPathResource("templates/excel/" + Mcm1001pConstants.TEMPLATE_FILE_NAME);
        try (var input = resource.getInputStream(); Workbook workbook = WorkbookFactory.create(input)) {
            var headers = delivery.getHeaderList();
            // 書き込み前の原紙を複製して、前の時間帯の値を混ぜない。
            for (int i = 1; i < headers.size(); i++) workbook.cloneSheet(0);
            for (int i = 0; i < headers.size(); i++) {
                var header = headers.get(i);
                Sheet sheet = workbook.getSheetAt(i);
                String base = WorkbookUtil.createSafeSheetName(header.getKeiyakujikantai());
                if (base == null || base.isBlank()) throw new IOException("契約時間帯が未設定です。");
                String name = base;
                for (int suffix = 2; workbook.getSheet(name) != null && workbook.getSheet(name) != sheet; suffix++) {
                    String tail = "_" + suffix;
                    name = base.substring(0, Math.min(base.length(), 31 - tail.length())) + tail;
                }
                workbook.setSheetName(i, name);
                writeHeader(sheet, header);
                writeDetails(sheet, delivery.getMeisaiList(), header);
            }
            return toByteArray(workbook);
        }
    }

    public String getRawOutputFileName(Mcm1001pDeliveryDto delivery) {
        return safePart(delivery.getTmIraiNo()) + "_" + safePart(delivery.getTorihikisakiCd())
                + "_" + Mcm1001pConstants.CREATE_FILE_NAME;
    }

    /** 既存呼出しとの互換用。HTTPヘッダーでは生の名前からContentDispositionを組み立てる。 */
    public String getOutputFileName(Mcm1001pDeliveryDto delivery) {
        return URLEncoder.encode(getRawOutputFileName(delivery), StandardCharsets.UTF_8).replace("+", "%20");
    }

    private static String safePart(String value) {
        return value == null ? "" : value.replaceAll("[\\\\/:*?\"<>|\\p{Cntrl}]", "_");
    }

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

    private void writeDetails(Sheet sheet, List<Mcm1001pViewDto> details,
                              Mcm1001pHeaderDto header) throws IOException {
        int start = header.getStartIndex(), end = header.getEndIndex();
        if (start < 0 || start > details.size() || end < start - 1 || end >= details.size()) {
            throw new IOException("依頼書の明細範囲が不正です。");
        }
        Row original = sheet.getRow(FIRST_DETAIL_ROW);
        short height = original == null ? sheet.getDefaultRowHeight() : original.getHeight();
        // 新しい行にも原紙の書式を継承する。列幅・印刷設定は原紙の設定を保つ。
        CellStyle[] originalStyles = new CellStyle[55];
        for (int c = 0; c < 55; c++) {
            originalStyles[c] = original != null && original.getCell(c) != null
                    ? original.getCell(c).getCellStyle() : sheet.getWorkbook().getCellStyleAt(0);
        }
        Map<String, CellStyle> borderStyles = new HashMap<>();
        for (int i = start; i <= end; i++) {
            int r = FIRST_DETAIL_ROW + i - start;
            if (r >= sheet.getWorkbook().getSpreadsheetVersion().getMaxRows()) {
                throw new IOException("依頼書の明細がExcel形式の上限を超えています。");
            }
            Row row = sheet.getRow(r);
            if (row == null) row = sheet.createRow(r);
            row.setHeight(height);
            for (int c = 0; c < 55; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) cell = row.createCell(c);
                cell.setCellStyle(originalStyles[c]);
                cell.setBlank();
            }
            var detail = details.get(i);
            setCellValue(sheet, r, 0, detail.getSeizomakerNk());
            setCellValue(sheet, r, 6, detail.getKikihinmeiNk());
            setCellValue(sheet, r, 20, detail.getKikikatashiki());
            setCellValue(sheet, r, 33, detail.getSuryoNm());
            setCellValue(sheet, r, 34, detail.getTehaiseiban());
            // 備考はVB同様空欄。文字列は数式に変換せず、そのまま書き込む。
            for (int[] group : COLUMN_GROUPS) {
                for (int c = group[0]; c <= group[1]; c++) {
                    Cell cell = row.getCell(c);
                    boolean first = i == start, last = i == end, left = c == group[0], right = c == group[1];
                    String key = originalStyles[c].getIndex()+":"+first+":"+last+":"+left+":"+right;
                    CellStyle style = borderStyles.get(key);
                    if (style == null) {
                        style = sheet.getWorkbook().createCellStyle();
                        style.cloneStyleFrom(originalStyles[c]);
                        style.setBorderTop(first ? BorderStyle.THIN : BorderStyle.HAIR);
                        style.setBorderBottom(last ? BorderStyle.THIN : BorderStyle.HAIR);
                        style.setBorderLeft(left ? BorderStyle.THIN : BorderStyle.NONE);
                        style.setBorderRight(right ? BorderStyle.THIN : BorderStyle.NONE);
                        borderStyles.put(key, style);
                    }
                    cell.setCellStyle(style);
                }
            }
        }
    }
}
