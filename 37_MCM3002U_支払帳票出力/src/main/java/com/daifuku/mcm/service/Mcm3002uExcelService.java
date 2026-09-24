package com.daifuku.mcm.service;

import com.daifuku.mcm.form.Mcm3002uForm.RowForm;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;
import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.YearMonth;
import java.util.*;
import java.util.zip.*;

/** VB Mcm3001pExcel / Mcm3005pExcel。提供されたxls原紙の書式を保持する。 */
@Service
public class Mcm3002uExcelService {
    public byte[] generateZip(String month, String paymentType, List<RowForm> rows) throws IOException {
        if (month == null || !month.matches("[0-9]{6}")
                || !("年払".equals(paymentType) || "月払".equals(paymentType))) {
            throw new IllegalArgumentException("帳票の出力条件が不正です。");
        }
        var ym = YearMonth.of(Integer.parseInt(month.substring(0, 4)), Integer.parseInt(month.substring(4)));
        if (rows == null || rows.isEmpty()) throw new IllegalArgumentException("対象データが存在しません。");
        var groups = new LinkedHashMap<BigDecimal, List<RowForm>>();
        for (var row : rows) {
            if (row.getTorihikisakiId() == null || row.getShiharaiKingaku() == null) {
                throw new IllegalArgumentException("取引先または支払金額が取得できません。");
            }
            groups.computeIfAbsent(row.getTorihikisakiId().stripTrailingZeros(), k -> new ArrayList<>()).add(row);
        }
        // 両帳票の行数を原紙の読込み・複製より前に確認する。
        // 発注確認書はヘッダーと合計を含め16行、支払い明細は14行を使用する。
        int maxDetails = org.apache.poi.ss.SpreadsheetVersion.EXCEL97.getMaxRows() - 16;
        if (groups.values().stream().anyMatch(group -> group.size() > maxDetails)) {
            throw new IllegalArgumentException("取引先別の明細件数がExcelの行数上限を超えています。");
        }
        // 両帳票が完成するまでレスポンスを開始しない。片方だけの成功を防ぐ。
        try (var out = new ByteArrayOutputStream(); var zip = new ZipOutputStream(out)) {
            for (boolean order : new boolean[]{true, false}) {
                var name = month + "_" + paymentType + (order ? "_発注確認書.xls" : "_支払い明細.xls");
                byte[] data = generateWorkbook(ym, paymentType, new ArrayList<>(groups.values()), order);
                zip.putNextEntry(new ZipEntry(name));
                zip.write(data);
                zip.closeEntry();
            }
            zip.finish();
            return out.toByteArray();
        }
    }

    private byte[] generateWorkbook(YearMonth month, String paymentType, List<List<RowForm>> groups,
                                    boolean order) throws IOException {
        String file = order ? "MCM3001P_発注確認書テンプレート.xls" : "MCM3005P_支払い明細テンプレート.xls";
        try (var in = new ClassPathResource("templates/excel/" + file).getInputStream();
             Workbook book = WorkbookFactory.create(in); var out = new ByteArrayOutputStream()) {
            if (book.getNumberOfSheets() != 1) throw new IOException("帳票テンプレートのシート構成が不正です。");
            for (int i = 1; i < groups.size(); i++) book.cloneSheet(0);
            var styles = new HashMap<String, CellStyle>();
            for (int i = 0; i < groups.size(); i++) {
                var rows = groups.get(i);
                Sheet sheet = book.getSheetAt(i);
                String base = WorkbookUtil.createSafeSheetName(text(rows.get(0).getTorihikisakiCd()));
                if (base.isBlank()) base = "取引先";
                String name = base;
                for (int n = 2; book.getSheet(name) != null && book.getSheet(name) != sheet; n++) {
                    String tail = "_" + n;
                    name = base.substring(0, Math.min(base.length(), 31 - tail.length())) + tail;
                }
                book.setSheetName(i, name);
                sheet.setRepeatingRows(new CellRangeAddress(order ? 0 : 8, order ? 14 : 8, -1, -1));
                sheet.setRepeatingColumns(null);
                writeSheet(sheet, month, paymentType, rows, order, styles);
            }
            book.getCreationHelper().createFormulaEvaluator().evaluateAll();
            book.setForceFormulaRecalculation(true);
            book.write(out);
            return out.toByteArray();
        }
    }

    private void writeSheet(Sheet sheet, YearMonth month, String paymentType, List<RowForm> rows,
                            boolean order, Map<String, CellStyle> styles) {
        var first = rows.get(0);
        String monthText = month.getYear() + "年" + String.format(java.util.Locale.ROOT, "%02d", month.getMonthValue());
        if (order) {
            put(sheet, 1, 1, text(first.getTorihikisakiNk()) + " " + text(first.getTorisyutantosyaNk()) + " 様");
            put(sheet, 4, 3, monthText);
            put(sheet, 4, 7, "( " + paymentType + " )");
        } else {
            put(sheet, 0, 0, monthText + "月分 支払い明細 (" + text(first.getTorihikisakiNk()) + ")");
        }
        int start = order ? 15 : 9, left = order ? 1 : 0, right = order ? 9 : 8;
        CellStyle[] original = new CellStyle[right + 1];
        for (int c = left; c <= right; c++) original[c] = cell(sheet, start, c).getCellStyle();
        short height = sheet.getRow(start).getHeight();
        var sum = BigDecimal.ZERO;
        for (int n = 0; n < rows.size(); n++) {
            int r = start + n;
            for (int c = left; c <= right; c++) cell(sheet, r, c).setCellStyle(original[c]);
            sheet.getRow(r).setHeight(height);
            var row = rows.get(n);
            // VB Screen側は支払い明細のみCIntを通す（銀行丸め）。発注確認書はDecimal。
            var amount = order ? row.getShiharaiKingaku() : row.getShiharaiKingaku().setScale(0, RoundingMode.HALF_EVEN);
            sum = sum.add(amount);
            if (order) {
                put(sheet,r,1,n+1); put(sheet,r,2,row.getKeiyakuNo());
                put(sheet,r,3,month.getYear()); put(sheet,r,4,month.getMonthValue());
                put(sheet,r,5,row.getJotai()); put(sheet,r,6,row.getNonyusakiNk());
                put(sheet,r,8,amount); put(sheet,r,9,row.getHardSeiban());
            } else {
                put(sheet,r,0,row.getSupportId()); put(sheet,r,1,row.getNonyusakiNk());
                put(sheet,r,2,row.getTenpoNk()); put(sheet,r,3,row.getKeiyakuNo());
                put(sheet,r,4,row.getTorihikisakiNk()); put(sheet,r,5,amount);
                put(sheet,r,6,0); put(sheet,r,7,row.getKaisu()); put(sheet,r,8,row.getHardSeiban());
            }
        }
        int last = start + rows.size() - 1;
        if (order) {
            put(sheet,last+1,7,"合計"); put(sheet,last+1,8,sum);
            for (int c=left;c<=right;c++) cell(sheet,last+1,c).setCellStyle(original[c]);
            borders(sheet,start,last+1,left,right,true,styles);
        } else {
            int total = rows.size()+11, count = rows.size()+13;
            put(sheet,total,4,"合計"); put(sheet,total,5,sum);
            cell(sheet,total,6).setCellFormula("SUM(G10:G"+(last+1)+")");
            put(sheet,count,4,"件数"); put(sheet,count,5,rows.size()+"件");
            cell(sheet,count,6).setCellFormula("COUNTIF(G10:G"+(last+1)+",\">0\")&\"件\"");
            borders(sheet,start,last,left,right,false,styles);
            for(int r:new int[]{total,count}) for(int c=4;c<=6;c++) {
                Cell cell=cell(sheet,r,c);
                String key="total:"+cell.getCellStyle().getIndex()+":"+c;
                final int column=c;
                cell.setCellStyle(styles.computeIfAbsent(key,k->{
                    var style=sheet.getWorkbook().createCellStyle();style.cloneStyleFrom(cell.getCellStyle());
                    style.setAlignment(HorizontalAlignment.RIGHT);
                    if(column!=4) style.setBorderBottom(BorderStyle.THIN);
                    return style;
                }));
            }
        }
    }

    /** CPExcelManager.setBorders：外枠は細線、内横線は極細、G:Hの内縦線はなし。 */
    private void borders(Sheet sheet,int first,int last,int left,int right,boolean order,Map<String,CellStyle> cache) {
        for(int r=first;r<=last;r++) for(int c=left;c<=right;c++) {
            Cell cell=cell(sheet,r,c);
            boolean noLeft=order && c==7, noRight=order && c==6, total=order && r==last;
            String key=cell.getCellStyle().getIndex()+":"+(r==first)+":"+(r==last)+":"+noLeft+":"+noRight+":"+total;
            var style=cache.get(key);
            if(style==null) {
                style=sheet.getWorkbook().createCellStyle();style.cloneStyleFrom(cell.getCellStyle());
                style.setBorderTop(r==first?BorderStyle.THIN:BorderStyle.HAIR);
                style.setBorderBottom(r==last?BorderStyle.THIN:BorderStyle.HAIR);
                style.setBorderLeft(noLeft?BorderStyle.NONE:BorderStyle.THIN);
                style.setBorderRight(noRight?BorderStyle.NONE:BorderStyle.THIN);
                if(total){style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());style.setFillPattern(FillPatternType.SOLID_FOREGROUND);style.setAlignment(HorizontalAlignment.RIGHT);}
                cache.put(key,style);
            }
            cell.setCellStyle(style);
        }
    }
    private static Cell cell(Sheet sheet,int r,int c) {
        Row row=sheet.getRow(r);if(row==null)row=sheet.createRow(r);
        Cell cell=row.getCell(c);return cell==null?row.createCell(c):cell;
    }
    private static void put(Sheet sheet,int r,int c,Object value) {
        Cell cell=cell(sheet,r,c);
        if(value instanceof Number number) cell.setCellValue(number.doubleValue());
        else cell.setCellValue(value==null?"":value.toString());
    }
    private static String text(String value){return value==null?"":value;}
}
