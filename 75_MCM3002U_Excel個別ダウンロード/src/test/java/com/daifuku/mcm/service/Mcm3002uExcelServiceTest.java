package com.daifuku.mcm.service;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import java.io.ByteArrayInputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.jupiter.api.Test;

import com.daifuku.mcm.form.Mcm3002uForm.RowForm;

/**
 * 【変換元】Mcm3001pExcel / Mcm3005pExcel
 * Mcm3002uExcelService の帳票出力を検証する。
 *
 * 優先度2対応：
 *   - テンプレート原紙が .xls（HSSF）形式で実在することの確認（拡張子不一致の回帰防止）
 *   - 取引先別シート分割が正しく行われることの確認
 */
class Mcm3002uExcelServiceTest {

    private final Mcm3002uExcelService service = new Mcm3002uExcelService();

    private RowForm row(String torihikisakiId, String torihikisakiCd, String torihikisakiNk, String amount) {
        var row = new RowForm();
        row.setTorihikisakiId(new BigDecimal(torihikisakiId));
        row.setTorihikisakiCd(torihikisakiCd);
        row.setTorihikisakiNk(torihikisakiNk);
        row.setTorisyutantosyaNk("担当者");
        row.setKeiyakuNo("K-0001");
        row.setJotai("新規");
        row.setNonyusakiNk("納入先");
        row.setTenpoNk("店舗");
        row.setSupportId("S-0001");
        row.setKaisu(BigDecimal.ONE);
        row.setHardSeiban("HW-0001");
        row.setShiharaiKingaku(new BigDecimal(amount));
        row.setKaishiDt(LocalDate.of(2026, 1, 1));
        row.setSyuryoDt(LocalDate.of(2026, 12, 31));
        return row;
    }

    /** VBと同じ順序で2帳票を返し、ZIPを介さずExcelとして開けること。 */
    @Test
    void generatesTwoSeparateExcelFilesInVbOrder() throws Exception {
        var reports = service.generateReports("202601", "年払",
                List.of(row("1", "T001", "取引先A", "10000")));
        assertThat(reports).extracting(Mcm3002uExcelService.ReportFile::filename)
                .containsExactly("202601_年払_発注確認書.xls", "202601_年払_支払い明細.xls");
        for (var report : reports) {
            try (Workbook wb = WorkbookFactory.create(new ByteArrayInputStream(report.data()))) {
                assertThat(wb.getNumberOfSheets()).isEqualTo(1);
            }
        }
    }

    @Test
    void splitsBothReportsByTorihikisaki() throws Exception {
        var reports = service.generateReports("202601", "年払", List.of(
                row("1", "T001", "取引先A", "10000"),
                row("1", "T001", "取引先A", "20000"),
                row("2", "T002", "取引先B", "30000")));
        for (var report : reports) {
            try (Workbook wb = WorkbookFactory.create(new ByteArrayInputStream(report.data()))) {
                assertThat(wb.getNumberOfSheets()).isEqualTo(2);
            }
        }
    }

    /** 対象データが空の場合は例外（VB版 MSG_0002 相当、Controller側でメッセージ表示）。 */
    @Test
    void emptyRowsThrows() {
        assertThatThrownBy(() -> service.generateReports("202601", "年払", List.of()))
                .isInstanceOf(IllegalArgumentException.class);
    }

    /** 支払月フォーマットが不正な場合は例外。 */
    @Test
    void invalidMonthFormatThrows() {
        List<RowForm> rows = List.of(row("1", "T001", "取引先A", "10000"));
        assertThatThrownBy(() -> service.generateReports("2026-01", "年払", rows))
                .isInstanceOf(IllegalArgumentException.class);
    }

    /** 支払区分が不正な場合は例外。 */
    @Test
    void invalidPaymentTypeThrows() {
        List<RowForm> rows = List.of(row("1", "T001", "取引先A", "10000"));
        assertThatThrownBy(() -> service.generateReports("202601", "半期払", rows))
                .isInstanceOf(IllegalArgumentException.class);
    }

    /** 取引先IDまたは支払金額が欠落している場合は例外（帳票出力の前提条件）。 */
    @Test
    void missingAmountThrows() {
        var row = row("1", "T001", "取引先A", "10000");
        row.setShiharaiKingaku(null);
        assertThatThrownBy(() -> service.generateReports("202601", "年払", List.of(row)))
                .isInstanceOf(IllegalArgumentException.class);
    }

    private Workbook report(List<RowForm> rows, boolean order) throws Exception {
        var reports = service.generateReports("202601", "月払", rows);
        return WorkbookFactory.create(new ByteArrayInputStream(reports.get(order ? 0 : 1).data()));
    }

    /** VBの指定セルへ全項目を配置し、コードの先頭ゼロと文字列を維持する。 */
    @Test
    void orderHeaderAndDetailsFollowVbPositions() throws Exception {
        var r = row("1", "T001", "取引先A", "1234.5");
        r.setKeiyakuNo("000123"); r.setHardSeiban("000456");
        try (var w = report(List.of(r), true)) {
            Sheet s = w.getSheetAt(0);
            assertThat(s.getRow(1).getCell(1).getStringCellValue()).isEqualTo("取引先A 担当者 様");
            assertThat(s.getRow(4).getCell(3).getStringCellValue()).isEqualTo("2026年01");
            assertThat(s.getRow(4).getCell(7).getStringCellValue()).isEqualTo("( 月払 )");
            assertThat(s.getRow(15).getCell(1).getNumericCellValue()).isEqualTo(1);
            assertThat(s.getRow(15).getCell(2).getStringCellValue()).isEqualTo("000123");
            assertThat(s.getRow(15).getCell(3).getNumericCellValue()).isEqualTo(2026);
            assertThat(s.getRow(15).getCell(4).getNumericCellValue()).isEqualTo(1);
            assertThat(s.getRow(15).getCell(5).getStringCellValue()).isEqualTo("新規");
            assertThat(s.getRow(15).getCell(6).getStringCellValue()).isEqualTo("納入先");
            assertThat(s.getRow(15).getCell(8).getNumericCellValue()).isEqualTo(1234.5);
            assertThat(s.getRow(15).getCell(9).getStringCellValue()).isEqualTo("000456");
        }
    }

    /** VBのSUM数式を再現し、Excelで金額を書き換えた後も合計が追従する。 */
    @Test
    void orderTotalRecalculatesAfterEditingDetail() throws Exception {
        try (var w = report(List.of(row("1", "T1", "A", "100"), row("1", "T1", "A", "200")), true)) {
            Sheet s = w.getSheetAt(0);
            var total = s.getRow(17).getCell(8);
            assertThat(total.getCellType()).isEqualTo(CellType.FORMULA);
            assertThat(total.getCellFormula()).isEqualTo("SUM(I16:I17)");
            assertThat(total.getNumericCellValue()).isEqualTo(300);
            s.getRow(15).getCell(8).setCellValue(400);
            w.getCreationHelper().createFormulaEvaluator().evaluateAll();
            assertThat(total.getNumericCellValue()).isEqualTo(600);
            assertThat(new DataFormatter().formatCellValue(total,
                    w.getCreationHelper().createFormulaEvaluator())).contains("¥");
        }
    }

    /** 2009年VB修正後は金額0・製番なしの明細も出力対象。 */
    @Test
    void zeroAmountAndMissingSerialAreIncluded() throws Exception {
        var r = row("1", "T1", "A", "0"); r.setHardSeiban(null);
        try (var w = report(List.of(r), true)) {
            assertThat(w.getSheetAt(0).getRow(15).getCell(8).getNumericCellValue()).isZero();
            assertThat(w.getSheetAt(0).getRow(15).getCell(9).getStringCellValue()).isEmpty();
            assertThat(w.getSheetAt(0).getRow(16).getCell(7).getStringCellValue()).isEqualTo("合計");
        }
    }

    @Test
    void idsWithDifferentDecimalScalesBelongToSameSupplier() throws Exception {
        try (var w = report(List.of(row("1", "T1", "A", "10"), row("1.00", "T1", "A", "20")), true)) {
            assertThat(w.getNumberOfSheets()).isEqualTo(1);
            assertThat(w.getSheetAt(0).getRow(17).getCell(8).getNumericCellValue()).isEqualTo(30);
        }
    }

    @Test
    void supplierSheetsHaveIndependentDetailsAndTotals() throws Exception {
        try (var w = report(List.of(row("1", "T1", "A", "10"), row("2", "T2", "B", "20")), true)) {
            assertThat(w.getSheetName(0)).isEqualTo("T1");
            assertThat(w.getSheetName(1)).isEqualTo("T2");
            assertThat(w.getSheetAt(0).getRow(16).getCell(8).getNumericCellValue()).isEqualTo(10);
            assertThat(w.getSheetAt(1).getRow(16).getCell(8).getNumericCellValue()).isEqualTo(20);
            assertThat(w.getPrintArea(1)).isEqualTo("'T2'!$A$1:$J$17");
        }
    }

    @Test
    void duplicateAndUnsafeSheetNamesStillProduceAllSuppliers() throws Exception {
        try (var w = report(List.of(row("1", "A/B", "A", "10"), row("2", "A/B", "B", "20"),
                row("3", "", "C", "30")), true)) {
            assertThat(w.getNumberOfSheets()).isEqualTo(3);
            assertThat(w.getSheetName(0)).isEqualTo("A B");
            assertThat(w.getSheetName(1)).isEqualTo("A B_2");
            assertThat(w.getSheetName(2)).isEqualTo("取引先");
        }
    }

    @Test
    void templatesKeepColumnWidthsFontsAndPageSettings() throws Exception {
        try (var w = report(List.of(row("1", "T1", "A", "100")), true);
             var in = getClass().getResourceAsStream("/templates/excel/MCM3001P_発注確認書テンプレート.xls");
             var template = WorkbookFactory.create(in)) {
            Sheet s = w.getSheetAt(0), original = template.getSheetAt(0);
            for (int c = 0; c <= 9; c++) assertThat(s.getColumnWidth(c)).isEqualTo(original.getColumnWidth(c));
            assertThat(s.getPrintSetup().getPaperSize()).isEqualTo(original.getPrintSetup().getPaperSize());
            assertThat(s.getPrintSetup().getLandscape()).isEqualTo(original.getPrintSetup().getLandscape());
            assertThat(s.getPrintSetup().getScale()).isEqualTo(original.getPrintSetup().getScale());
            assertThat(s.getPrintSetup().getFitWidth()).isEqualTo(original.getPrintSetup().getFitWidth());
            assertThat(s.getPrintSetup().getFitHeight()).isEqualTo(original.getPrintSetup().getFitHeight());
            assertThat(s.getMargin(Sheet.LeftMargin)).isEqualTo(original.getMargin(Sheet.LeftMargin));
            assertThat(s.getMargin(Sheet.RightMargin)).isEqualTo(original.getMargin(Sheet.RightMargin));
            assertThat(s.getMargin(Sheet.TopMargin)).isEqualTo(original.getMargin(Sheet.TopMargin));
            assertThat(s.getMargin(Sheet.BottomMargin)).isEqualTo(original.getMargin(Sheet.BottomMargin));
            assertThat(s.getMergedRegions()).isEqualTo(original.getMergedRegions());
            assertThat(s.getRepeatingRows()).isEqualTo(new CellRangeAddress(0, 14, -1, -1));
            var actualFont = w.getFontAt(s.getRow(15).getCell(8).getCellStyle().getFontIndex());
            var expectedFont = template.getFontAt(original.getRow(15).getCell(8).getCellStyle().getFontIndex());
            assertThat(actualFont.getFontName()).isEqualTo(expectedFont.getFontName());
            assertThat(actualFont.getFontHeight()).isEqualTo(expectedFont.getFontHeight());
        }
    }

    @Test
    void extendedDetailsKeepYenAndVbBorders() throws Exception {
        try (var w = report(Collections.nCopies(100, row("1", "T1", "A", "123")), true)) {
            Sheet s = w.getSheetAt(0);
            assertThat(s.getRow(114).getCell(1).getNumericCellValue()).isEqualTo(100);
            assertThat(s.getRow(114).getHeight()).isEqualTo(s.getRow(15).getHeight());
            assertThat(s.getRow(114).getCell(8).getCellStyle().getDataFormatString()).contains("¥");
            assertThat(s.getRow(114).getCell(6).getCellStyle().getBorderRight()).isEqualTo(BorderStyle.NONE);
            assertThat(s.getRow(114).getCell(7).getCellStyle().getBorderLeft()).isEqualTo(BorderStyle.NONE);
            assertThat(s.getRow(15).getCell(1).getCellStyle().getBorderTop()).isEqualTo(BorderStyle.THIN);
            assertThat(s.getRow(114).getCell(1).getCellStyle().getBorderBottom()).isEqualTo(BorderStyle.HAIR);
            assertThat(s.getRow(115).getCell(8).getCellStyle().getBorderBottom()).isEqualTo(BorderStyle.THIN);
            assertThat(s.getRow(115).getCell(8).getCellStyle().getFillForegroundColor())
                    .isEqualTo(IndexedColors.GREY_25_PERCENT.getIndex());
            assertThat(s.getRow(115).getCell(8).getNumericCellValue()).isEqualTo(12300);
            assertThat(w.getPrintArea(0)).isEqualTo("'T1'!$A$1:$J$116");
        }
    }

    @Test
    void singleDetailPrintsOnlyThroughTotal() throws Exception {
        try (var w = report(List.of(row("1", "T1", "A", "10")), true)) {
            assertThat(w.getPrintArea(0)).isEqualTo("'T1'!$A$1:$J$17");
        }
    }

    /** 同時出力の金額はVBのCIntと同じ銀行丸め、合計は¥表示の数式。 */
    @Test
    void paymentTotalRecalculatesAndKeepsYenAfterOriginalRowsRunOut() throws Exception {
        try (var w = report(Collections.nCopies(120, row("1", "T1", "A", "1234.5")), false)) {
            Sheet s = w.getSheetAt(0);
            assertThat(s.getRow(9).getCell(5).getNumericCellValue()).isEqualTo(1234);
            var total = s.getRow(131).getCell(5);
            assertThat(total.getCellFormula()).isEqualTo("SUM(F10:F129)");
            assertThat(total.getNumericCellValue()).isEqualTo(148080);
            assertThat(total.getCellStyle().getDataFormatString()).contains("¥");
            assertThat(s.getRow(131).getCell(6).getCellStyle().getDataFormatString()).contains("¥");
            assertThat(s.getRow(133).getCell(5).getStringCellValue()).isEqualTo("120件");
            assertThat(s.getRow(133).getCell(6).getStringCellValue()).isEqualTo("0件");
            assertThat(s.getRow(133).getCell(6).getCellStyle().getDataFormatString()).isEqualTo("General");
            s.getRow(9).getCell(5).setCellValue(2000);
            s.getRow(9).getCell(6).setCellValue(50);
            w.getCreationHelper().createFormulaEvaluator().evaluateAll();
            assertThat(total.getNumericCellValue()).isEqualTo(148846);
            assertThat(s.getRow(131).getCell(6).getNumericCellValue()).isEqualTo(50);
            assertThat(s.getRow(133).getCell(6).getStringCellValue()).isEqualTo("1件");
            assertThat(w.getPrintArea(0)).isEqualTo("'T1'!$A$1:$I$134");
        }
    }

    @Test
    void sixThousandDetailsProduceCompleteReportWithoutExcessStyles() throws Exception {
        try (var w = report(Collections.nCopies(6000, row("1", "T1", "A", "10")), true)) {
            assertThat(w.getSheetAt(0).getRow(6014).getCell(1).getNumericCellValue()).isEqualTo(6000);
            assertThat(w.getSheetAt(0).getRow(6015).getCell(8).getNumericCellValue()).isEqualTo(60000);
            assertThat(w.getNumCellStyles()).isLessThan(200);
        }
    }

    @Test
    void rowLimitIsRejectedBeforeGeneratingWorkbooks() {
        assertThatThrownBy(() -> service.generateReports("202601", "年払",
                Collections.nCopies(65521, row("1", "T1", "A", "1"))))
                .isInstanceOf(IllegalArgumentException.class).hasMessageContaining("行数上限");
    }

    @Test
    void invalidYearOrCalendarMonthIsReportedAsInputError() {
        var rows = List.of(row("1", "T1", "A", "1"));
        for (String month : List.of("000001", "202600", "202613")) {
            assertThatThrownBy(() -> service.generateReports(month, "年払", rows))
                    .isInstanceOf(IllegalArgumentException.class).hasMessageContaining("支払い月");
        }
    }

    @Test
    void nullDetailIsRejectedWithGuidance() {
        assertThatThrownBy(() -> service.generateReports("202601", "年払", Collections.singletonList(null)))
                .isInstanceOf(IllegalArgumentException.class).hasMessageContaining("取引先または支払金額");
    }
}
