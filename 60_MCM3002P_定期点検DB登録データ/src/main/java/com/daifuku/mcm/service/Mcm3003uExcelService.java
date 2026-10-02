package com.daifuku.mcm.service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import com.daifuku.mcm.constants.Mcm3002pConstants;
import com.daifuku.mcm.constants.Mcm3003pConstants;
import com.daifuku.mcm.form.Mcm3003uForm.RowForm;

/**
 * 【変換元】Mcm3002pExcel.vb（定期点検DB登録データ）/ Mcm3003pExcel.vb（定期点検リスト）
 *          CPExcelManager.vb setBorders / Mcm3002pData.vb / Mcm3003pData.vb
 * MCM3003U 保守点検出力指示 Excel出力サービス
 *
 * MCM_作業用/50_帳票テンプレート の原紙（.xls / HSSF）を
 * src/main/resources/templates/excel へ取り込み、書式を保持したまま
 * 2行目以降へ明細を流し込む。原紙の見出し・列幅・印刷設定・表示書式は一切変更しない。
 *
 * 原紙の構造（実ファイルをPOIで確認した結果）:
 *   MCM3002P: シート1枚「定期点検DB登録データ」/ 52列 / 縦・A4・縮小82%・横1ページ
 *             1列目「５件以上フラグ」、2列目「点検日」（書式 yyyy"年"m"月";@）、
 *             12〜51列目に8列ブロック×5組、52列目「※以下続く」
 *   MCM3003P: シート1枚「定期点検リスト」/ 10列 / 横・A4・横1ページ
 *             10列目「契約点検月」（書式 mm"月"）、9列目「回数」は右寄せ
 *             11〜15列目はVB側の後付け項目で原紙に見出し・書式が無い
 *
 * 原紙は 58行目（index 57）までしか書式が用意されていないため、
 * 明細先頭行（index 1）のセル書式と行高を退避し、以降の行へ複写する。
 */
@Service
public class Mcm3003uExcelService extends com.daifuku.mcm.common.BaseExcelService {

    /** テンプレート原紙の配置ディレクトリ（クラスパス） */
    private static final String TEMPLATE_DIR = "templates/excel/";

    /** VB版は rowIndex=1 から開始し最初の書き込み前に +1 するため、明細はExcelの2行目（POIでは index 1）から。 */
    private static final int DATA_START_ROW = 1;

    /** 原紙が .xls（HSSF）形式のため行数上限は EXCEL97 準拠。 */
    private static final int MAX_ROW = SpreadsheetVersion.EXCEL97.getMaxRows();

    // ===================================================================
    // ＤＢ登録用データ（MCM3002P）
    // 【変換元】Mcm3002pExcel.vb WriteDataInExcel
    // ===================================================================

    /**
     * 定期点検DB登録データを生成する。
     *
     * VB版のロジック:
     *   プラントID・点検月・点検可能曜日が変化した時、または 1行に 5 件展開した時に行を折り返す。
     *   折り返した行に固定項目（点検月〜納入先電話番号）を書き、
     *   機器ブロック（ブランド名〜取引先／8列）を横方向に最大5件並べる。
     *   5件目を書いた行は 1列目に "●"（＝5件以上あることを示す）を立てる。
     */
    public byte[] generateDbTorokuyo(List<RowForm> tenkenList) throws IOException {
        requireRows(tenkenList);
        // .xlsの上限超過を、全セルを作成する前に検知する。
        checkDbRowLimit(tenkenList);

        int lastCol = Mcm3002pConstants.COL_TORIHIKISAKI_NK
                + (Mcm3002pConstants.MAX_BLOCK_PER_ROW - 1) * Mcm3002pConstants.COLINDEX_COUNT;

        try (Workbook book = openTemplate(Mcm3002pConstants.TEMPLATE_FILE_NAME);
             var out = new ByteArrayOutputStream()) {

            Sheet sheet = book.getSheetAt(0);
            RowTemplate template = RowTemplate.capture(sheet, DATA_START_ROW, lastCol);

            // VB: rowIndex は 1 始まり。Java では 0 始まりへ読み替える。
            int rowIndex = DATA_START_ROW - 1;
            int colCount = 0;
            boolean firstRow = true;

            LocalDate tsukiBack = null;
            BigDecimal plantIdBack = null;
            String tenkenKanoYobiBack = null;

            for (RowForm data : tenkenList) {
                LocalDate tsuki = data.getTsuki();
                BigDecimal plantId = data.getPlantId();
                String tenkenKanoYobi = data.getTenkenkanoyobi();

                boolean breakRow = firstRow
                        || colCount == Mcm3002pConstants.MAX_BLOCK_PER_ROW - 1
                        || !Objects.equals(tsuki, tsukiBack)
                        || !equalsDecimal(plantId, plantIdBack)
                        || !Objects.equals(tenkenKanoYobi, tenkenKanoYobiBack);

                if (breakRow) {
                    rowIndex++;
                    checkRowLimit(rowIndex);
                    colCount = 0;
                    firstRow = false;
                    plantIdBack = plantId;
                    tsukiBack = tsuki;
                    tenkenKanoYobiBack = tenkenKanoYobi;

                    template.apply(sheet, rowIndex);

                    // 点検月は原紙の書式（yyyy"年"m"月"）を活かすため日付値として書き込む
                    putDate(sheet, rowIndex, Mcm3002pConstants.COL_TSUKI, tsuki);
                    // アポイントはVB版のSELECTに含まれないため常に空欄（Mcm3002pData.APPOINT が未設定）
                    put(sheet, rowIndex, Mcm3002pConstants.COL_APPOINT, "");
                    put(sheet, rowIndex, Mcm3002pConstants.COL_NONYUSAKI_NK, data.getNonyusakiNk());
                    put(sheet, rowIndex, Mcm3002pConstants.COL_NONYUSAKIKOJO_NK, data.getNonyusakikojoNk());
                    put(sheet, rowIndex, Mcm3002pConstants.COL_PLANT_NK, data.getPlantNk());
                    put(sheet, rowIndex, Mcm3002pConstants.COL_MEISHO4_NK, data.getMeisho4Nk());
                    put(sheet, rowIndex, Mcm3002pConstants.COL_SUPPORT_ID, formatSupportId(data.getSupportId()));
                    put(sheet, rowIndex, Mcm3002pConstants.COL_NONYUTANTOSYA_NK, data.getNonyutantosyaNk());
                    put(sheet, rowIndex, Mcm3002pConstants.COL_NONYUBUSYO_NK, data.getNonyubusyoNk());
                    put(sheet, rowIndex, Mcm3002pConstants.COL_NONYUTEL_NO, data.getNonyutelNo());
                } else {
                    colCount++;
                }

                int colPlus = colCount * Mcm3002pConstants.COLINDEX_COUNT;
                put(sheet, rowIndex, Mcm3002pConstants.COL_BRAND_NK + colPlus, data.getBrandNk());
                put(sheet, rowIndex, Mcm3002pConstants.COL_KIKIHINMEI_NK + colPlus, data.getKikihinmeiNk());
                put(sheet, rowIndex, Mcm3002pConstants.COL_HOSHUGAISHA_NK + colPlus, data.getHoshugaishaNk());
                put(sheet, rowIndex, Mcm3002pConstants.COL_HOSHUGAISHAJIGYOSYO_NK + colPlus, data.getHoshugaishajigyosyoNk());
                put(sheet, rowIndex, Mcm3002pConstants.COL_HOSHUGAISHATANTOSYA_NK + colPlus, data.getHoshugaishatantosyaNk());
                put(sheet, rowIndex, Mcm3002pConstants.COL_HOSHUGAISHATEL_NO + colPlus, data.getHoshugaishatelNo());
                put(sheet, rowIndex, Mcm3002pConstants.COL_HOSHUGAISHAFAX_NO + colPlus, data.getHoshugaishafaxNo());
                put(sheet, rowIndex, Mcm3002pConstants.COL_TORIHIKISAKI_NK + colPlus, data.getTorihikisakiNk());

                if (colCount == Mcm3002pConstants.MAX_BLOCK_PER_ROW - 1) {
                    put(sheet, rowIndex, Mcm3002pConstants.COL_5KENFLG, Mcm3002pConstants.FLG_5KEN);
                }
            }

            book.write(out);
            return out.toByteArray();
        }
    }

    // ===================================================================
    // 作業用リスト（MCM3003P）
    // 【変換元】Mcm3003pExcel.vb WriteDataInExcel
    // ===================================================================

    /**
     * 定期点検リスト（作業用リスト）を生成する。
     *
     * VB版のロジック:
     *   1件を1行に出力する。契約日・解約日は値が入っている場合のみ書き込む
     *   （VB: If tenkenData.KEIYAKU_DT &gt; "#12:00:00 AM#"）。
     *   最後に 2行目〜最終行、店舗名列〜備考列に罫線を引く（CPExcelManager.setBorders 相当）。
     */
    public byte[] generateSagyoyoList(List<RowForm> tenkenList) throws IOException {
        requireRows(tenkenList);
        checkRowLimit(DATA_START_ROW + tenkenList.size() - 1);

        try (Workbook book = openTemplate(Mcm3003pConstants.TEMPLATE_FILE_NAME);
             var out = new ByteArrayOutputStream()) {

            Sheet sheet = book.getSheetAt(0);
            RowTemplate template = RowTemplate.capture(sheet, DATA_START_ROW, Mcm3003pConstants.COL_BIKO);

            // 契約日・解約日は原紙に書式が無いため、日付書式を明示的に付与する。
            // （VB版はExcel Interopが日付値に既定の日付書式を自動適用するが、POIは適用しない）
            CellStyle dateStyle = createDateStyle(book, template.style(Mcm3003pConstants.COL_KEIYAKU_DT));

            int rowIndex = DATA_START_ROW - 1;
            for (RowForm data : tenkenList) {
                rowIndex++;
                template.apply(sheet, rowIndex);

                put(sheet, rowIndex, Mcm3003pConstants.COL_MEISHO4_NK, data.getMeisho4Nk());
                put(sheet, rowIndex, Mcm3003pConstants.COL_NONYUSAKI_NK, data.getNonyusakiNk());
                put(sheet, rowIndex, Mcm3003pConstants.COL_NONYUSAKIKANA_KN, data.getNonyusakikanaKn());
                put(sheet, rowIndex, Mcm3003pConstants.COL_SUPPORT_ID, formatSupportId(data.getSupportId()));
                put(sheet, rowIndex, Mcm3003pConstants.COL_PLANT_NK, data.getPlantNk());
                put(sheet, rowIndex, Mcm3003pConstants.COL_BRAND_NK, data.getBrandNk());
                put(sheet, rowIndex, Mcm3003pConstants.COL_KIKIHINMEI_NK, data.getKikihinmeiNk());
                put(sheet, rowIndex, Mcm3003pConstants.COL_TORIHIKISAKI_NK, data.getTorihikisakiNk());
                put(sheet, rowIndex, Mcm3003pConstants.COL_TENKENKAISU, data.getTenkenkaisu());
                // 点検月は原紙の書式（mm"月"）を活かすため日付値として書き込む
                putDate(sheet, rowIndex, Mcm3003pConstants.COL_TSUKI, data.getTsuki());
                // VB: 契約日・解約日は値が存在する場合のみ書き込む
                if (data.getKeiyakuDt() != null) {
                    putDate(sheet, rowIndex, Mcm3003pConstants.COL_KEIYAKU_DT, data.getKeiyakuDt());
                    cell(sheet, rowIndex, Mcm3003pConstants.COL_KEIYAKU_DT).setCellStyle(dateStyle);
                }
                if (data.getKaiyakuDt() != null) {
                    putDate(sheet, rowIndex, Mcm3003pConstants.COL_KAIYAKU_DT, data.getKaiyakuDt());
                    cell(sheet, rowIndex, Mcm3003pConstants.COL_KAIYAKU_DT).setCellStyle(dateStyle);
                }
                put(sheet, rowIndex, Mcm3003pConstants.COL_KEIYAKU_NO, data.getKeiyakuNo());
                put(sheet, rowIndex, Mcm3003pConstants.COL_HOSHUGAISHA_NK, data.getHoshugaishaNk());
                put(sheet, rowIndex, Mcm3003pConstants.COL_BIKO, data.getBiko());
            }

            // 【変換元】Mcm3003pExcel.vb: setBorders(sheet, 2, COL_MEISHO4_NK, rowIndex, COL_BIKO, 縦線あり)
            setBorders(book, sheet, DATA_START_ROW, rowIndex,
                       Mcm3003pConstants.COL_MEISHO4_NK, Mcm3003pConstants.COL_BIKO);

            book.write(out);
            return out.toByteArray();
        }
    }

    // ===================================================================
    // 共通処理
    // ===================================================================

    /** テンプレート原紙をクラスパスから読み込む。シート構成が想定外なら失敗させる。 */
    @Override
    protected Workbook openTemplate(String fileName) throws IOException {
        try (var in = new ClassPathResource(TEMPLATE_DIR + fileName).getInputStream()) {
            Workbook book = WorkbookFactory.create(in);
            if (book.getNumberOfSheets() != 1) {
                book.close();
                throw new IOException("帳票テンプレートのシート構成が不正です: " + fileName);
            }
            return book;
        }
    }

    /**
     * 明細先頭行の書式・行高を退避し、任意の行へ複写するための保持クラス。
     * 原紙は58行目までしか書式が用意されていないため、それ以降の行にも同じ体裁を与える。
     */
    private record RowTemplate(CellStyle[] styles, short height) {

        static RowTemplate capture(Sheet sheet, int rowIndex, int lastVbColumn) {
            CellStyle[] styles = new CellStyle[lastVbColumn];
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                for (int c = 0; c < lastVbColumn; c++) {
                    Cell cell = row.getCell(c);
                    if (cell != null) {
                        styles[c] = cell.getCellStyle();
                    }
                }
            }
            short height = row == null ? sheet.getDefaultRowHeight() : row.getHeight();
            return new RowTemplate(styles, height);
        }

        /** VB列番号（1始まり）に対応する退避済み書式。無ければ null。 */
        CellStyle style(int vbColumn) {
            int index = vbColumn - 1;
            return index >= 0 && index < styles.length ? styles[index] : null;
        }

        void apply(Sheet sheet, int rowIndex) {
            for (int c = 0; c < styles.length; c++) {
                if (styles[c] == null) {
                    continue;
                }
                cell(sheet, rowIndex, c + 1).setCellStyle(styles[c]);
            }
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                row.setHeight(height);
            }
        }
    }

    /**
     * サポートIDを「Pxx-yyyyy」形式へ変換する。
     * 【変換元】Mcm3002pExcel.vb / Mcm3003pExcel.vb setSupportId
     *   supportId を '0' で7桁まで右詰めし、"P" + 先頭2桁 + "-" + 続く5桁を返す。
     */
    static String formatSupportId(String supportId) {
        if (supportId == null || supportId.isBlank()) {
            return "";
        }
        StringBuilder padded = new StringBuilder(supportId);
        while (padded.length() < 7) {
            padded.append('0');
        }
        return "P" + padded.substring(0, 2) + "-" + padded.substring(2, 7);
    }

    /**
     * 【変換元】CPExcelManager.vb setBorders（VerticalLineStyle_Ari）
     *   外枠は細線、内側の縦線は細線、内側の横線は極細。
     */
    private void setBorders(Workbook book, Sheet sheet, int firstRow, int lastRow,
                            int firstVbCol, int lastVbCol) {
        if (lastRow < firstRow) {
            return;
        }
        Map<String, CellStyle> cache = new HashMap<>();
        for (int r = firstRow; r <= lastRow; r++) {
            for (int c = firstVbCol; c <= lastVbCol; c++) {
                Cell target = cell(sheet, r, c);
                boolean top = r == firstRow;
                boolean bottom = r == lastRow;
                String key = target.getCellStyle().getIndex() + ":" + top + ":" + bottom;
                CellStyle style = cache.get(key);
                if (style == null) {
                    style = book.createCellStyle();
                    style.cloneStyleFrom(target.getCellStyle());
                    style.setBorderTop(top ? BorderStyle.THIN : BorderStyle.HAIR);
                    style.setBorderBottom(bottom ? BorderStyle.THIN : BorderStyle.HAIR);
                    style.setBorderLeft(BorderStyle.THIN);
                    style.setBorderRight(BorderStyle.THIN);
                    cache.put(key, style);
                }
                target.setCellStyle(style);
            }
        }
    }

    private void checkDbRowLimit(List<RowForm> rows) {
        int outputRows = 0;
        int blocks = Mcm3002pConstants.MAX_BLOCK_PER_ROW;
        RowForm previous = null;
        for (RowForm data : rows) {
            if (data == null) throw new IllegalStateException("点検データに不正な行があります。");
            if (blocks == Mcm3002pConstants.MAX_BLOCK_PER_ROW || previous == null
                    || !Objects.equals(data.getTsuki(), previous.getTsuki())
                    || !equalsDecimal(data.getPlantId(), previous.getPlantId())
                    || !Objects.equals(data.getTenkenkanoyobi(), previous.getTenkenkanoyobi())) {
                outputRows++;
                checkRowLimit(DATA_START_ROW + outputRows - 1);
                blocks = 0;
            }
            blocks++;
            previous = data;
        }
    }

    private void requireRows(List<RowForm> rows) {
        // 【変換元】MSG_0002E「検索結果が1件も存在しません。」
        if (rows == null || rows.isEmpty()) {
            throw new IllegalArgumentException("検索結果が1件も存在しません。");
        }
    }

    private void checkRowLimit(int rowIndex) {
        if (rowIndex >= MAX_ROW) {
            throw new IllegalArgumentException("出力件数がExcelの行数上限を超えています。点検対象月の範囲を狭めてください。");
        }
    }

    private static boolean equalsDecimal(BigDecimal a, BigDecimal b) {
        if (a == null || b == null) {
            return a == b;
        }
        return a.compareTo(b) == 0;
    }

    /** 原紙の書式を引き継いだ上で日付書式（yyyy/mm/dd）を付与したスタイルを作る。 */
    private CellStyle createDateStyle(Workbook book, CellStyle base) {
        CellStyle style = book.createCellStyle();
        if (base != null) {
            style.cloneStyleFrom(base);
        }
        style.setDataFormat(book.getCreationHelper().createDataFormat().getFormat("yyyy/mm/dd"));
        return style;
    }

    /** VB列番号（1始まり）で文字列を書き込む。書式は既存のまま維持する。 */
    private static void put(Sheet sheet, int rowIndex, int vbColumn, String value) {
        cell(sheet, rowIndex, vbColumn).setCellValue(value == null ? "" : value);
    }

    /** VB列番号（1始まり）で日付値を書き込む。表示書式は原紙のものを使う。 */
    private static void putDate(Sheet sheet, int rowIndex, int vbColumn, LocalDate value) {
        Cell target = cell(sheet, rowIndex, vbColumn);
        if (value == null) {
            target.setBlank();
            return;
        }
        target.setCellValue(value);
    }

    /** VB列番号（1始まり）→ POI列インデックス（0始まり）でセルを取得・生成する。 */
    private static Cell cell(Sheet sheet, int rowIndex, int vbColumn) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        int colIndex = vbColumn - 1;
        Cell target = row.getCell(colIndex);
        return target == null ? row.createCell(colIndex) : target;
    }
}
