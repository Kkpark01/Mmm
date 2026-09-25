package com.daifuku.mcm.service;

import com.daifuku.mcm.constants.Mcm1003pConstants;
import com.daifuku.mcm.dto.Mcm1003pKihonDto;
import com.daifuku.mcm.dto.Mcm1003pMeisaiDto;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

/**
 * 【変換元】Mcm1003pExcel.vb - MCM1003P 契約手続依頼書 Excel出力サービス
 * 【変換元行数】約350行
 * 【作成者】T.Matsui (2009/03/04)
 *
 * 元コード: CPExcelManager を継承し、Microsoft.Office.Interop.Excel で出力。
 *   Initialize() → TemplateFileName, CreateFileName セット
 *   WriteDataInExcel() → ヘッダー・顧客情報・契約・解約・明細を書込
 *   setTantoIn() → Excelシェイプに担当者印を設定
 * 変換後:   Apache POI (HSSFWorkbook) でExcel出力。
 *
 * 【複数シート対応】
 *   1シート目: ヘッダー + 明細（最大20行）
 *   2シート目以降: 明細のみ（最大50行/シート）
 *   テンプレートは2シート構成（index 0=ヘッダー付き, index 1=明細のみ）
 */
@Service
public class Mcm1003pExcelService {

    private static final Logger log = LoggerFactory.getLogger(Mcm1003pExcelService.class);

    /**
     * 契約手続依頼書をExcel出力する。
     *
     * 【変換元】Mcm1003pExcel.vb - WriteDataInExcel()
     *
     * @param kihonData      基本情報DTO
     * @param meisaiList     機器明細リスト
     * @param loginUserLastName ログインユーザーの苗字（担当者印用）
     * @return Excelファイルのバイト配列
     * @throws IOException テンプレート読込・書込エラー時
     */
    public byte[] generate(Mcm1003pKihonDto kihonData,
                           List<Mcm1003pMeisaiDto> meisaiList,
                           String loginUserLastName) throws IOException {

        log.info("MCM1003P Excel生成開始: 明細件数={}", meisaiList.size());

        ClassPathResource templateResource =
                new ClassPathResource("templates/excel/" + Mcm1003pConstants.TEMPLATE_FILE_NAME);

        try (InputStream is = templateResource.getInputStream();
             Workbook workbook = WorkbookFactory.create(is);
             ByteArrayOutputStream bos = new ByteArrayOutputStream()) {

            /*
             * 【変換元】Mcm1003pExcel.vb - WriteDataInExcel()
             *   元コード: Dim pageSosu As Integer
             *             If uvaList2.Count > ROWCOUNT_MEISAIFST Then
             *               pageSosu = 1 + Math.Ceiling((uvaList2.Count - ROWCOUNT_MEISAIFST) / ROWCOUNT_MEISAISEC)
             *             Else
             *               pageSosu = 1
             *             End If
             */
            int meisaiCount = meisaiList.size();
            int pageSosu;
            if (meisaiCount > Mcm1003pConstants.ROWCOUNT_MEISAIFST) {
                pageSosu = 1 + (int) Math.ceil(
                        (double) (meisaiCount - Mcm1003pConstants.ROWCOUNT_MEISAIFST)
                                / Mcm1003pConstants.ROWCOUNT_MEISAISEC);
            } else {
                pageSosu = 1;
            }

            // ===================================================================
            // 1シート目（index 0）：ヘッダー + 顧客情報 + 契約・解約 + 明細
            // ===================================================================
            Sheet sheet1 = workbook.getSheetAt(0);

            // --- ヘッダー（区分・作成日・ページ番号） ---
            writeHeaderSection(sheet1, kihonData, pageSosu);

            // --- 顧客情報 ---
            writeCustomerSection(sheet1, kihonData);

            // --- 契約セクション（更新時はタイトル変更） ---
            writeContractSection(sheet1, kihonData);

            // --- 解約セクション ---
            writeCancellationSection(sheet1, kihonData);

            // --- 返金情報セクション ---
            writeRefundSection(sheet1, kihonData);

            // --- 担当者印 ---
            setTantoIn(sheet1, loginUserLastName);

            // --- 1シート目の明細（最大20行） ---
            int firstSheetMeisaiCount = Math.min(meisaiCount, Mcm1003pConstants.ROWCOUNT_MEISAIFST);
            writeMeisaiRows(sheet1,
                    meisaiList.subList(0, firstSheetMeisaiCount),
                    Mcm1003pConstants.STARTINDEX_MEISAIFST);

            // ===================================================================
            // 2シート目以降：明細のみ（最大50行/シート）
            // ===================================================================
            if (pageSosu > 1) {
                int remainingStart = Mcm1003pConstants.ROWCOUNT_MEISAIFST;
                for (int page = 3; page <= pageSosu; page++) workbook.cloneSheet(1);

                for (int page = 2; page <= pageSosu; page++) {
                    Sheet sheetN = workbook.getSheetAt(page - 1);
                    workbook.setPrintArea(page - 1, 0, 10, 0, 51);

                    // ページ番号設定
                    setCellValueByIndex(sheetN,
                            Mcm1003pConstants.MEISAISEC_ROWINDEX_PAGE_NO,
                            Mcm1003pConstants.MEISAISEC_COLINDEX_PAGE_NO,
                            "(" + page + "/" + pageSosu + ")");

                    // 明細行の書込
                    int endIdx = Math.min(remainingStart + Mcm1003pConstants.ROWCOUNT_MEISAISEC, meisaiCount);
                    writeMeisaiRows(sheetN,
                            meisaiList.subList(remainingStart, endIdx),
                            Mcm1003pConstants.STARTINDEX_MEISAISEC);

                    remainingStart = endIdx;
                }
            } else {
                /*
                 * 【変換元】Mcm1003pExcel.vb
                 *   元コード: If pageSosu = 1 Then objExcelApp.Worksheets(2).Delete
                 *   → 2シート目が不要な場合は削除
                 */
                if (workbook.getNumberOfSheets() > 1) {
                    workbook.removeSheetAt(1);
                }
            }

            workbook.write(bos);
            log.info("MCM1003P Excel生成完了: ページ数={}", pageSosu);
            return bos.toByteArray();
        }
    }

    // ===================================================================
    // ヘッダーセクション書込
    // ===================================================================

    /**
     * ヘッダー（区分・作成日・ページ番号）を書込む。
     *
     * 【変換元】Mcm1003pExcel.vb - WriteDataInExcel() 内ヘッダー部分
     */
    private void writeHeaderSection(Sheet sheet, Mcm1003pKihonDto data, int pageSosu) {
        // 区分
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_KUBUN,
                Mcm1003pConstants.KEIYAKU_COLINDEX_KUBUN,
                data.getKubun());

        /*
         * 元コード: sheet.Cells(ROWINDEX_CREATED_DT, COLINDEX_CREATED_DT).Value = "作成日：" & uvaData.CreatedDt
         */
        String createdDtDisplay = "作成日：" + (data.getCreatedDt() != null ? data.getCreatedDt() : "");
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_CREATED_DT,
                Mcm1003pConstants.KEIYAKU_COLINDEX_CREATED_DT,
                createdDtDisplay);

        // ページ番号
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_PAGE_NO,
                Mcm1003pConstants.KEIYAKU_COLINDEX_PAGE_NO,
                "(1/" + pageSosu + ")");
    }

    // ===================================================================
    // 顧客情報セクション書込
    // ===================================================================

    /**
     * 顧客情報（取引先・納入先・住所・電話・担当者等）を書込む。
     *
     * 【変換元】Mcm1003pExcel.vb - WriteDataInExcel() 内顧客情報部分
     */
    private void writeCustomerSection(Sheet sheet, Mcm1003pKihonDto data) {
        // 取引先名
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_TORIHIKISAKI_NK,
                Mcm1003pConstants.KEIYAKU_COLINDEX_TORIHIKISAKI_NK,
                data.getTorihikisakiNk());

        /*
         * 元コード: sheet.Cells(...).Value = uvaData.TorisyutantosyaNk & "様"
         */
        String tantosyaDisplay = honorific(data.getTorisyutantosyaNk());
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_TORISYUTANTOSYA_NK,
                Mcm1003pConstants.KEIYAKU_COLINDEX_TORISYUTANTOSYA_NK,
                tantosyaDisplay);

        // 納入先名
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_NONYUSAKI_NK,
                Mcm1003pConstants.KEIYAKU_COLINDEX_NONYUSAKI_NK,
                data.getNonyusakiNk());

        /*
         * 元コード:
         *   If uvaData.YubinNo <> "" Then
         *     sheet.Cells(...).Value = "〒" & Mid(uvaData.YubinNo, 1, 3) & "-" & Mid(uvaData.YubinNo, 4) & " " & uvaData.Nonyusakijusyo1Nk
         *   Else
         *     sheet.Cells(...).Value = uvaData.Nonyusakijusyo1Nk
         *   End If
         */
        String addressDisplay;
        if (data.getYubinNo() != null && !data.getYubinNo().isEmpty()) {
            String yubin = data.getYubinNo().replace("-", "");
            String formatted;
            if (yubin.length() >= 4) {
                formatted = "〒" + yubin.substring(0, 3) + "-" + yubin.substring(3);
            } else {
                formatted = "〒" + yubin;
            }
            addressDisplay = formatted + " " + (data.getNonyusakijusyo1Nk() != null ? data.getNonyusakijusyo1Nk() : "");
        } else {
            addressDisplay = data.getNonyusakijusyo1Nk();
        }
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_NONYUSAKIJUSYO1_NK,
                Mcm1003pConstants.KEIYAKU_COLINDEX_NONYUSAKIJUSYO1_NK,
                addressDisplay);

        // 電話番号
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_NONYUTEL_NO,
                Mcm1003pConstants.KEIYAKU_COLINDEX_NONYUTEL_NO,
                data.getNonyutelNo());

        // 担当部署
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_NONYUBUSYO_NK,
                Mcm1003pConstants.KEIYAKU_COLINDEX_NONYUBUSYO_NK,
                data.getNonyubusyoNk());

        /*
         * 元コード: sheet.Cells(...).Value = uvaData.NonyutantosyaNk & "様"
         */
        String nonyuTantosyaDisplay = honorific(data.getNonyutantosyaNk());
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_NONYUTANTOSYA_NK,
                Mcm1003pConstants.KEIYAKU_COLINDEX_NONYUTANTOSYA_NK,
                nonyuTantosyaDisplay);

        // サポートID
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_SUPPORT_ID,
                Mcm1003pConstants.KEIYAKU_COLINDEX_SUPPORT_ID,
                data.getSupportId());
    }

    // ===================================================================
    // 契約セクション書込
    // ===================================================================

    /**
     * 契約セクション（見積金額・支払月・契約開始日等）を書込む。
     *
     * 【変換元】Mcm1003pExcel.vb - WriteDataInExcel() 内契約部分
     *   元コード: If uvaData.Kubun = KUBUN_UPDATE Then
     *             sheet.Cells(ROWINDEX_TITLE, COLINDEX_TITLE).Value = TITLE_KEIYAKU
     */
    private void writeContractSection(Sheet sheet, Mcm1003pKihonDto data) {
        // 更新時はタイトルを「新契約」に変更
        if (Mcm1003pConstants.KUBUN_UPDATE.equals(data.getKubun())) {
            setCellValueByIndex(sheet,
                    Mcm1003pConstants.KEIYAKU_ROWINDEX_TITLE,
                    Mcm1003pConstants.KEIYAKU_COLINDEX_TITLE,
                    Mcm1003pConstants.TITLE_KEIYAKU);
        }

        // 見積金額（仕切）
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_SIKIRI_KEIYAKU,
                Mcm1003pConstants.KEIYAKU_COLINDEX_SIKIRI_KEIYAKU,
                data.getSikiriKeiyaku());

        // 支払月
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_SIHARAI_KEIYAKU,
                Mcm1003pConstants.KEIYAKU_COLINDEX_SIHARAI_KEIYAKU,
                data.getSiharaiKeiyaku());

        // 契約開始日
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_KEIYAKUKAISI_DT,
                Mcm1003pConstants.KEIYAKU_COLINDEX_KEIYAKUKAISI_DT,
                data.getKeiyakukaisiDt());

        // 見積No
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_TM_MITSUMORI_NO,
                Mcm1003pConstants.KEIYAKU_COLINDEX_TM_MITSUMORI_NO,
                data.getTmMitsumoriNo());

        // 契約時間帯
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_KEIYAKUJIKANTAI,
                Mcm1003pConstants.KEIYAKU_COLINDEX_KEIYAKUJIKANTAI,
                data.getKeiyakujikantai());

        // 手配製番
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_TEHAISEIBAN,
                Mcm1003pConstants.KEIYAKU_COLINDEX_TEHAISEIBAN,
                data.getTehaiseiban());

        // 備考（契約）
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KEIYAKU_ROWINDEX_BIKO_KEIYAKU,
                Mcm1003pConstants.KEIYAKU_COLINDEX_BIKO_KEIYAKU,
                data.getBikoKeiyaku());
    }

    // ===================================================================
    // 解約セクション書込
    // ===================================================================

    /**
     * 解約セクション（仕切・支払月・解約日・契約No等）を書込む。
     *
     * 【変換元】Mcm1003pExcel.vb - WriteDataInExcel() 内解約部分
     *   元コード: If uvaData.Kubun = KUBUN_UPDATE Then
     *             sheet.Cells(KAIYAKU_ROWINDEX_TITLE, KAIYAKU_COLINDEX_TITLE).Value = TITLE_KAIYAKU
     */
    private void writeCancellationSection(Sheet sheet, Mcm1003pKihonDto data) {
        // 更新時はタイトルを「旧契約」に変更
        if (Mcm1003pConstants.KUBUN_UPDATE.equals(data.getKubun())) {
            setCellValueByIndex(sheet,
                    Mcm1003pConstants.KAIYAKU_ROWINDEX_TITLE,
                    Mcm1003pConstants.KAIYAKU_COLINDEX_TITLE,
                    Mcm1003pConstants.TITLE_KAIYAKU);
        }

        // 見積金額（仕切・解約）
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KAIYAKU_ROWINDEX_SIKIRI_KAIYAKU,
                Mcm1003pConstants.KAIYAKU_COLINDEX_SIKIRI_KAIYAKU,
                data.getSikiriKaiyaku());

        // 支払月（解約）
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KAIYAKU_ROWINDEX_SIHARAI_KAIYAKU,
                Mcm1003pConstants.KAIYAKU_COLINDEX_SIHARAI_KAIYAKU,
                data.getSiharaiKaiyaku());

        // 解約日
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KAIYAKU_ROWINDEX_KAIYAKU_DT,
                Mcm1003pConstants.KAIYAKU_COLINDEX_KAIYAKU_DT,
                data.getKaiyakuDt());

        // 契約No
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KAIYAKU_ROWINDEX_KEIYAKU_NO,
                Mcm1003pConstants.KAIYAKU_COLINDEX_KEIYAKU_NO,
                data.getKeiyakuNo());

        // 備考（解約）
        setCellValueByIndex(sheet,
                Mcm1003pConstants.KAIYAKU_ROWINDEX_BIKO_KAIYAKU,
                Mcm1003pConstants.KAIYAKU_COLINDEX_BIKO_KAIYAKU,
                data.getBikoKaiyaku());
    }

    // ===================================================================
    // 返金情報セクション書込
    // ===================================================================

    /**
     * 返金情報（契約期間・返金期間・金額・備考）を書込む。
     *
     * 【変換元】Mcm1003pExcel.vb - WriteDataInExcel() 内返金情報部分
     */
    private void writeRefundSection(Sheet sheet, Mcm1003pKihonDto data) {
        // 契約期間
        setCellValueByIndex(sheet,
                Mcm1003pConstants.HENKIN_ROWINDEX_KEIYAKUKIKAN,
                Mcm1003pConstants.HENKIN_COLINDEX_KEIYAKUKIKAN,
                data.getKeiyakukikan());

        // 返金期間
        setCellValueByIndex(sheet,
                Mcm1003pConstants.HENKIN_ROWINDEX_HENKINKIKAN,
                Mcm1003pConstants.HENKIN_COLINDEX_HENKINKIKAN,
                data.getHenkinkikan());

        // 金額
        setCellValueByIndex(sheet,
                Mcm1003pConstants.HENKIN_ROWINDEX_KINGAKU,
                Mcm1003pConstants.HENKIN_COLINDEX_KINGAKU,
                data.getKingaku());

        // 備考（返金）
        setCellValueByIndex(sheet,
                Mcm1003pConstants.HENKIN_ROWINDEX_BIKO_HENKIN,
                Mcm1003pConstants.HENKIN_COLINDEX_BIKO_HENKIN,
                data.getBikoHenkin());
    }

    // ===================================================================
    // 担当者印
    // ===================================================================

    private static String honorific(String name) {
        return name == null || name.isBlank() ? "" : name + "様";
    }

    /** VB uses the named stamp shapes; leave approval stamps untouched. */
    private void setTantoIn(Sheet sheet, String name) {
        if (sheet instanceof HSSFSheet h && h.getDrawingPatriarch() != null)
            writeStamp(h.getDrawingPatriarch().getChildren(), name == null ? "" : name);
    }

    private void writeStamp(List<HSSFShape> shapes, String name) {
        for (HSSFShape shape : shapes) {
            if (shape instanceof HSSFShapeGroup group) writeStamp(group.getChildren(), name);
            String key = shape.getShapeName();
            if (key == null || !(shape instanceof HSSFSimpleShape text)) continue;
            key = key.replace("\u0000", "");
            String value = switch (key) {
                case "stp_lastname" -> "DTS";
                case "stp_firstname" -> name;
                case "stp_date" -> LocalDate.now().format(DateTimeFormatter.ofPattern("yy.MM.dd"));
                default -> null;
            };
            if (value != null) text.setString(new HSSFRichTextString(value));
        }
    }

    // ===================================================================
    // 明細行書込
    // ===================================================================

    /**
     * 明細行を書込む。
     *
     * 【変換元】Mcm1003pExcel.vb - WriteDataInExcel() 内明細ループ部分
     *   元コード:
     *     For j = 0 To maxRow - 1
     *       sheet.Cells(startIndex + j, COLINDEX_HYOJIJUN).Value = uvaList2(dataIndex).HYOJIJUN
     *       sheet.Cells(startIndex + j, COLINDEX_KIKIHINMEI_NK).Value = uvaList2(dataIndex).KIKIHINMEI_NK
     *       ...
     *       dataIndex += 1
     *     Next
     *
     * @param sheet      対象シート
     * @param meisaiList 書込む明細リスト
     * @param startIndex 開始行（1-based）
     */
    private void writeMeisaiRows(Sheet sheet, List<Mcm1003pMeisaiDto> meisaiList, int startIndex) {
        for (int j = 0; j < meisaiList.size(); j++) {
            Mcm1003pMeisaiDto meisai = meisaiList.get(j);
            int rowIdx = (startIndex - 1) + j; // 0-based変換

            Row row = sheet.getRow(rowIdx);
            if (row == null) {
                row = sheet.createRow(rowIdx);
            }

            // 表示順
            setCellValue(row, Mcm1003pConstants.MEISAI_COLINDEX_HYOJIJUN - 1,
                    meisai.getHyojijun());

            // 機器品名
            setCellValue(row, Mcm1003pConstants.MEISAI_COLINDEX_KIKIHINMEI_NK - 1,
                    meisai.getKikihinmeiNk());

            // 機器型式
            setCellValue(row, Mcm1003pConstants.MEISAI_COLINDEX_KIKIKATASHIKI - 1,
                    meisai.getKikikatashiki());

            // 数量（数値）
            Cell suryoCell = getOrCreateCell(row, Mcm1003pConstants.MEISAI_COLINDEX_SURYO_NM - 1);
            if (meisai.getSuryoNm() != null) {
                suryoCell.setCellValue(meisai.getSuryoNm().doubleValue());
            }

            // 点検回数
            setCellValue(row, Mcm1003pConstants.MEISAI_COLINDEX_TENKENKAISU - 1,
                    meisai.getTenkenkaisu());

            // 点検月
            setCellValue(row, Mcm1003pConstants.MEISAI_COLINDEX_TENKEN_TSUKI - 1,
                    meisai.getTenkenTsuki());

            // 備考
            setCellValue(row, Mcm1003pConstants.MEISAI_COLINDEX_BIKO_MEISAI - 1,
                    meisai.getBiko());
        }
    }

    // ===================================================================
    // ヘルパーメソッド
    // ===================================================================

    /**
     * 指定した行・列（1-based）のセルに文字列値を設定する。
     *
     * @param sheet    対象シート
     * @param rowIndex 行インデックス（1-based）
     * @param colIndex 列インデックス（1-based）
     * @param value    設定する文字列値
     */
    private void setCellValueByIndex(Sheet sheet, int rowIndex, int colIndex, String value) {
        int rowIdx = rowIndex - 1; // 0-based変換
        int colIdx = colIndex - 1; // 0-based変換
        Row row = sheet.getRow(rowIdx);
        if (row == null) {
            row = sheet.createRow(rowIdx);
        }
        Cell cell = row.getCell(colIdx);
        if (cell == null) {
            cell = row.createCell(colIdx);
        }
        if (value != null) {
            cell.setCellValue(value);
        }
    }

    /**
     * セルに文字列値を設定する。
     */
    private void setCellValue(Row row, int colIdx, String value) {
        Cell cell = getOrCreateCell(row, colIdx);
        if (value != null) {
            cell.setCellValue(value);
        }
    }

    /**
     * セルを取得（なければ作成）。
     */
    private Cell getOrCreateCell(Row row, int colIdx) {
        Cell cell = row.getCell(colIdx);
        if (cell == null) {
            cell = row.createCell(colIdx);
        }
        return cell;
    }
}
