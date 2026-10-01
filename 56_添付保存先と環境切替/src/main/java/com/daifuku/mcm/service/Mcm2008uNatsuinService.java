package com.daifuku.mcm.service;

import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.constants.Mcm2008uConstants;
import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.form.Mcm2008uForm;
import com.daifuku.mcm.repository.Mcm2008uRepository;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFAnchor;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAnchorClientData;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.support.TransactionSynchronization;
import org.springframework.transaction.support.TransactionSynchronizationManager;

import javax.xml.namespace.QName;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/**
 * 【変換元】Mcm2008uExcel.vb / CPExcelManager.UpdateExcelReport()
 * MCM2008U 審査・承認（店舗見積）承認時の添付Excel捺印サービス。
 *
 * VB版は Excel COM で見本の捺印枠（stp_circle / stp_topline / stp_btmline）を
 * 審査印・承認印の位置へコピー＆貼り付けし、図形 {審査|承認}_lastname / _date / _firstname に
 * 文字を設定していた。Web版（Apache POI）では図形のコピー＆貼り付けが安定しないため、
 * 以下の方式で同じ見た目を再現する。
 * <ul>
 *   <li>丸枠・上下線：見本（stp_*）と同じ寸法・相対位置で新規に描画する（赤・1.5pt）</li>
 *   <li>文字：テンプレートに配置済みの {shinsa|shonin}_* 図形へ設定する（書式は見本 stp_* 準拠）</li>
 * </ul>
 *
 * 捺印内容（VB版準拠）:
 * <ul>
 *   <li>_lastname  … "DTS" 固定（McmConstants.NATSUON_UE）</li>
 *   <li>_date      … 処理日（yy.MM.dd）</li>
 *   <li>_firstname … ログインユーザーの姓（VB版 UserLastName と同じ抽出ロジック）</li>
 * </ul>
 *
 * 承認処理がロールバックされた場合は、捺印前のファイル内容に復元する。
 */
@Service
public class Mcm2008uNatsuinService {

    private static final Logger logger = LoggerFactory.getLogger(Mcm2008uNatsuinService.class);

    /** 列インデックスの上限（XLSX: XFD列 = 16384列） */
    private static final int MAX_COL = 16383;

    /** 行インデックスの上限（XLSX: 1,048,576行） */
    private static final int MAX_ROW = 1048575;

    /** 名前の空白が無い場合に姓とみなす先頭文字数（VB版 McmLoginForm.vb 準拠） */
    private static final int LAST_NAME_DEFAULT_LENGTH = 3;

    private static final String SUFFIX_LASTNAME = "_lastname";
    private static final String SUFFIX_DATE = "_date";
    private static final String SUFFIX_FIRSTNAME = "_firstname";
    private static final String SUFFIX_CIRCLE = "_circle";
    private static final String SUFFIX_TOPLINE = "_topline";
    private static final String SUFFIX_BTMLINE = "_btmline";

    /** 捺印枠の線幅（pt）。見本 stp_circle の a:ln w="19050"（EMU）= 1.5pt */
    private static final double FRAME_LINE_WIDTH_PT = 1.5;

    @Autowired
    private Mcm2008uRepository repository;

    /** 添付ファイルの保存ルート。MCM2003U と同一の保管場所を参照する。 */
    @Autowired private com.daifuku.mcm.common.AttachmentStorage storage;

    /**
     * 店舗見積に紐付く添付Excelへ、承認状態に応じた捺印を行う。
     * 【変換元】Mcm2008uScreen.vb SYONINButton_Click（捺印押下）/ Mcm2008uExcel.WriteDataInExcel()
     * <ul>
     *   <li>承認状態が審査中(1) → 審査印（shinsa_*）</li>
     *   <li>承認状態が承認中(2) → 承認印（shonin_*）</li>
     * </ul>
     * 捺印対象は拡張子 .xlsx の添付のみ（VB版 CPConstant.FILETYPE_EXCEL 準拠）。
     * 呼び出し元のトランザクションがロールバックされた場合、ファイルを捺印前の内容に戻す。
     *
     * @param umKihonMitsumoriId 店舗基本見積ID
     * @param syouninJotai       更新前の承認状態コード
     * @param userName           ログインユーザーの氏名（MCM_MO_TANTO.TANTO_NK）
     * @throws McmBusinessException 添付ファイルが存在しない、または捺印に失敗した場合
     */
    public void stamp(BigDecimal umKihonMitsumoriId, String syouninJotai, String userName) {
        String prefix;
        if (McmConstants.SHONINJOTAI_SHINSACHU_CD.equals(syouninJotai)) {
            prefix = Mcm2008uConstants.SHAPE_NAME_SHINSA;
        } else if (McmConstants.SHONINJOTAI_SHONINCHU_CD.equals(syouninJotai)) {
            prefix = Mcm2008uConstants.SHAPE_NAME_SHONIN;
        } else {
            return;
        }

        String lastName = extractLastName(userName);
        String date = LocalDate.now().format(DateTimeFormatter.ofPattern(Mcm2008uConstants.NATSUIN_DATE_FORMAT));

        List<Mcm2008uForm.TenpuRowForm> tenpuList = repository.findTenpuByKihonMitsumoriId(umKihonMitsumoriId);
        for (Mcm2008uForm.TenpuRowForm tenpu : tenpuList) {
            String fileName = tenpu.getTenpufileNk();
            // VB版 UpdateExcelReport: 拡張子が .xlsx 以外は処理しない
            if (fileName == null
                    || !fileName.toLowerCase(Locale.ROOT).endsWith(Mcm2008uConstants.FILETYPE_EXCEL)) {
                continue;
            }
            Path file = resolveFile(tenpu);
            stampFile(file, fileName, prefix, lastName, date);
        }
    }

    /**
     * 氏名から姓を抽出する。
     * 【変換元】McmLoginForm.vb btnLogin_Click() の UserLastName 算出ロジック
     * <ul>
     *   <li>半角空白、無ければ全角空白の位置で区切り、その手前を姓とする</li>
     *   <li>空白が無い場合は先頭3文字を姓とする（3文字未満は全体）</li>
     *   <li>空または null の場合は空文字</li>
     * </ul>
     *
     * @param userName 氏名
     * @return 姓
     */
    static String extractLastName(String userName) {
        if (userName == null) {
            return "";
        }
        String name = userName.trim();
        if (name.isEmpty()) {
            return "";
        }
        int index = name.indexOf(' ');
        if (index < 0) {
            index = name.indexOf('\u3000');
        }
        if (index < 0) {
            return name.substring(0, Math.min(LAST_NAME_DEFAULT_LENGTH, name.length()));
        }
        return name.substring(0, index);
    }

    // ===================================================================
    // ファイル単位の処理
    // ===================================================================

    /**
     * 添付ファイルの物理パスを解決する（MCM2003U 添付資料と同一の保管場所・検証）。
     */
    private Path resolveFile(Mcm2008uForm.TenpuRowForm tenpu) {
        return storage.resolve(tenpu.getDirectory(), tenpu.getTenpufileNk());
    }

    /**
     * 1ファイルに捺印する。捺印対象シート・図形が無いファイルは変更しない。
     */
    private void stampFile(Path file, String fileName, String prefix, String lastName, String date) {
        byte[] original;
        byte[] stamped;
        try {
            original = Files.readAllBytes(file);
            try (XSSFWorkbook book = new XSSFWorkbook(new ByteArrayInputStream(original));
                 ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                boolean changed = false;
                for (int i = 0; i < book.getNumberOfSheets(); i++) {
                    XSSFSheet sheet = book.getSheetAt(i);
                    changed |= stampSheet(sheet, fileName, prefix, lastName, date);
                }
                if (!changed) {
                    logger.info("MCM2008U 捺印対象の図形が無いため捺印しません: {}", fileName);
                    return;
                }
                book.write(out);
                stamped = out.toByteArray();
            }
        } catch (McmBusinessException e) {
            throw e;
        } catch (Exception e) {
            logger.error("MCM2008U 捺印に失敗しました: {}", fileName, e);
            throw new McmBusinessException("添付資料への捺印に失敗しました。（" + fileName + "）", e);
        }

        // 書込み前に復元処理を登録し、承認処理のロールバック時は捺印前に戻す。
        registerRestore(file, fileName, original);
        try {
            Files.write(file, stamped);
        } catch (IOException e) {
            logger.error("MCM2008U 捺印ファイルの保存に失敗しました: {}", fileName, e);
            throw new McmBusinessException("添付資料への捺印に失敗しました。（" + fileName + "）", e);
        }
    }

    /**
     * 承認処理（トランザクション）がコミットされなかった場合に、ファイルを捺印前の内容へ戻す。
     */
    private void registerRestore(Path file, String fileName, byte[] original) {
        if (!TransactionSynchronizationManager.isSynchronizationActive()) {
            return;
        }
        TransactionSynchronizationManager.registerSynchronization(new TransactionSynchronization() {
            @Override
            public void afterCompletion(int status) {
                if (status == STATUS_COMMITTED) {
                    return;
                }
                try {
                    Files.write(file, original);
                } catch (IOException e) {
                    logger.error("MCM2008U 捺印ファイルの復元に失敗しました: {}", fileName, e);
                }
            }
        });
    }

    // ===================================================================
    // シート単位の処理
    // ===================================================================

    /**
     * 1シートに捺印する。
     * 【変換元】Mcm2008uExcel.WriteDataInExcel() — シート名が「見積総括表」「保守見積総括表(原価）」で
     * 始まるシートのみ対象。
     *
     * @return 捺印した場合 true
     */
    private boolean stampSheet(XSSFSheet sheet, String fileName, String prefix, String lastName, String date) {
        XSSFDrawing drawing = sheet.getDrawingPatriarch();
        if (drawing == null) {
            return false;
        }

        Map<String, XSSFSimpleShape> shapes = new HashMap<>();
        for (XSSFShape shape : drawing.getShapes()) {
            if (shape instanceof XSSFSimpleShape simple) {
                String name = shapeName(simple);
                if (name != null) {
                    // Excel の図形名は大文字小文字を区別しない（VB版は "SHINSA_lastname" で参照）
                    shapes.put(name.toLowerCase(Locale.ROOT), simple);
                }
            }
        }

        XSSFSimpleShape targetLast = shapes.get(prefix + SUFFIX_LASTNAME);
        XSSFSimpleShape targetDate = shapes.get(prefix + SUFFIX_DATE);
        XSSFSimpleShape targetFirst = shapes.get(prefix + SUFFIX_FIRSTNAME);
        if (targetLast == null || targetDate == null || targetFirst == null) {
            return false;
        }

        if (!isTargetSheet(sheet.getSheetName())) {
            logger.warn("MCM2008U 捺印図形はありますが対象外のシート名のため捺印しません: file={}, sheet={}",
                fileName, sheet.getSheetName());
            return false;
        }

        String sample = Mcm2008uConstants.SHAPE_NAME_SAMPLE;
        XSSFSimpleShape sampleLast = shapes.get(sample + SUFFIX_LASTNAME);
        XSSFSimpleShape sampleDate = shapes.get(sample + SUFFIX_DATE);
        XSSFSimpleShape sampleFirst = shapes.get(sample + SUFFIX_FIRSTNAME);
        XSSFSimpleShape sampleCircle = shapes.get(sample + SUFFIX_CIRCLE);
        XSSFSimpleShape sampleTop = shapes.get(sample + SUFFIX_TOPLINE);
        XSSFSimpleShape sampleBtm = shapes.get(sample + SUFFIX_BTMLINE);
        if (sampleLast == null || sampleCircle == null || sampleTop == null || sampleBtm == null) {
            // VB版も見本枠（stp_circle 等）が無い場合は捺印できない
            throw new McmBusinessException(
                "捺印枠（stp_circle／stp_topline／stp_btmline）がテンプレートに見つかりません。（" + fileName + "）");
        }

        // 再捺印時に枠が重複しないよう、同名の枠があれば削除してから描画する
        removeShapes(drawing, prefix + SUFFIX_CIRCLE, prefix + SUFFIX_TOPLINE, prefix + SUFFIX_BTMLINE);

        long[] nextId = {maxShapeId(drawing)};
        XSSFClientAnchor targetRef = anchorOf(targetLast, fileName);
        XSSFClientAnchor sampleRef = anchorOf(sampleLast, fileName);
        drawFrame(sheet, drawing, sampleCircle, ShapeTypes.ELLIPSE, prefix + SUFFIX_CIRCLE,
            targetRef, sampleRef, nextId, fileName);
        drawFrame(sheet, drawing, sampleTop, ShapeTypes.LINE, prefix + SUFFIX_TOPLINE,
            targetRef, sampleRef, nextId, fileName);
        drawFrame(sheet, drawing, sampleBtm, ShapeTypes.LINE, prefix + SUFFIX_BTMLINE,
            targetRef, sampleRef, nextId, fileName);

        // 【変換元】setNatsuin(): 苗字欄="DTS"、日付欄=処理日、名前欄=ログインユーザーの姓
        setStampText(targetLast, sampleLast, McmConstants.NATSUON_UE);
        setStampText(targetDate, sampleDate, date);
        setStampText(targetFirst, sampleFirst, lastName);
        return true;
    }

    /**
     * 捺印対象のシート名か判定する（前方一致）。全角・半角の括弧の違いは同一とみなす。
     */
    private boolean isTargetSheet(String sheetName) {
        String name = normalizeParen(sheetName);
        return name.startsWith(normalizeParen(Mcm2008uConstants.SHEET_NAME_TENPO_SOKATSU))
            || name.startsWith(normalizeParen(Mcm2008uConstants.SHEET_NAME_TENPO_GENKA));
    }

    private String normalizeParen(String value) {
        return value == null ? "" : value.replace('（', '(').replace('）', ')');
    }

    // ===================================================================
    // 図形操作
    // ===================================================================

    /**
     * 見本の枠図形（stp_*）と同じ寸法・相対位置で、捺印位置に枠図形を新規描画する。
     * 位置は「見本の文字図形(stp_lastname) → 捺印先の文字図形({shinsa|shonin}_lastname)」の移動量を
     * 見本の枠図形に加えて算出する。
     */
    private void drawFrame(XSSFSheet sheet, XSSFDrawing drawing, XSSFSimpleShape sample, int shapeType,
                           String name, XSSFClientAnchor targetRef, XSSFClientAnchor sampleRef,
                           long[] nextId, String fileName) {
        XSSFClientAnchor src = anchorOf(sample, fileName);

        long moveX = colPos(sheet, targetRef.getCol1(), targetRef.getDx1())
            - colPos(sheet, sampleRef.getCol1(), sampleRef.getDx1());
        long moveY = rowPos(sheet, targetRef.getRow1(), targetRef.getDy1())
            - rowPos(sheet, sampleRef.getRow1(), sampleRef.getDy1());

        int[] from = colMarker(sheet, colPos(sheet, src.getCol1(), src.getDx1()) + moveX);
        int[] to = colMarker(sheet, colPos(sheet, src.getCol2(), src.getDx2()) + moveX);
        int[] rowFrom = rowMarker(sheet, rowPos(sheet, src.getRow1(), src.getDy1()) + moveY);
        int[] rowTo = rowMarker(sheet, rowPos(sheet, src.getRow2(), src.getDy2()) + moveY);

        XSSFClientAnchor anchor = new XSSFClientAnchor(
            from[1], rowFrom[1], to[1], rowTo[1], from[0], rowFrom[0], to[0], rowTo[0]);

        XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
        // POI既定のテーマ参照スタイル（塗り・影など）が付くと見本と見た目が変わるため外す
        if (shape.getCTShape().isSetStyle()) {
            shape.getCTShape().unsetStyle();
        }
        shape.setShapeType(shapeType);
        shape.setNoFill(true);
        shape.setLineStyleColor(0xFF, 0x00, 0x00);
        shape.setLineWidth(FRAME_LINE_WIDTH_PT);

        // 図形IDの重複は Excel の「修復」対象になるため、既存の最大ID+1 を採番する
        nextId[0]++;
        CTShape ct = shape.getCTShape();
        ct.getNvSpPr().getCNvPr().setId(nextId[0]);
        ct.getNvSpPr().getCNvPr().setName(name);

        copyAnchorSettings(drawing, shapeName(sample), nextId[0]);
    }

    /**
     * 見本の枠図形のアンカー設定（印刷有無・移動/サイズ変更の扱い）を新規図形に引き継ぐ。
     * 見本は印刷対象外（fPrintsWithSheet="0"）のため、枠だけ印刷されることを防ぐ。
     */
    private void copyAnchorSettings(XSSFDrawing drawing, String sampleName, long newId) {
        CTTwoCellAnchor srcAnchor = null;
        CTTwoCellAnchor dstAnchor = null;
        for (CTTwoCellAnchor anchor : drawing.getCTDrawing().getTwoCellAnchorList()) {
            if (!anchor.isSetSp()) {
                continue;
            }
            String name = anchor.getSp().getNvSpPr().getCNvPr().getName();
            long id = anchor.getSp().getNvSpPr().getCNvPr().getId();
            if (srcAnchor == null && sampleName != null && sampleName.equalsIgnoreCase(name)) {
                srcAnchor = anchor;
            }
            if (id == newId) {
                dstAnchor = anchor;
            }
        }
        if (srcAnchor == null || dstAnchor == null) {
            return;
        }

        if (srcAnchor.isSetEditAs()) {
            dstAnchor.setEditAs(srcAnchor.getEditAs());
        } else if (dstAnchor.isSetEditAs()) {
            dstAnchor.unsetEditAs();
        }

        CTAnchorClientData src = srcAnchor.getClientData();
        CTAnchorClientData dst = dstAnchor.getClientData();
        if (src == null || dst == null) {
            return;
        }
        if (src.isSetFPrintsWithSheet()) {
            dst.setFPrintsWithSheet(src.getFPrintsWithSheet());
        } else if (dst.isSetFPrintsWithSheet()) {
            dst.unsetFPrintsWithSheet();
        }
        if (src.isSetFLocksWithSheet()) {
            dst.setFLocksWithSheet(src.getFLocksWithSheet());
        } else if (dst.isSetFLocksWithSheet()) {
            dst.unsetFLocksWithSheet();
        }
    }

    /**
     * 文字図形のテキストを差し替える。書式は見本図形（stp_*）の文字書式を引き継ぎ、
     * 見本が無い場合は捺印先図形の既定書式（endParaRPr）を用いる。
     */
    private void setStampText(XSSFSimpleShape target, XSSFSimpleShape sample, String text) {
        CTShape ct = target.getCTShape();
        if (!ct.isSetTxBody()) {
            // 文字を持たない図形は想定外（テンプレートでは全て txBody を持つ）
            target.setText(text);
            return;
        }
        CTTextBody body = ct.getTxBody();
        CTTextParagraph paragraph = body.sizeOfPArray() > 0 ? body.getPArray(0) : body.addNewP();
        while (body.sizeOfPArray() > 1) {
            body.removeP(1);
        }
        while (paragraph.sizeOfRArray() > 0) {
            paragraph.removeR(0);
        }
        while (paragraph.sizeOfFldArray() > 0) {
            paragraph.removeFld(0);
        }
        while (paragraph.sizeOfBrArray() > 0) {
            paragraph.removeBr(0);
        }

        CTTextCharacterProperties props = sampleRunProperties(sample);
        if (props == null && paragraph.isSetEndParaRPr()) {
            props = paragraph.getEndParaRPr();
        }

        CTRegularTextRun run = paragraph.addNewR();
        if (props != null) {
            run.setRPr((CTTextCharacterProperties) props.copy());
        }
        run.setT(text == null ? "" : text);
    }

    /**
     * 見本図形の先頭文字列の書式を取得する。
     */
    private CTTextCharacterProperties sampleRunProperties(XSSFSimpleShape sample) {
        if (sample == null || !sample.getCTShape().isSetTxBody()) {
            return null;
        }
        CTTextBody body = sample.getCTShape().getTxBody();
        for (CTTextParagraph paragraph : body.getPList()) {
            for (CTRegularTextRun run : paragraph.getRList()) {
                if (run.isSetRPr()) {
                    return run.getRPr();
                }
            }
        }
        return null;
    }

    /**
     * 指定した名前の図形（2セルアンカー）を削除する。
     */
    private void removeShapes(XSSFDrawing drawing, String... names) {
        CTDrawing ctDrawing = drawing.getCTDrawing();
        for (int i = ctDrawing.sizeOfTwoCellAnchorArray() - 1; i >= 0; i--) {
            CTTwoCellAnchor anchor = ctDrawing.getTwoCellAnchorArray(i);
            if (!anchor.isSetSp()) {
                continue;
            }
            String name = anchor.getSp().getNvSpPr().getCNvPr().getName();
            for (String target : names) {
                if (target.equalsIgnoreCase(name)) {
                    ctDrawing.removeTwoCellAnchor(i);
                    break;
                }
            }
        }
    }

    /**
     * 描画内の図形IDの最大値を取得する（グループ内の図形も含む）。
     */
    private long maxShapeId(XSSFDrawing drawing) {
        long max = 0;
        XmlCursor cursor = drawing.getCTDrawing().newCursor();
        try {
            QName idAttr = new QName("id");
            while (!cursor.toNextToken().isNone()) {
                if (cursor.isStart() && "cNvPr".equals(cursor.getName().getLocalPart())) {
                    String value = cursor.getAttributeText(idAttr);
                    if (value != null) {
                        try {
                            max = Math.max(max, Long.parseLong(value.trim()));
                        } catch (NumberFormatException ignored) {
                            // 数値以外のIDは採番対象外
                        }
                    }
                }
            }
        } finally {
            cursor.dispose();
        }
        return max;
    }

    private String shapeName(XSSFSimpleShape shape) {
        CTShape ct = shape.getCTShape();
        if (ct.getNvSpPr() == null || ct.getNvSpPr().getCNvPr() == null) {
            return null;
        }
        return ct.getNvSpPr().getCNvPr().getName();
    }

    private XSSFClientAnchor anchorOf(XSSFShape shape, String fileName) {
        XSSFAnchor anchor = shape.getAnchor();
        if (anchor instanceof XSSFClientAnchor client) {
            return client;
        }
        throw new McmBusinessException("捺印枠の配置情報を取得できません。（" + fileName + "）");
    }

    // ===================================================================
    // 位置計算（EMU）
    // ===================================================================

    /** 列の先頭からの絶対位置（EMU） */
    private long colPos(XSSFSheet sheet, int col, long offset) {
        long x = offset;
        for (int c = 0; c < col; c++) {
            x += colWidth(sheet, c);
        }
        return x;
    }

    /** 行の先頭からの絶対位置（EMU） */
    private long rowPos(XSSFSheet sheet, int row, long offset) {
        long y = offset;
        for (int r = 0; r < row; r++) {
            y += rowHeight(sheet, r);
        }
        return y;
    }

    /** 絶対位置（EMU）を {列, 列内オフセット} に変換する */
    private int[] colMarker(XSSFSheet sheet, long x) {
        long rest = Math.max(0, x);
        int col = 0;
        while (col < MAX_COL) {
            long width = colWidth(sheet, col);
            if (rest < width) {
                break;
            }
            rest -= width;
            col++;
        }
        return new int[]{col, (int) rest};
    }

    /** 絶対位置（EMU）を {行, 行内オフセット} に変換する */
    private int[] rowMarker(XSSFSheet sheet, long y) {
        long rest = Math.max(0, y);
        int row = 0;
        while (row < MAX_ROW) {
            long height = rowHeight(sheet, row);
            if (rest < height) {
                break;
            }
            rest -= height;
            row++;
        }
        return new int[]{row, (int) rest};
    }

    private long colWidth(XSSFSheet sheet, int col) {
        return Math.round(sheet.getColumnWidthInPixels(col) * Units.EMU_PER_PIXEL);
    }

    private long rowHeight(XSSFSheet sheet, int row) {
        XSSFRow r = sheet.getRow(row);
        float points = (r != null) ? r.getHeightInPoints() : sheet.getDefaultRowHeightInPoints();
        return Math.round(points * Units.EMU_PER_POINT);
    }
}
