package com.daifuku.mcm.service;

import com.daifuku.mcm.common.BaseExcelService;
import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.form.Mcm1011uForm.KeiyakuRowForm;
import java.io.*;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ddf.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.support.*;

/** 正式VB Mcm1011uExcel: 契約手続依頼書①の審査印・承認印。 */
@Service
public class Mcm1011uNatsuinService extends BaseExcelService {
    @Autowired private Mcm1005uAttachmentService attachments;

    /** 全帳票を検証・生成してから置換する。通常のDBロールバックでは原本へ戻す。 */
    public void stamp(List<KeiyakuRowForm> rows, String userName) {
        if (!TransactionSynchronizationManager.isActualTransactionActive())
            throw new IllegalStateException("Stamping requires a transaction");
        var prepared = new LinkedHashMap<Path, byte[]>();
        var originals = new LinkedHashMap<Path, byte[]>();
        for (var row : rows) for (var file : row.getPendingAttachments()) {
            if (!Set.of("1", "2").contains(String.valueOf(file.getShoninjotai()))) continue;
            String name = file.getTenpufileNk();
            // 添付PDF/Word等には捺印しない。拡張子だけ変更して承認できないようExcelは実内容も検査。
            if (name == null || !name.toLowerCase(Locale.ROOT).matches(".*\\.(xls|xlsx|xlsm)$")) continue;
            try {
                var entity = new com.daifuku.mcm.entity.McmTkTenpuEntity();
                entity.setDirectory(file.getDirectory()); entity.setTenpufileNk(name);
                Path path = attachments.file(entity).toRealPath();
                if (originals.containsKey(path)) throw new McmBusinessException("複数の申請が同じ添付ファイルを参照しています。保存先を確認してください。");
                byte[] original = Files.readAllBytes(path);
                byte[] result = stampBytes(original, file.getShoninjotai(), userName, LocalDate.now());
                originals.put(path, original);
                if (!Arrays.equals(original, result)) prepared.put(path, result);
            } catch (McmBusinessException ex) { throw ex; }
            catch (Exception ex) {
                logger.error("MCM1011U 捺印準備失敗 attachment={}", file.getTkTenpuId(), ex);
                throw new McmBusinessException("添付帳票の捺印に失敗しました。ファイルと保存先を確認してください。");
            }
        }
        for (var entry : prepared.entrySet()) replace(entry.getKey(), originals.get(entry.getKey()), entry.getValue());
    }

    /** 副作用のない帳票変換。原本と同形式で出力。VB同様、別種のExcel資料は変更しない。 */
    public byte[] stampBytes(byte[] source, String state, String name, LocalDate date) throws IOException {
        String prefix = switch (state) { case "1" -> "shinsa"; case "2" -> "shonin";
            default -> throw new McmBusinessException("捺印できない承認状態です。"); };
        try (Workbook book = WorkbookFactory.create(new ByteArrayInputStream(source))) {
            int count = 0;
            for (Sheet sheet : book) if (sheet.getSheetName().startsWith("契約手続依頼書①")) {
                if (sheet instanceof HSSFSheet h) stampHssf(h, prefix, name, date);
                else if (sheet instanceof XSSFSheet x) stampXssf(x, prefix, name, date);
                else throw new McmBusinessException("対応していない帳票形式です。");
                count++;
            }
            if (count == 0) return source;
            return toByteArray(book);
        }
    }

    private String text(String suffix, String name, LocalDate date) {
        return switch (suffix) { case "lastname" -> "DTS";
            case "date" -> date.format(DateTimeFormatter.ofPattern("yy.MM.dd"));
            default -> Mcm2008uNatsuinService.extractLastName(name); };
    }
    private String key(String name) { return name == null ? "" : name.replace("\u0000", "").toLowerCase(Locale.ROOT); }
    private McmBusinessException missing() { return new McmBusinessException("添付帳票に捺印用の図形がありません。帳票を確認してください。"); }

    private void stampHssf(HSSFSheet sheet, String prefix, String name, LocalDate date) {
        HSSFPatriarch drawing = sheet.getDrawingPatriarch();
        if (drawing == null) throw missing();
        var shapes = new HashMap<String, HSSFSimpleShape>();
        for (var s : drawing.getChildren()) if (s instanceof HSSFSimpleShape simple) shapes.put(key(s.getShapeName()), simple);
        for (String part : List.of("lastname", "firstname", "date", "circle", "topline", "btmline"))
            if (!shapes.containsKey("stp_" + part)) throw missing();
        for (String part : List.of("lastname", "firstname", "date")) if (!shapes.containsKey(prefix + "_" + part)) throw missing();
        var origin = anchor(shapes.get("stp_lastname")); var target = anchor(shapes.get(prefix + "_lastname"));
        for (String part : List.of("circle", "topline", "btmline")) {
            String frameName = prefix + "_" + part;
            if (shapes.containsKey(frameName)) drawing.removeShape(shapes.get(frameName));
            var sample = shapes.get("stp_" + part);
            var frame = drawing.createSimpleShape((HSSFClientAnchor) shifted(sheet, anchor(sample), origin, target));
            frame.setShapeType(sample.getShapeType()); frame.setLineStyle(sample.getLineStyle());
            frame.setLineStyleColor(sample.getLineStyleColor()); frame.setLineWidth(sample.getLineWidth()); frame.setNoFill(true);
            byte[] bytes = (frameName + "\u0000").getBytes(java.nio.charset.StandardCharsets.UTF_16LE);
            var property = new EscherComplexProperty(EscherPropertyTypes.GROUPSHAPE__SHAPENAME, false, bytes.length);
            property.setComplexData(bytes); frame.getOptRecord().setEscherProperty(property);
            frame.getOptRecord().sortProperties();
        }
        for (String part : List.of("lastname", "date", "firstname")) {
            var sample = shapes.get("stp_" + part).getString();
            var value = new HSSFRichTextString(text(part, name, date));
            if (sample.numFormattingRuns() > 0) value.applyFont(sample.getFontOfFormattingRun(0));
            shapes.get(prefix + "_" + part).setString(value);
        }
    }
    private HSSFClientAnchor anchor(HSSFSimpleShape shape) {
        if (!(shape.getAnchor() instanceof HSSFClientAnchor a)) throw missing();
        return a;
    }

    private void stampXssf(XSSFSheet sheet, String prefix, String name, LocalDate date) {
        XSSFDrawing drawing = sheet.getDrawingPatriarch();
        if (drawing == null) throw missing();
        var shapes = new HashMap<String, XSSFSimpleShape>();
        for (var s : drawing.getShapes()) if (s instanceof XSSFSimpleShape simple)
            shapes.put(key(simple.getCTShape().getNvSpPr().getCNvPr().getName()), simple);
        for (String part : List.of("lastname", "firstname", "date", "circle", "topline", "btmline"))
            if (!shapes.containsKey("stp_" + part)) throw missing();
        for (String part : List.of("lastname", "firstname", "date")) if (!shapes.containsKey(prefix + "_" + part)) throw missing();
        var origin = anchor(shapes.get("stp_lastname")); var target = anchor(shapes.get(prefix + "_lastname"));
        for (String part : List.of("circle", "topline", "btmline")) {
            String frameName = prefix + "_" + part;
            // POIのremoveShapeはないため、同名図形はそのまま再利用して重複を防ぐ。
            var sample = shapes.get("stp_" + part);
            var frame = shapes.get(frameName);
            if (frame == null) {
                frame = drawing.createSimpleShape((XSSFClientAnchor) shifted(sheet, anchor(sample), origin, target));
                var position = frame.getCTShape().getSpPr().getXfrm().copy();
                frame.getCTShape().getSpPr().set(sample.getCTShape().getSpPr());
                frame.getCTShape().getSpPr().getXfrm().set(position);
                frame.getCTShape().getNvSpPr().getCNvPr().setName(frameName);
            }
        }
        for (String part : List.of("lastname", "date", "firstname")) {
            var dest = shapes.get(prefix + "_" + part);
            var sample = shapes.get("stp_" + part);
            var body = dest.getCTShape().getTxBody();
            body.set(sample.getCTShape().getTxBody());
            var p = body.getPArray(0);
            var style = p.sizeOfRArray() > 0 && p.getRArray(0).isSetRPr() ? p.getRArray(0).getRPr().copy() : null;
            while (body.sizeOfPArray() > 1) body.removeP(1);
            while (p.sizeOfRArray() > 0) p.removeR(0);
            var run = p.addNewR(); if (style != null) run.addNewRPr().set(style);
            run.setT(text(part, name, date));
        }
    }
    private XSSFClientAnchor anchor(XSSFSimpleShape shape) {
        if (!(shape.getAnchor() instanceof XSSFClientAnchor a)) throw missing();
        return a;
    }

    // HSSFのセル内相対座標 / XSSFのEMUをピクセルへ統一し、見本と文字欄の移動量を適用。
    private double x(Sheet s, int col, int dx, boolean xssf) {
        double result = 0; for (int i=0;i<col;i++) result += s.getColumnWidthInPixels(i);
        return result + (xssf ? dx / 9525.0 : s.getColumnWidthInPixels(col) * dx / 1024.0);
    }
    private double height(Sheet s, int row) { return (s.getRow(row)==null ? s.getDefaultRowHeightInPoints() : s.getRow(row).getHeightInPoints()) * 96.0/72.0; }
    private double y(Sheet s, int row, int dy, boolean xssf) {
        double result=0;for(int i=0;i<row;i++)result+=height(s,i);
        return result+(xssf ? dy/9525.0 : height(s,row)*dy/256.0);
    }
    private int[] marker(Sheet s, double pos, boolean col, boolean xssf) {
        int i=0, max=col ? 255 : 65535;
        if(xssf)max=col ? 16383 : 1048575;
        double size=col?s.getColumnWidthInPixels(i):height(s,i);
        while(i<max && pos>=size){pos-=size;i++;size=col?s.getColumnWidthInPixels(i):height(s,i);}
        if(pos<0 || size<=0)throw missing();
        return new int[]{i,(int)Math.round(xssf?pos*9525.0:pos/size*(col?1024:256))};
    }
    private ClientAnchor shifted(Sheet s, ClientAnchor a, ClientAnchor origin, ClientAnchor target) {
        boolean isX=s instanceof XSSFSheet;
        double dx=x(s,target.getCol1(),target.getDx1(),isX)-x(s,origin.getCol1(),origin.getDx1(),isX);
        double dy=y(s,target.getRow1(),target.getDy1(),isX)-y(s,origin.getRow1(),origin.getDy1(),isX);
        var c1=marker(s,x(s,a.getCol1(),a.getDx1(),isX)+dx,true,isX);
        var c2=marker(s,x(s,a.getCol2(),a.getDx2(),isX)+dx,true,isX);
        var r1=marker(s,y(s,a.getRow1(),a.getDy1(),isX)+dy,false,isX);
        var r2=marker(s,y(s,a.getRow2(),a.getDy2(),isX)+dy,false,isX);
        ClientAnchor result=s.getWorkbook().getCreationHelper().createClientAnchor();
        result.setCol1(c1[0]);result.setDx1(c1[1]);result.setCol2(c2[0]);result.setDx2(c2[1]);
        result.setRow1(r1[0]);result.setDy1(r1[1]);result.setRow2(r2[0]);result.setDy2(r2[1]);
        return result;
    }

    private void replace(Path file, byte[] original, byte[] stamped) {
        Path backup = file.resolveSibling(file.getFileName() + ".mcm1011.bak");
        try {
            // CREATE_NEWで別処理・異常終了後の原本を上書きしない。
            Files.write(backup, original, StandardOpenOption.CREATE_NEW);
            if (!Arrays.equals(Files.readAllBytes(file), original)) {
                Files.delete(backup);
                throw new IOException("File changed during approval");
            }
            TransactionSynchronizationManager.registerSynchronization(new TransactionSynchronization() {
                @Override public void afterCompletion(int status) {
                    try {
                        if (status == STATUS_COMMITTED) Files.delete(backup);
                        else if (status == STATUS_ROLLED_BACK) {
                            byte[] current = Files.readAllBytes(file);
                            if (!Arrays.equals(current, stamped) && !Arrays.equals(current, original))
                                throw new IOException("File was changed externally; retain recovery backup");
                            Files.move(backup, file, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
                        }
                        else logger.error("MCM1011U DB終了状態不明。捺印原本を保留: {}", backup);
                    } catch (IOException ex) { logger.error("MCM1011U 捺印原本の後処理失敗。復旧が必要: {}", backup, ex); }
                }
            });
            Path temp = Files.createTempFile(file.getParent(), ".mcm1011-", ".tmp");
            try { Files.write(temp, stamped); Files.move(temp, file, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE); }
            finally { Files.deleteIfExists(temp); }
        } catch (IOException ex) {
            logger.error("MCM1011U 捺印置換失敗: {}", file, ex);
            throw new McmBusinessException("添付帳票を保存できません。使用中のファイル・保存先・前回処理の復旧状態を確認してください。");
        }
    }
}
