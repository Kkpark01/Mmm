package com.daifuku.mcm.service;

import java.io.*;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.daifuku.mcm.service.McmCustomerReportService.Group;
import com.daifuku.mcm.service.McmCustomerReportService.Line;
import static com.daifuku.mcm.service.McmCustomerTotalsService.*;

/**
 * Mcm2001pConstant.vb の出力セルを基準にした再構成書式。
 * 元のExcelテンプレート未受領のため、罫線・列幅・固定文言は完全復元ではない。
 * 金額を再入力せず、総括表から明細の計算セルを参照する。
 */
final class McmCustomerReportLayout {
    private XSSFWorkbook book;
    private CellStyle normal,number,date,heading,title,percent;
    private final Map<Sheet,Integer> widths=new LinkedHashMap<>();
    private final Map<Group,String> quoteTotals=new HashMap<>(),costTotals=new HashMap<>();
    private final Map<Integer,CellStyle> precision=new HashMap<>();
    private Map<String,Object> h;

    byte[] render(Map<String,Object> header,List<Map<String,Object>> periods,
                  List<Map<String,Object>> brands,List<Group> groups)throws IOException {
        h=header;
        try(var wb=new XSSFWorkbook();var bytes=new ByteArrayOutputStream()){
            book=wb;styles();
            Sheet summary=sheet("見積総括表",Math.max(38,20+periods.size()*6),3);
            Sheet costSummary=sheet("保守見積総括表（原価）",Math.max(40,25+periods.size()*5),3);
            int i=0;
            for(Group g:groups){String suffix="_"+(++i);quote(g,suffix);inventory(g,suffix);cost(g,suffix);hard(g,suffix);soft(g,suffix);}
            summary(summary,periods,brands,groups,false);
            summary(costSummary,periods,brands,groups,true);
            book.getCreationHelper().createFormulaEvaluator().evaluateAll();
            for(Sheet s:book){
                for(Row row:s)for(Cell c:row)if(c.getCellType()==CellType.FORMULA&&c.getCachedFormulaResultType()==CellType.ERROR)
                    throw new IllegalStateException("帳票の計算に失敗しました："+s.getSheetName()+" "+c.getAddress());
                finish(s);
            }
            book.setForceFormulaRecalculation(true);book.write(bytes);return bytes.toByteArray();
        }
    }
    private void styles(){
        Font font=book.createFont();font.setFontName("ＭＳ Ｐゴシック");font.setFontHeightInPoints((short)10);
        normal=book.createCellStyle();normal.setFont(font);normal.setVerticalAlignment(VerticalAlignment.CENTER);normal.setWrapText(true);
        number=book.createCellStyle();number.cloneStyleFrom(normal);number.setAlignment(HorizontalAlignment.RIGHT);number.setDataFormat(book.createDataFormat().getFormat("#,##0.######;[Red]-#,##0.######"));
        date=book.createCellStyle();date.cloneStyleFrom(normal);date.setDataFormat(book.createDataFormat().getFormat("yyyy/mm/dd"));
        percent=book.createCellStyle();percent.cloneStyleFrom(number);percent.setDataFormat(book.createDataFormat().getFormat("0.0%"));
        Font bold=book.createFont();bold.setFontName(font.getFontName());bold.setFontHeightInPoints((short)10);bold.setBold(true);
        heading=book.createCellStyle();heading.cloneStyleFrom(normal);heading.setFont(bold);heading.setAlignment(HorizontalAlignment.CENTER);heading.setBorderBottom(BorderStyle.THIN);
        Font large=book.createFont();large.setFontName(font.getFontName());large.setFontHeightInPoints((short)15);large.setBold(true);
        title=book.createCellStyle();title.cloneStyleFrom(normal);title.setFont(large);title.setAlignment(HorizontalAlignment.CENTER);
    }
    private Sheet sheet(String name,int columns,int width){
        Sheet s=book.createSheet(name);widths.put(s,columns);s.setDisplayGridlines(false);s.setDefaultRowHeightInPoints(19);
        for(int i=0;i<columns;i++)s.setColumnWidth(i,width*256);return s;
    }
    private Cell put(Sheet s,int r,int c,int end,Object v){
        Row row=s.getRow(r-1);if(row==null)row=s.createRow(r-1);
        Cell cell=row.getCell(c-1,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        cell.setCellStyle(v instanceof Number?number:v instanceof LocalDate?date:normal);
        if(v instanceof Number n)cell.setCellValue(n.doubleValue());else if(v instanceof LocalDate d)cell.setCellValue(d);else if(str(v).isBlank()&&!str(v).contains("\n"))cell.setBlank();else cell.setCellValue(str(v));
        if(end>c)s.addMergedRegion(new CellRangeAddress(r-1,r-1,c-1,end-1));
        if(v!=null){int capacity=0;for(int i=c;i<=end;i++)capacity+=s.getColumnWidth(i-1)/256;
            int lines=0;for(String line:str(v).split("\\R",-1))lines+=Math.max(1,(line.codePoints().map(ch->ch<128?1:2).sum()+Math.max(1,capacity)-1)/Math.max(1,capacity));
            row.setHeightInPoints(Math.max(row.getHeightInPoints(),Math.min(409,lines*13+5)));}
        return cell;
    }
    private void label(Sheet s,int r,int c,int end,String v){put(s,r,c,end,v).setCellStyle(heading);}
    private void title(Sheet s,int r,int c,int end,String v){put(s,r,c,end,v).setCellStyle(title);s.getRow(r-1).setHeightInPoints(30);}
    private void formula(Sheet s,int r,int c,int end,String f){Cell cell=put(s,r,c,end,null);cell.setCellStyle(number);cell.setCellFormula(f);}
    /** Excelの行高上限を超える注記は複数行に分け、後続セクションを押し下げる。 */
    private int paragraph(Sheet s,int row,int c,int end,Object value){
        String text=str(value);int capacity=0;for(int i=c;i<=end;i++)capacity+=s.getColumnWidth(i-1)/256;
        int limit=Math.max(30,capacity*8);StringBuilder part=new StringBuilder();int weight=0;
        for(int cp:text.codePoints().toArray()){
            int w=cp=='\n'?capacity:cp<128?1:2;
            if(weight+w>limit&&!part.isEmpty()){put(s,row++,c,end,part.toString());part.setLength(0);weight=0;}
            part.appendCodePoint(cp);weight+=w;
        }
        put(s,row++,c,end,part.toString());return row;
    }
    private void ratio(Sheet s,int r,int c,String numerator,String denominator){formula(s,r,c,c,"IF("+denominator+"=0,\"\","+numerator+"/"+denominator+")");s.getRow(r-1).getCell(c-1).setCellStyle(percent);}
    private static String ref(Sheet s,int r,int c){return "'"+s.getSheetName().replace("'","''")+"'!"+CellReference.convertNumToColString(c-1)+r;}
    private static String str(Object o){return o==null?"":o instanceof BigDecimal b?b.stripTrailingZeros().toPlainString():o.toString();}
    private static String t(Map<String,Object> m,String k){return str(m.get(k));}
    private static LocalDate day(Object o){return o==null?null:LocalDate.parse(o.toString().substring(0,10));}
    private static String period(Group g){return str(day(g.period().get("KAISI_DT")))+" ～ "+str(day(g.period().get("SYURYO_DT")));}
    private static String brand(Group g){return (t(g.brand(),"BRAND_NK")+" "+t(g.brand(),"BRANDSYOSAI_NK")).strip();}
    private static boolean on(Map<String,Object> m,String k){return num(m,k).signum()!=0;}
    private static BigDecimal quantity(Line l){return new BigDecimal(l.row().getSuryoNm());}
    private static BigDecimal price(Line l){return nz(l.row().getHyojunKin()).multiply(quantity(l));}
    private static BigDecimal cost(Line l){return nz(l.row().getSikiriKin()).multiply(quantity(l));}
    private static List<Line> lines(Group g,Map<String,Object> config){return g.lines().stream().filter(l->l.row().getUmKikoseiId().compareTo(num(config,"UM_KIKIKOSEI_ID"))==0).toList();}
    private static List<Map<String,Object>> unselected(Group g,Map<String,Object> config){return g.unselected().stream().filter(m->num(m,"KIKIKOSEI_ID").compareTo(num(config,"KIKIKOSEI_ID"))==0).toList();}
    private static List<Map<String,Object>> displayConfigurations(Group g){
        var configs=new ArrayList<>(g.configurations());
        for(var m:g.unselected())if(configs.stream().noneMatch(k->num(k,"KIKIKOSEI_ID").compareTo(num(m,"KIKIKOSEI_ID"))==0))configs.add(m);
        return configs;
    }
    private void audit(Sheet s,int r,int labelCol,int valueCol,int end){
        label(s,r,labelCol,valueCol-1,"見積NO");put(s,r,valueCol,end,h.get("UM_MITSUMORI_NO"));
        label(s,r+1,labelCol,valueCol-1,"作成日");put(s,r+1,valueCol,end,day(h.get("MITSUMORI_DT")));
        label(s,r+2,labelCol,valueCol-1,"管理NO");put(s,r+2,valueCol,end,t(h,"UM_KIHON_MITSUMORI_ID"));
    }
    private void summary(Sheet s,List<Map<String,Object>> periods,List<Map<String,Object>> brands,List<Group> groups,boolean costs){
        int end=widths.get(s),row=costs?15:17,col=costs?22:21,step=costs?5:6;
        if(costs){title(s,1,2,end,"保守見積総括表（原価）");audit(s,2,25,29,end);
            label(s,6,2,5,"納入先");put(s,6,6,24,h.get("NONYUSAKI_NK"));
            label(s,7,2,5,"プラント");put(s,7,6,24,h.get("PLANT_NK"));
            label(s,8,2,5,"サポートID");put(s,8,6,13,h.get("SUPPORT_ID"));
            label(s,9,2,5,"見積レベル");put(s,9,6,13,h.get("MITSUMORILEVEL"));
            label(s,9,14,17,"保守時間");put(s,9,18,24,h.get("KEIYAKUJIKANTAI"));
        }else{audit(s,2,27,31,end);
            put(s,3,3,25,t(h,"SOFUMEISHO1_NK")+" "+t(h,"SOFUMEISHO2_NK"));put(s,4,3,25,t(h,"SOFUMEISHO3_NK")+" "+t(h,"SOFUMEISHO4_NK"));put(s,5,3,25,h.get("SOFUTANTO_NK"));
            title(s,8,3,end,t(h,"NONYUSAKI_NK")+" 見積総括表（"+t(h,"KEIYAKUJIKANTAI")+"H）");put(s,9,3,end,"サポートID："+t(h,"SUPPORT_ID")+" / "+t(h,"PLANT_NK"));
        }
        label(s,row-1,costs?3:5,col-1,costs?"ブランド / 費目":"ブランド / 保守費");
        for(int p=0;p<periods.size();p++){
            int c=col+p*step;put(s,row-3,c,c+step-1,day(periods.get(p).get("KAISI_DT")));put(s,row-2,c,c+step-1,day(periods.get(p).get("SYURYO_DT")));label(s,row-1,c,c+step-1,"金額（円）");
        }
        int start=row;
        for(var b:brands){
            if(costs){put(s,row,3,20,t(b,"BRANDSYOSAI_NK"));put(s,row,21,21,1);}
            else {put(s,row,5,19,t(b,"BRANDSYOSAI_NK"));put(s,row,20,20,"1式");}
            for(int p=0;p<periods.size();p++){
                Group g=find(groups,b,periods.get(p));int c=col+p*step;
                BigDecimal remote=on(b,"REMOTE_FLG")?num(b,"REMOTE_KIN"):BigDecimal.ZERO;
                BigDecimal excluded=remote;
                if(!costs)excluded=excluded.add(on(b,"DREMOS_FLG")?num(b,"DREMOS_KIN"):BigDecimal.ZERO).add(g.lines().stream().filter(l->"1".equals(l.row().getControllerFlg())).map(McmCustomerReportLayout::price).reduce(BigDecimal.ZERO,BigDecimal::add));
                formula(s,row,c,c+step-1,quoteTotals.get(g)+"-("+excluded.toPlainString()+")");
            }
            row++;
            if(!costs){
                // VB総括表はコントローラをブランド保守費の下に独立して表示する。
                Group first=find(groups,b,periods.get(0));
                var controllerNames=new LinkedHashSet<String>();for(Line l:first.lines())if("1".equals(l.row().getControllerFlg()))controllerNames.add(l.row().getMeisaiNm());
                for(String name:controllerNames){
                    put(s,row,10,19,name);BigDecimal qty=first.lines().stream().filter(l->"1".equals(l.row().getControllerFlg())&&Objects.equals(name,l.row().getMeisaiNm())).map(McmCustomerReportLayout::quantity).reduce(BigDecimal.ZERO,BigDecimal::add);put(s,row,20,20,qty);
                    for(int p=0;p<periods.size();p++){Group g=find(groups,b,periods.get(p));int c=col+p*step;put(s,row,c,c+step-1,g.lines().stream().filter(l->"1".equals(l.row().getControllerFlg())&&Objects.equals(name,l.row().getMeisaiNm())).map(McmCustomerReportLayout::price).reduce(BigDecimal.ZERO,BigDecimal::add));}row++;
                }
            }
        }
        if(!costs)row=optionRows(s,row,col,step,brands,periods,"DREMOS_FLG","DREMOS_KIN","DREMOS使用費");
        row=optionRows(s,row,col,step,brands,periods,"REMOTE_FLG","REMOTE_KIN","リモート一次導入費");
        int sum=Math.max(costs?42:51,row+1);label(s,sum,costs?3:5,col-1,"見積合計");
        // VBの「保守見積総括表(原価)」は期間別の保守見積額を集計する。仕切原価との混同を避ける。
        for(int p=0;p<periods.size();p++){
            int c=col+p*step;String column=CellReference.convertNumToColString(c-1);
            formula(s,sum,c,c+step-1,"SUM("+column+start+":"+column+(row-1)+")");
        }
        if(costs){
            // Web版で扱っていた期間単位の出精値引きも落とさず、VBの見積額合計と分離する。
            label(s,sum+2,3,col-1,"仕切原価合計");label(s,sum+3,3,col-1,"取引先出精値引き");label(s,sum+4,3,col-1,"仕切原価差引");
            for(int p=0;p<periods.size();p++){
                int c=col+p*step;final var selected=periods.get(p);String column=CellReference.convertNumToColString(c-1);
                formula(s,sum+2,c,c+step-1,"SUM("+String.join(",",groups.stream().filter(g->g.period()==selected).map(costTotals::get).toList())+")");
                put(s,sum+3,c,c+step-1,num(selected,"SYUSSEINEBIKI_KIN"));formula(s,sum+4,c,c+step-1,column+(sum+2)+"-"+column+(sum+3));
            }
        }
        if(!costs){label(s,sum+1,2,end,"見積条件");paragraph(s,sum+2,2,end,h.get("MITSUMORI_JOUKEN"));}
        s.setRepeatingRows(new CellRangeAddress(start-4,start-2,-1,-1));s.createFreezePane(col-1,start-1);
    }
    private static Group find(List<Group> groups,Map<String,Object> brand,Map<String,Object> period){return groups.stream().filter(g->g.brand()==brand&&g.period()==period).findFirst().orElseThrow();}
    private int optionRows(Sheet s,int row,int col,int step,List<Map<String,Object>> brands,List<Map<String,Object>> periods,String flag,String key,String name){
        if(brands.stream().noneMatch(b->on(b,flag)))return row;
        label(s,row++,3,col-1,name);
        for(var b:brands)if(on(b,flag)){put(s,row,3,col-1,t(b,"BRANDSYOSAI_NK"));for(int p=0;p<periods.size();p++){int c=col+p*step;put(s,row,c,c+step-1,num(b,key));}row++;}
        return row;
    }
    private void quote(Group g,String suffix){
        Sheet s=sheet("システム保守見積"+suffix,40,3);title(s,1,2,40,"システム保守見積");audit(s,3,25,29,40);
        put(s,2,35,40,t(g.brand(),"KEIYAKUJIKANTAI")+"H");
        put(s,3,2,23,t(h,"IRAIMEISHO1_NK")+" "+t(h,"IRAIMEISHO2_NK"));put(s,4,2,23,t(h,"IRAIMEISHO3_NK")+" "+t(h,"IRAIMEISHO4_NK"));put(s,5,2,23,brand(g));
        label(s,7,2,6,"納入先");put(s,7,7,40,h.get("NONYUSAKI_NK"));label(s,9,2,6,"適用期間");put(s,9,7,22,period(g));label(s,9,23,25,"有効期限");put(s,9,26,40,h.get("MITSUMORIKIGEN"));
        label(s,11,3,3,"No");label(s,11,4,20,"品名");label(s,11,21,24,"分類");label(s,11,25,27,"数量");label(s,11,28,40,"標準価格（円）");
        int row=12,n=0;
        for(var k:g.configurations()){
            put(s,row,3,3,++n);put(s,row,4,20,k.get("KIKIKOSEI_NK"));put(s,row,21,24,k.get("HOSYUHOHO"));put(s,row,25,27,k.get("SET_NM"));
            put(s,row,28,40,lines(g,k).stream().map(McmCustomerReportLayout::price).reduce(BigDecimal.ZERO,BigDecimal::add));row++;
        }
        int softRow=Math.max(36,row+1);
        String[] names={"システムサポート費","DTSサポート費","ダイフク技術料"};String[] keys={"HOSEISYSTEMSUPPORT_KIN","HOSEIDTSSUPPORT_KIN","HOSEIDAIFUKUGIJUTSU_KIN"};
        for(int i=0;i<3;i++){put(s,softRow+i,4,27,names[i]);put(s,softRow+i,28,40,num(g.brand(),keys[i]));}
        BigDecimal delta=num(g.brand(),"HOSEISOFTHOSHU_KIN").subtract(Arrays.stream(keys).map(k->num(g.brand(),k)).reduce(BigDecimal.ZERO,BigDecimal::add));
        put(s,softRow+3,4,27,"ソフト保守費補正差額");put(s,softRow+3,28,40,delta);
        put(s,softRow+4,4,27,"調整費");put(s,softRow+4,28,40,num(g.values(),"CHOSEI_KIN").negate());
        put(s,softRow+5,4,27,"DREMOS使用費");put(s,softRow+5,28,40,on(g.brand(),"DREMOS_FLG")?num(g.brand(),"DREMOS_KIN"):BigDecimal.ZERO);
        put(s,softRow+6,4,27,"リモート一次導入費");put(s,softRow+6,28,40,on(g.brand(),"REMOTE_FLG")?num(g.brand(),"REMOTE_KIN"):BigDecimal.ZERO);
        int total=softRow+7;label(s,total,4,27,"見積合計");formula(s,total,28,40,"SUM(AB12:AB"+(total-1)+")");quoteTotals.put(g,ref(s,total,28));
        int note=total+2;label(s,note++,2,40,"見積注記");note=paragraph(s,note,2,40,g.values().get("MITSUMORICHUKI"));label(s,note++,2,40,"保守方法");note=paragraph(s,note,2,40,g.brand().get("SOFTHOSYUHOHO"));label(s,note++,2,40,"備考");paragraph(s,note,2,40,g.values().get("BIKO"));
        s.setRepeatingRows(new CellRangeAddress(10,10,-1,-1));s.createFreezePane(3,11);
    }
    private void inventory(Group g,String suffix){
        Sheet s=sheet("保守仕様明細目録"+suffix,9,15);title(s,1,1,9,"システム保守仕様明細目録");label(s,2,7,8,"管理NO");put(s,2,9,9,t(h,"UM_KIHON_MITSUMORI_ID"));put(s,3,1,9,t(h,"NONYUSAKI_NK")+" / "+period(g));put(s,5,1,9,brand(g));
        int row=7;
        for(var k:displayConfigurations(g)){
            label(s,row++,1,9,t(k,"KIKIKOSEI_NK"));String[] labels={"品名","型式","契約数","総数","保守契約時間帯","点検回数","点検日程","保守方法","備考"};for(int c=0;c<9;c++)label(s,row,c+1,c+1,labels[c]);row++;
            for(Line l:lines(g,k)){var r=l.row();Object[] values={r.getMeisaiNm(),r.getKikikatashiki(),quantity(l),decimalOrBlank(r.getSosuNm()),l.hours(),r.getTenkenkaisu(),McmCustomerReportService.weekday(r.getTenkenyobi()),McmCustomerReportService.method(r.getHosyuhoho()),r.getBiko()};for(int c=0;c<9;c++)put(s,row,c+1,c+1,values[c]);row++;}
            for(var m:unselected(g,k)){put(s,row,1,1,m.get("KIKIHINMEI_NK"));put(s,row,2,2,m.get("KIKIKATASHIKI"));put(s,row,3,3,0);put(s,row,4,4,m.get("SOSU_NM"));put(s,row,9,9,m.get("BIKO"));row++;}
        }
        s.setColumnWidth(0,24*256);s.setColumnWidth(1,22*256);s.setColumnWidth(2,8*256);s.setColumnWidth(3,8*256);s.setColumnWidth(4,25*256);s.setColumnWidth(5,8*256);s.setColumnWidth(8,30*256);
        s.createFreezePane(2,6);
    }
    private static Object decimalOrBlank(String s){return s==null||s.isBlank()?null:new BigDecimal(s);}
    private void detailHeader(Sheet s,Group g,String name,boolean hard){
        int end=widths.get(s);title(s,1,2,end,name);
        int c=hard?4:5,labelEnd=c-1,split=hard?11:8,right=hard?16:11;
        label(s,2,2,labelEnd,"納入先");put(s,2,c,split,h.get("NONYUSAKI_NK"));label(s,2,split+1,right-1,"見積NO");put(s,2,right,end,h.get("UM_MITSUMORI_NO"));
        label(s,3,2,labelEnd,"プラント");put(s,3,c,split,h.get("PLANT_NK"));label(s,3,split+1,right-1,"管理NO");put(s,3,right,end,t(h,"UM_KIHON_MITSUMORI_ID"));
        label(s,4,2,labelEnd,"適用期間");put(s,4,c,split,period(g));put(s,4,split+1,right-1,day(h.get("MITSUMORI_DT")));put(s,4,right,end,h.get("MITSUMORISAKUSEISYA_NK"));
        put(s,5,2,end,brand(g)+" / 保守時間："+t(g.brand(),"KEIYAKUJIKANTAI")+"H");
    }
    private void hard(Group g,String suffix){
        Sheet s=sheet("ハード保守明細"+suffix,16,12);detailHeader(s,g,"ハードウェア保守費明細",true);s.setColumnWidth(0,3*256);s.setColumnWidth(1,5*256);s.setColumnWidth(3,2*256);s.setColumnWidth(4,24*256);s.setColumnWidth(5,20*256);s.setColumnWidth(15,25*256);
        int row=6,n=0;var totals=new ArrayList<Integer>();
        for(var k:displayConfigurations(g)){
            if(on(k,"CONTROLLER_FLG"))continue;
            label(s,row++,2,16,t(k,"KIKIKOSEI_NK"));
            int[] cols={2,3,5,6,7,8,9,10,11,12,13,14,15,16};String[] labels={"No","メーカー","品名","型式","契約数","総数","仕切価格","発注先","仕切（%）","定価単価","定価価格合計","契約仕切（%）","契約価格","備考"};
            for(int i=0;i<cols.length;i++)label(s,row,cols[i],i==1?4:cols[i],labels[i]);row++;int first=row;
            for(Line l:lines(g,k)){if("1".equals(l.row().getControllerFlg()))continue;var r=l.row();put(s,row,2,2,++n);put(s,row,3,4,r.getSeizomakerNk());put(s,row,5,5,r.getMeisaiNm());put(s,row,6,6,r.getKikikatashiki());put(s,row,7,7,quantity(l));put(s,row,8,8,decimalOrBlank(r.getSosuNm()));put(s,row,9,9,cost(l));put(s,row,10,10,l.supplier());put(s,row,12,12,r.getHyojunKin());formula(s,row,13,13,"G"+row+"*L"+row);ratio(s,row,11,"I"+row,"M"+row);put(s,row,16,16,r.getBiko());row++;}
            for(var m:unselected(g,k)){put(s,row,2,2,++n);put(s,row,3,4,m.get("SEIZOMAKER_NK"));put(s,row,5,5,m.get("KIKIHINMEI_NK"));put(s,row,6,6,m.get("KIKIKATASHIKI"));put(s,row,7,7,0);put(s,row,8,8,m.get("SOSU_NM"));put(s,row,9,9,"-");put(s,row,13,13,"-");put(s,row,16,16,m.get("BIKO"));row++;}
            label(s,row,2,8,"小計");formula(s,row,9,10,row==first?"0":"SUM(I"+first+":I"+(row-1)+")");formula(s,row,12,13,row==first?"0":"SUM(M"+first+":M"+(row-1)+")");ratio(s,row,11,"I"+row,"L"+row);put(s,row,14,15,"-");totals.add(row);row+=2;
        }
        label(s,row,2,8,"合計");formula(s,row,9,10,sumCells(totals,"I"));formula(s,row,12,13,sumCells(totals,"L"));s.createFreezePane(6,5);
        row+=2;label(s,row++,2,16,"取引先別費用一覧");label(s,row,2,8,"取引先");label(s,row,9,11,"仕切原価");label(s,row++,12,16,"標準価格");
        for(var supplier:suppliers(g.lines().stream().filter(l->!"1".equals(l.row().getControllerFlg())).toList()).values()){
            put(s,row,2,8,supplierName(supplier.get(0)));put(s,row,9,11,supplier.stream().map(McmCustomerReportLayout::cost).reduce(BigDecimal.ZERO,BigDecimal::add));put(s,row,12,16,supplier.stream().map(McmCustomerReportLayout::price).reduce(BigDecimal.ZERO,BigDecimal::add));row++;
        }
    }
    private static Map<String,List<Line>> suppliers(List<Line> lines){
        Map<String,List<Line>> groups=new LinkedHashMap<>();for(Line l:lines)groups.computeIfAbsent(l.supplierId(),key->new ArrayList<>()).add(l);return groups;
    }
    private static String supplierName(Line line){return line.supplierId().isEmpty()?"社内対応":line.supplier();}
    private static String sumCells(List<Integer> rows,String column){return rows.isEmpty()?"0":"SUM("+String.join(",",rows.stream().map(r->column+r).toList())+")";}
    private void cost(Group g,String suffix){
        Sheet s=sheet("見積原価表"+suffix,12,14);s.setColumnWidth(0,3*256);s.setColumnWidth(1,5*256);s.setColumnWidth(2,24*256);s.setColumnWidth(3,20*256);s.setColumnWidth(4,5*256);s.setColumnWidth(5,8*256);s.setColumnWidth(11,25*256);
        title(s,2,2,12,"見積原価表");
        label(s,3,2,4,"納入先");put(s,3,5,9,h.get("NONYUSAKI_NK"));label(s,3,10,10,"見積NO");put(s,3,11,12,h.get("UM_MITSUMORI_NO"));
        label(s,4,2,4,"プラント");put(s,4,5,8,h.get("PLANT_NK"));put(s,4,9,10,t(g.brand(),"KEIYAKUJIKANTAI")+"H");put(s,4,11,12,t(h,"UM_KIHON_MITSUMORI_ID"));
        label(s,5,2,4,"適用期間");put(s,5,5,8,period(g));put(s,5,9,10,day(h.get("MITSUMORI_DT")));put(s,5,11,12,h.get("MITSUMORISAKUSEISYA_NK"));put(s,6,2,12,brand(g));
        label(s,8,3,3,"項目");label(s,8,4,5,"メーカー");String[] labels={"数量","コスト金額","発注先","プライス","契約金額","原価率","備考"};for(int i=0;i<labels.length;i++)label(s,8,i+6,i+6,labels[i]);
        int row=9;
        for(boolean controller:new boolean[]{false,true}){
            var category=g.lines().stream().filter(l->"1".equals(l.row().getControllerFlg())==controller).toList();if(category.isEmpty())continue;
            label(s,row++,3,12,controller?"コントローラ保守費":"ハードウェア保守費");
            for(var supplier:suppliers(category).values()){
                int supplierRow=row;put(s,row,3,6,supplierName(supplier.get(0)));put(s,row++,9,9,supplier.stream().map(McmCustomerReportLayout::price).reduce(BigDecimal.ZERO,BigDecimal::add));
                Map<String,List<Line>> makers=new LinkedHashMap<>();for(Line l:supplier)makers.computeIfAbsent(str(l.row().getSeizomakerNk()),key->new ArrayList<>()).add(l);
                for(var entry:makers.entrySet()){
                    put(s,row,4,5,entry.getKey());put(s,row,6,6,entry.getValue().stream().map(McmCustomerReportLayout::quantity).reduce(BigDecimal.ZERO,BigDecimal::add));put(s,row,7,7,entry.getValue().stream().map(McmCustomerReportLayout::cost).reduce(BigDecimal.ZERO,BigDecimal::add));put(s,row,8,8,supplierName(entry.getValue().get(0)));row++;
                }
                // 数量・原価はメーカー行、プライスは取引先行のみへ出力し、合計で重複加算しない。
                ratio(s,supplierRow,11,"SUM(G"+(supplierRow+1)+":G"+(row-1)+")","I"+supplierRow);
            }
        }
        String[] names={"システムサポート費","DTSサポート費","ダイフク技術料"},keys={"HOSEISYSTEMSUPPORT_KIN","HOSEIDTSSUPPORT_KIN","HOSEIDAIFUKUGIJUTSU_KIN"};
        for(int i=0;i<3;i++){put(s,row,3,5,names[i]);put(s,row,7,7,0);put(s,row,8,8,"社内対応");put(s,row,9,9,num(g.brand(),keys[i]));row++;}
        BigDecimal delta=num(g.brand(),"HOSEISOFTHOSHU_KIN").subtract(Arrays.stream(keys).map(k->num(g.brand(),k)).reduce(BigDecimal.ZERO,BigDecimal::add));
        put(s,row,3,5,"ソフト保守費補正差額");put(s,row++,9,9,delta);
        put(s,row,3,5,"調整費");put(s,row++,9,9,num(g.values(),"CHOSEI_KIN").negate());
        put(s,row,3,5,"DREMOS使用費");put(s,row,7,7,on(g.brand(),"DREMOS_FLG")?34800:0);put(s,row,8,8,"NTT");put(s,row++,9,9,on(g.brand(),"DREMOS_FLG")?num(g.brand(),"DREMOS_KIN"):0);
        put(s,row,3,5,"リモート一次導入費");put(s,row,7,7,0);put(s,row++,9,9,on(g.brand(),"REMOTE_FLG")?num(g.brand(),"REMOTE_KIN"):0);
        int total=Math.max(59,row+1);label(s,total,3,6,"合計");formula(s,total,7,7,"SUM(G9:G"+(row-1)+")");formula(s,total,9,9,"SUM(I9:I"+(row-1)+")");ratio(s,total,11,"G"+total,"I"+total);costTotals.put(g,ref(s,total,7));s.setRepeatingRows(new CellRangeAddress(7,7,-1,-1));s.createFreezePane(3,8);
    }
    private void soft(Group g,String suffix){
        Sheet s=sheet("ソフト保守費明細"+suffix,12,12);title(s,1,2,12,"ソフトウェア保守費明細");
        label(s,2,1,2,"納入先");put(s,2,3,8,h.get("NONYUSAKI_NK"));label(s,2,9,9,"見積NO");put(s,2,10,12,h.get("UM_MITSUMORI_NO"));
        label(s,3,1,2,"プラント");put(s,3,3,8,h.get("PLANT_NK"));label(s,3,9,9,"管理NO");put(s,3,10,12,t(h,"UM_KIHON_MITSUMORI_ID"));
        label(s,4,1,2,"適用期間");put(s,4,3,6,period(g));put(s,4,7,9,day(h.get("MITSUMORI_DT")));put(s,4,10,12,h.get("MITSUMORISAKUSEISYA_NK"));put(s,5,2,12,brand(g));
        label(s,7,2,6,"費目");label(s,7,7,9,"基準額（円）");label(s,7,10,12,"補正額（円）");
        String[] names={"システム設計費","基本設計費","プログラム作成費","システムサポート費","DTSサポート費","ダイフク技術料","ソフト保守費合計"};String[] keys={"SYSTEMSEKKEI_KIN","KIHONSEKKEI_KIN","PROGRAMSAKUSEI_KIN","SYSTEMSUPPORT_KIN","DTSSUPPORT_KIN","DAIFUKUGIJUTSU_KIN","SOFTHOSHU_KIN"};
        for(int i=0;i<names.length;i++){put(s,i+8,2,6,names[i]);put(s,i+8,7,9,g.brand().get(keys[i]));if(i>=3)put(s,i+8,10,12,g.brand().get("HOSEI"+keys[i]));}
        label(s,16,2,12,"保守方法");paragraph(s,17,2,12,g.brand().get("SOFTHOSYUHOHO"));s.setRepeatingRows(new CellRangeAddress(6,6,-1,-1));
    }
    private void finish(Sheet s){
        for(Row row:s)for(Cell cell:row)if(cell.getCellStyle()==number||cell.getCellStyle().getIndex()==number.getIndex()){
            if(cell.getCellType()!=CellType.NUMERIC&&!(cell.getCellType()==CellType.FORMULA&&cell.getCachedFormulaResultType()==CellType.NUMERIC))continue;
            int scale=Math.max(0,Math.min(6,BigDecimal.valueOf(cell.getNumericCellValue()).stripTrailingZeros().scale()));
            cell.setCellStyle(precision.computeIfAbsent(scale,k->{CellStyle style=book.createCellStyle();style.cloneStyleFrom(number);String fmt="#,##0"+(k==0?"":"."+"0".repeat(k));style.setDataFormat(book.createDataFormat().getFormat(fmt+";[Red]-"+fmt));return style;}));
        }
        s.setFitToPage(true);s.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);s.getPrintSetup().setLandscape(true);s.getPrintSetup().setFitWidth((short)1);s.getPrintSetup().setFitHeight((short)0);
        s.setMargin(Sheet.LeftMargin,0.25);s.setMargin(Sheet.RightMargin,0.25);s.setMargin(Sheet.TopMargin,0.35);s.setMargin(Sheet.BottomMargin,0.35);s.getFooter().setRight("Page &P / &N");
        book.setPrintArea(book.getSheetIndex(s),0,widths.get(s)-1,0,s.getLastRowNum());
    }
}
