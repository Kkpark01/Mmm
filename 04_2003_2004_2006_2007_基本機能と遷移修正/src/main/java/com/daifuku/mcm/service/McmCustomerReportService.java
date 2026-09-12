package com.daifuku.mcm.service;

import java.io.*;
import java.math.BigDecimal;
import java.util.*;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.*;
import com.daifuku.mcm.repository.Mcm2005uRepository;
import com.daifuku.mcm.form.Mcm2005uForm.TankaRowForm;
import static com.daifuku.mcm.service.McmCustomerTotalsService.*;

/** 保存済みの同一スナップショットから店舗見積帳票を生成する。 */
@Service
public class McmCustomerReportService {
    private final JdbcTemplate jdbc;
    private final Mcm2005uRepository details;
    public McmCustomerReportService(JdbcTemplate jdbc,Mcm2005uRepository details){this.jdbc=jdbc;this.details=details;}

    record Line(TankaRowForm row,String supplier,String hours,String supplierId) {}
    record Group(Map<String,Object> brand,Map<String,Object> period,Map<String,Object> values,
                 List<Map<String,Object>> configurations,List<Line> lines,List<Map<String,Object>> unselected) {}

    @Transactional(readOnly=true,isolation=Isolation.REPEATABLE_READ)
    public byte[] generate(BigDecimal id){
        var headers=jdbc.queryForList("SELECT * FROM MCM.MCM_UM_KIHON_MITSUMORI WHERE UM_KIHON_MITSUMORI_ID=?",id);
        if(headers.size()!=1)throw new IllegalStateException("見積を選択してください。");
        var periods=jdbc.queryForList("SELECT * FROM MCM.MCM_UM_MITSUMORI WHERE UM_KIHON_MITSUMORI_ID=? ORDER BY KAISI_DT,UM_MITSUMORI_ID",id);
        var brands=jdbc.queryForList("SELECT B.*,M.DREMOS_KIN,M.REMOTE_KIN FROM MCM.MCM_UM_KIHON_BRAND B LEFT JOIN MCM.MCM_MA_BRAND M ON M.BRAND_ID=B.BRAND_ID WHERE B.UM_KIHON_MITSUMORI_ID=? ORDER BY B.UM_KIHON_BRAND_ID",id);
        if(periods.isEmpty()||brands.isEmpty())throw new IllegalStateException("帳票に出力するブランド・期間がありません。");
        var groups=new ArrayList<Group>();
        for(var b:brands){
            var configs=jdbc.queryForList("SELECT * FROM MCM.MCM_UM_KIKIKOSEI WHERE UM_KIHON_BRAND_ID=? ORDER BY HYOJIJUN,UM_KIKIKOSEI_ID",b.get("UM_KIHON_BRAND_ID"));
            for(var p:periods){
                var values=jdbc.queryForList("SELECT * FROM MCM.MCM_UM_BRAND WHERE UM_KIHON_BRAND_ID=? AND UM_MITSUMORI_ID=?",b.get("UM_KIHON_BRAND_ID"),p.get("UM_MITSUMORI_ID"));
                if(values.size()!=1)throw new IllegalStateException("帳票のブランド期間情報が不足しています。");
                var lines=new ArrayList<Line>();
                for(var r:details.findTankaRows(num(p,"UM_MITSUMORI_ID"),num(b,"UM_KIHON_BRAND_ID"))){
                    var source=jdbc.queryForList("SELECT R.TORIHIKISAKI_ID,R.TORIHIKISAKI_NK FROM MCM.MCM_UM_TANKA U LEFT JOIN MCM.MCM_MA_TORIHOSYUJIKAN J ON J.TORIHOSYUJIKAN_ID=U.TORIHOSYUJIKAN_ID LEFT JOIN MCM.MCM_MA_TORIHIKISAKI R ON R.TORIHIKISAKI_ID=J.TORIHIKISAKI_ID WHERE U.UM_TANKA_ID=?",r.getUmTankaId());
                    lines.add(new Line(r,source.isEmpty()?"":text(source.get(0),"TORIHIKISAKI_NK"),hours(r.getDaifukuhosyujikanId()),source.isEmpty()?"":text(source.get(0),"TORIHIKISAKI_ID")));
                }
                // VB MCM_MA_KOTAIMEISAI_VTableAdapter.Fill と同じマスター掲載条件。
                var unselected=jdbc.queryForList("""
                    SELECT V.MAE_KIKIKOSEI_ID AS KIKIKOSEI_ID,V.MAE_KIKIKOSEI_NK AS KIKIKOSEI_NK,
                           V.MAF_KIKIMEISAI_ID AS KIKIMEISAI_ID,V.MAJ_SEIZOMAKER_NK AS SEIZOMAKER_NK,
                           V.MAF_KIKIHINMEI_NK AS KIKIHINMEI_NK,V.MAF_KIKIKATASHIKI AS KIKIKATASHIKI,
                           SUM(ISNULL(V.MAF_SURYO_NM,0)) AS SOSU_NM,V.MAE_CONTROLLER_FLG AS CONTROLLER_FLG,V.MAF_BIKO AS BIKO
                    FROM MCM.MCM_MA_KOTAIMEISAI_V V
                    WHERE V.MAE_PLANT_ID=? AND V.MAG_BRANDKOSEI_ID=?
                      AND V.MAH_MITSUMORIHYOJI_FLG='0' AND V.MAH_DUMMY_FLG='0' AND V.MAF_NOUNYU_KBN IS NULL
                      AND NOT EXISTS (SELECT 1 FROM MCM.MCM_UM_TANKA T
                          JOIN MCM.MCM_UM_KIKIMEISAI M ON M.UM_KIKIMEISAI_ID=T.UM_KIKIMEISAI_ID
                          JOIN MCM.MCM_UM_KIKIKOSEI K ON K.UM_KIKIKOSEI_ID=M.UM_KIKIKOSEI_ID
                          WHERE K.UM_KIHON_BRAND_ID=? AND T.UM_MITSUMORI_ID=?
                            AND M.KIKIMEISAI_ID=V.MAF_KIKIMEISAI_ID AND K.KIKIKOSEI_ID=V.MAE_KIKIKOSEI_ID)
                    GROUP BY V.MAE_KIKIKOSEI_ID,V.MAE_KIKIKOSEI_NK,V.MAF_KIKIMEISAI_ID,V.MAJ_SEIZOMAKER_NK,
                             V.MAF_KIKIHINMEI_NK,V.MAF_KIKIKATASHIKI,V.MAE_CONTROLLER_FLG,V.MAF_BIKO,V.MAE_HYOJIJUN,V.MAF_HYOJIJUN
                    ORDER BY V.MAE_HYOJIJUN,V.MAF_HYOJIJUN,V.MAE_KIKIKOSEI_ID,V.MAF_KIKIMEISAI_ID
                    """,headers.get(0).get("PLANT_ID"),b.get("BRANDKOSEI_ID"),b.get("UM_KIHON_BRAND_ID"),p.get("UM_MITSUMORI_ID"));
                groups.add(new Group(b,p,values.get(0),configs,lines,unselected));
            }
        }
        try { return new McmCustomerReportLayout().render(headers.get(0),periods,brands,groups); }
        catch(IOException e){throw new IllegalStateException("帳票を作成できませんでした。再度発行してください。",e);}
    }
    private String hours(BigDecimal id){if(id==null)return "";var rows=jdbc.queryForList("SELECT HOSYUJIKAN_DT,KAISIYOBI,SYURYOYOBI,KAISIJIKAN_DT,SYURYOJIKAN_DT FROM MCM.MCM_MA_TORIHOSYUJIKAN WHERE TORIHOSYUJIKAN_ID=?",id);if(rows.isEmpty())return "ID "+id.stripTrailingZeros().toPlainString();var r=rows.get(0);return text(r,"HOSYUJIKAN_DT")+"時間 "+dayName(text(r,"KAISIYOBI"))+"～"+dayName(text(r,"SYURYOYOBI"))+" "+text(r,"KAISIJIKAN_DT")+"～"+text(r,"SYURYOJIKAN_DT");}
    public static String weekday(String c){return Map.of("1","月～金","2","月～土","3","月～日","4","土・日").getOrDefault(textValue(c),textValue(c));}
    public static String method(String c){return Map.of("F","オンサイト","S","センドバック","C","コンテック製品","H","持ち帰り","I","スポット","T","TEL対応").getOrDefault(textValue(c),textValue(c));}
    private static String dayName(String c){return Map.of("1","月","2","火","3","水","4","木","5","金","6","土","7","日").getOrDefault(c,c);}
    private static String text(Map<String,Object> r,String key){return textValue(r.get(key));}
    private static String textValue(Object v){return v==null?"":v instanceof BigDecimal n?n.stripTrailingZeros().toPlainString():v.toString();}
    private static String date(Object v){return v==null?"":v.toString().substring(0,10).replace('-','/');}

}
