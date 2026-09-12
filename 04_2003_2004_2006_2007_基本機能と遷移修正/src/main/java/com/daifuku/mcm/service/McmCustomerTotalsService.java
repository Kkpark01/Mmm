package com.daifuku.mcm.service;

import java.math.BigDecimal;
import java.util.*;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import com.daifuku.mcm.repository.Mcm2005uRepository;

/** VBの数量×単価、ブランド期間保守費、期間見積合計を同一トランザクションで計算する。 */
@Service
public class McmCustomerTotalsService {
    private final JdbcTemplate jdbc;
    private final Mcm2005uRepository details;
    public McmCustomerTotalsService(JdbcTemplate jdbc,Mcm2005uRepository details){this.jdbc=jdbc;this.details=details;}

    @Transactional
    public void recalculate(BigDecimal id,String user){
        jdbc.queryForList("SELECT UM_KIHON_MITSUMORI_ID FROM MCM.MCM_UM_KIHON_MITSUMORI WITH (UPDLOCK,HOLDLOCK) WHERE UM_KIHON_MITSUMORI_ID=?",id);
        var periods=jdbc.queryForList("SELECT UM_MITSUMORI_ID FROM MCM.MCM_UM_MITSUMORI WHERE UM_KIHON_MITSUMORI_ID=? ORDER BY KAISI_DT,UM_MITSUMORI_ID",id);
        var brands=jdbc.queryForList("SELECT B.*,M.DREMOS_KIN,M.REMOTE_KIN FROM MCM.MCM_UM_KIHON_BRAND B LEFT JOIN MCM.MCM_MA_BRAND M ON M.BRAND_ID=B.BRAND_ID WHERE B.UM_KIHON_MITSUMORI_ID=?",id);
        for(var p:periods){
            BigDecimal total=BigDecimal.ZERO,hardTotal=BigDecimal.ZERO;
            for(var b:brands){
                var rows=details.findTankaRows(num(p,"UM_MITSUMORI_ID"),num(b,"UM_KIHON_BRAND_ID"));
                BigDecimal hard=BigDecimal.ZERO,controller=BigDecimal.ZERO;
                for(var r:rows){BigDecimal amount=nz(r.getHyojunKin()).multiply(new BigDecimal(r.getSuryoNm()));if("1".equals(r.getControllerFlg()))controller=controller.add(amount);else hard=hard.add(amount);}
                var bp=jdbc.queryForList("SELECT CHOSEI_KIN FROM MCM.MCM_UM_BRAND WHERE UM_KIHON_BRAND_ID=? AND UM_MITSUMORI_ID=?",b.get("UM_KIHON_BRAND_ID"),p.get("UM_MITSUMORI_ID"));
                if(bp.size()!=1)throw new IllegalStateException("ブランドと期間の組合せが不足しています。機器選定を確認してください。");
                BigDecimal hoshu=hard.add(num(b,"HOSEISOFTHOSHU_KIN")).subtract(num(bp.get(0),"CHOSEI_KIN"));
                jdbc.update("UPDATE MCM.MCM_UM_BRAND SET HARDHOSYU_KIN=?,HOSHU_KIN=?,LASTUPDATE_DT=GETDATE(),LASTUPDATE_BY=? WHERE UM_KIHON_BRAND_ID=? AND UM_MITSUMORI_ID=?",hard,hoshu,user,b.get("UM_KIHON_BRAND_ID"),p.get("UM_MITSUMORI_ID"));
                hardTotal=hardTotal.add(hard);total=total.add(hoshu).add(controller);
                if(num(b,"DREMOS_FLG").intValue()==1)total=total.add(num(b,"DREMOS_KIN"));
                if(num(b,"REMOTE_FLG").intValue()==1)total=total.add(num(b,"REMOTE_KIN"));
            }
            jdbc.update("UPDATE MCM.MCM_UM_MITSUMORI SET HARDHOSYU_KIN=?,MITSUMORI_GKIN=?,LASTUPDATE_DT=GETDATE(),LASTUPDATE_BY=? WHERE UM_MITSUMORI_ID=?",hardTotal,total,user,p.get("UM_MITSUMORI_ID"));
        }
        // 手入力された補正ソフト保守費は維持し、基本ハード保守費は期間最大値を反映する。
        for(var b:brands)jdbc.update("UPDATE MCM.MCM_UM_KIHON_BRAND SET HARDHOSYU_KIN=(SELECT MAX(HARDHOSYU_KIN) FROM MCM.MCM_UM_BRAND WHERE UM_KIHON_BRAND_ID=?),LASTUPDATE_DT=GETDATE(),LASTUPDATE_BY=? WHERE UM_KIHON_BRAND_ID=?",b.get("UM_KIHON_BRAND_ID"),user,b.get("UM_KIHON_BRAND_ID"));
        jdbc.update("UPDATE MCM.MCM_UM_KIHON_MITSUMORI SET LASTUPDATE_DT=GETDATE(),LASTUPDATE_BY=? WHERE UM_KIHON_MITSUMORI_ID=?",user,id);
    }

    public void checkHours(BigDecimal supplier,BigDecimal daifuku){
        if(daifuku==null||daifuku.signum()==0){if(supplier!=null&&supplier.signum()>0)throw new IllegalStateException("ダイフク保守時間を選択してください。");return;}
        var own=hours(daifuku);if(supplier==null||supplier.signum()==0)return;
        var source=hours(supplier);
        for(int d=0;d<7;d++)if(own[d][0]==1&&(source[d][0]!=1||own[d][1]<source[d][1]||own[d][2]>source[d][2]))
            throw new IllegalStateException("ダイフク保守時間は取引先の保守時間内に設定してください。");
    }
    private int[][] hours(BigDecimal id){
        var list=jdbc.queryForList("SELECT * FROM MCM.MCM_MA_TORIHOSYUJIKAN WHERE TORIHOSYUJIKAN_ID=?",id);
        if(list.size()!=1)throw new IllegalStateException("保守時間マスタを確認してください。");
        var m=list.get(0);int start=num(m,"KAISIYOBI").intValue(),end=num(m,"SYURYOYOBI").intValue();
        if(start<1||start>7||end<1||end>7)throw new IllegalStateException("保守時間マスタの曜日を確認してください。");
        int[][] a=new int[7][3];
        for(int d=start,i=0;i<7;i++,d=d%7+1){a[d-1]=new int[]{1,num(m,"HOSYUJIKAN_DT").intValue()==24?0:num(m,"KAISIJIKAN_DT").intValue(),num(m,"HOSYUJIKAN_DT").intValue()==24?2400:num(m,"SYURYOJIKAN_DT").intValue()};if(d==end)break;}
        int exception=num(m,"REIGAIYOBI").intValue();if(exception>0&&exception<=7)a[exception-1]=new int[]{1,num(m,"REIGAIKAISIJIKAN_DT").intValue(),num(m,"REIGAISYURYOJIKAN_DT").intValue()};return a;
    }
    static BigDecimal nz(BigDecimal v){return v==null?BigDecimal.ZERO:v;}
    static BigDecimal num(Map<String,Object> r,String key){return r.get(key)==null?BigDecimal.ZERO:new BigDecimal(r.get(key).toString());}
}
