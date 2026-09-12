package com.daifuku.mcm.service;

import java.math.BigDecimal;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

/** SP_UM/SP_UKの整合処理。画面更新は対象伝票内に限定し、全件の孤児削除は運用SQLに分離する。 */
@Service
public class McmCustomerIntegrityService {
    private final JdbcTemplate jdbc;
    @Value("${mcm.customer.integrity-mode:sql}") private String mode="sql";
    public McmCustomerIntegrityService(JdbcTemplate jdbc){this.jdbc=jdbc;}
    @Transactional
    public void estimate(BigDecimal id){
        require("SELECT COUNT(*) FROM MCM.MCM_UM_KIHON_BRAND WHERE UM_KIHON_MITSUMORI_ID=?",id);
        require("SELECT COUNT(*) FROM MCM.MCM_UM_MITSUMORI WHERE UM_KIHON_MITSUMORI_ID=?",id);
        if("procedure".equals(mode)){jdbc.update("EXEC MCM.MCM_WEB_SP_UM ?",id);return;}
        if(!"sql".equals(mode))throw new IllegalStateException("整合処理の設定を確認してください。");
        jdbc.update("DELETE FROM MCM.MCM_UM_TANKA WHERE UM_KIKIMEISAI_ID IN (SELECT M.UM_KIKIMEISAI_ID FROM MCM.MCM_UM_KIKIMEISAI M JOIN MCM.MCM_UM_KIKIKOSEI K ON K.UM_KIKIKOSEI_ID=M.UM_KIKIKOSEI_ID JOIN MCM.MCM_UM_KIHON_BRAND B ON B.UM_KIHON_BRAND_ID=K.UM_KIHON_BRAND_ID WHERE B.UM_KIHON_MITSUMORI_ID=?) AND NOT EXISTS (SELECT 1 FROM MCM.MCM_UM_MITSUMORI P WHERE P.UM_MITSUMORI_ID=MCM_UM_TANKA.UM_MITSUMORI_ID AND P.UM_KIHON_MITSUMORI_ID=?)",id,id);
        jdbc.update("DELETE FROM MCM.MCM_UM_BRAND WHERE UM_KIHON_BRAND_ID IN (SELECT UM_KIHON_BRAND_ID FROM MCM.MCM_UM_KIHON_BRAND WHERE UM_KIHON_MITSUMORI_ID=?) AND NOT EXISTS (SELECT 1 FROM MCM.MCM_UM_MITSUMORI P WHERE P.UM_MITSUMORI_ID=MCM_UM_BRAND.UM_MITSUMORI_ID AND P.UM_KIHON_MITSUMORI_ID=?)",id,id);
    }
    @Transactional
    public void contract(BigDecimal id){
        require("SELECT COUNT(*) FROM MCM.MCM_UK_KIKAN WHERE UK_KEIYAKU_ID=?",id);
        if("procedure".equals(mode)){jdbc.update("EXEC MCM.MCM_WEB_SP_UK ?",id);return;}
        if(!"sql".equals(mode))throw new IllegalStateException("整合処理の設定を確認してください。");
        // 期間・機器の削除は Mcm2006uRepository.deleteEquipment が子から処理済み。
    }
    private void require(String sql,BigDecimal id){if(id==null||jdbc.queryForObject(sql,Integer.class,id)==0)throw new IllegalStateException("保存対象の親子情報が不足しています。再選定してください。");}
}
