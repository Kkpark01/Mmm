/**
 * 【変換元】Mcm0010uScreen.vb
 * 【説明】納入先マスタ検索画面のサービスクラス
 * 元メソッド対応:
 *   SEARCHButton_Click → search()
 *   RowDeleteButton_Click → deletePlant()
 *   DataRelation連動 → getPlantsByNonyusakiId(), getBrandsByPlantId(), getKikiByPlantId()
 */
package com.daifuku.mcm.service;

import com.daifuku.mcm.dto.Mcm0010uBrandKoseiDto;
import com.daifuku.mcm.dto.Mcm0010uKikiKoseiDto;
import com.daifuku.mcm.entity.NonyusakiEntity;
import com.daifuku.mcm.entity.PlantEntity;
import com.daifuku.mcm.form.Mcm0010uForm;
import com.daifuku.mcm.repository.Mcm2004uJdbcRepository;
import org.springframework.jdbc.core.BeanPropertyRowMapper;
import org.springframework.transaction.annotation.Isolation;
import java.util.LinkedHashSet;
import java.util.Objects;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Service
public class Mcm0010uService {

    @Autowired
    private Mcm2004uService permissions;

    @Autowired
    private Mcm2004uJdbcRepository userQuotes;

    @Autowired
    private JdbcTemplate jdbcTemplate;

    /**
     * 検索条件が全て空白かチェックする
     * 元VB: IsNull(NONYUSAKI_CDTextBox.Text) And IsNull(...) And ...
     */
    public boolean isAllEmpty(Mcm0010uForm form) {
        return isBlank(form.getNonyusakiCd())
            && isBlank(form.getSupportId())
            && isBlank(form.getNonyusakiNk())
            && isBlank(form.getPlantNk())
            && isBlank(form.getJusyo());
    }

    /**
     * 納入先マスタ検索
     * 元VB: Me.Fill(Me.Mcm0010uDataSet.MCM_MA_NONYUSAKI, ...)
     */
    public List<NonyusakiEntity> searchNonyusaki(Mcm0010uForm form) {
        var sql = new StringBuilder("SELECT MAA.* FROM MCM.MCM_MA_NONYUSAKI MAA WHERE 1=1 ");
        var args = new ArrayList<Object>();
        appendNonyusaki(sql, args, form);
        if (!isBlank(form.getSupportId())) { sql.append(" AND EXISTS (SELECT 1 FROM MCM.MCM_MA_PLANT P WHERE P.NONYUSAKI_ID=MAA.NONYUSAKI_ID AND P.SUPPORT_ID LIKE ?)"); args.add(prefix(form.getSupportId())); }
        if (!isBlank(form.getPlantNk())) { sql.append(" AND EXISTS (SELECT 1 FROM MCM.MCM_MA_PLANT P WHERE P.NONYUSAKI_ID=MAA.NONYUSAKI_ID AND UPPER(P.PLANT_NK) LIKE UPPER(?))"); args.add(contains(form.getPlantNk())); }
        var result=jdbcTemplate.query(sql.append(" ORDER BY MAA.NONYUSAKI_ID").toString(),new BeanPropertyRowMapper<>(NonyusakiEntity.class),args.toArray());
        result.forEach(e -> e.setPlanttsuika("プラント追加")); return result;
    }

    /**
     * プラントマスタ検索
     * 元VB: Me.Fill(Me.Mcm0010uDataSet.MCM_MA_PLANT, ...)
     * @param authorityDivision 権限区分 ("INSPECTION" or "UPDATE")
     */
    public List<PlantEntity> searchPlants(Mcm0010uForm form, String authorityDivision) {
        var sql = new StringBuilder("SELECT MAB.*, CASE WHEN ?='2' OR EXISTS (SELECT 1 FROM MCM.MCM_MA_KIKIKOSEI K WHERE K.PLANT_ID=MAB.PLANT_ID) THEN N'機器構成' ELSE '' END AS KIKIKOSEILINK FROM MCM.MCM_MA_PLANT MAB WHERE 1=1 ");
        var args=new ArrayList<Object>(); args.add("UPDATE".equals(authorityDivision)?"2":authorityDivision);
        appendExistsConditions(sql,args,form,"MAB");
        return jdbcTemplate.query(sql.append(" ORDER BY MAB.PLANT_ID").toString(),new BeanPropertyRowMapper<>(PlantEntity.class),args.toArray());
    }

    /**
     * 納入先IDでプラントを取得（リレーション連動用）
     * 元VB: DataRelation MCM_MA_PLANT_MCM_MA_NONYUSAKI
     */
    public List<PlantEntity> getPlantsByNonyusakiId(BigDecimal nonyusakiId) {
        return jdbcTemplate.query("SELECT * FROM MCM.MCM_MA_PLANT WHERE NONYUSAKI_ID=? ORDER BY PLANT_ID",new BeanPropertyRowMapper<>(PlantEntity.class),nonyusakiId);
    }

    /**
     * ブランド構成マスタ検索
     * 元VB SQL (Mcm0010u_MCM_MA_BRAND_KOSEITableAdapter.xml):
     *   SELECT COUNT(*) AS SURYO, MAC.BRAND_NK, MAD.PLANT_ID, MAD.BRAND_ID,
     *          MAD.NOUNYU_KBN, MAD.BRANDSYOSAI_NK
     *   FROM MCM_MA_BRAND_KOSEI MAD, MCM_MA_BRAND MAC
     *   WHERE MAD.BRAND_ID = MAC.BRAND_ID
     *   GROUP BY MAC.BRAND_NK, MAD.PLANT_ID, MAD.BRAND_ID,
     *            MAD.NOUNYU_KBN, MAD.BRANDSYOSAI_NK, MAC.HYOJIJUN
     *   ORDER BY MAC.HYOJIJUN ASC
     * + EXISTS条件で検索パラメータ絞り込み
     */
    public List<Mcm0010uBrandKoseiDto> searchBrandKosei(Mcm0010uForm form) {
        StringBuilder sql = new StringBuilder();
        sql.append("SELECT COUNT(*) AS SURYO, MAC.BRAND_NK, MAD.PLANT_ID, MAD.BRAND_ID, ");
        sql.append("MAD.NOUNYU_KBN, MAD.BRANDSYOSAI_NK ");
        sql.append("FROM MCM.MCM_MA_BRAND_KOSEI MAD ");
        sql.append("INNER JOIN MCM.MCM_MA_BRAND MAC ON MAD.BRAND_ID = MAC.BRAND_ID ");
        sql.append("WHERE 1=1 ");

        List<Object> params = new ArrayList<>();
        appendExistsConditions(sql, params, form, "MAD");

        sql.append("GROUP BY MAC.BRAND_NK, MAD.PLANT_ID, MAD.BRAND_ID, ");
        sql.append("MAD.NOUNYU_KBN, MAD.BRANDSYOSAI_NK, MAC.HYOJIJUN ");
        sql.append("ORDER BY MAC.HYOJIJUN ASC");

        List<Map<String, Object>> rows = jdbcTemplate.queryForList(sql.toString(), params.toArray());
        List<Mcm0010uBrandKoseiDto> result = new ArrayList<>();
        for (Map<String, Object> row : rows) {
            Mcm0010uBrandKoseiDto dto = new Mcm0010uBrandKoseiDto();
            dto.setPlantId(toBigDecimal(row.get("PLANT_ID")));
            dto.setBrandId(toBigDecimal(row.get("BRAND_ID")));
            dto.setNounyuKbn(toStr(row.get("NOUNYU_KBN")));
            dto.setBrandNk(toStr(row.get("BRAND_NK")));
            dto.setBrandsyosaiNk(toStr(row.get("BRANDSYOSAI_NK")));
            dto.setSuryo(toLong(row.get("SURYO")));
            result.add(dto);
        }
        return result;
    }

    /**
     * 機器構成マスタ検索
     * 元VB SQL (Mcm0010u_MCM_MA_KIKIKOSEITableAdapter.xml):
     *   SELECT SUM(MAF.SURYO_NM * MAE.SET_NM) as SURYO,
     *          MAJ.SEIZOMAKER_NK, MAF.KIKIHINMEI_NK, MAF.KIKIKATASHIKI, MAE.PLANT_ID
     *   FROM MCM_MA_KIKIKOSEI MAE
     *   INNER JOIN MCM_MA_KIKIMEISAI MAF ON MAE.KIKIKOSEI_ID = MAF.KIKIKOSEI_ID
     *   INNER JOIN MCM_MA_ATSUKAIKIKI MAH ON MAF.ATSUKAIKIKI_ID = MAH.ATSUKAIKIKI_ID
     *   INNER JOIN MCM_MA_SEIZOMAKER MAJ ON MAH.SEIZOMAKER_ID = MAJ.SEIZOMAKER_ID
     *   GROUP BY MAJ.SEIZOMAKER_NK, MAF.KIKIHINMEI_NK, MAF.KIKIKATASHIKI,
     *            MAE.PLANT_ID, MAJ.HYOJIJUN, MAH.HYOJIJUN
     *   ORDER BY MAH.HYOJIJUN ASC, MAJ.HYOJIJUN ASC
     */
    public List<Mcm0010uKikiKoseiDto> searchKikiKosei(Mcm0010uForm form) {
        StringBuilder sql = new StringBuilder();
        sql.append("SELECT SUM(MAF.SURYO_NM * MAE.SET_NM) AS SURYO, ");
        sql.append("MAJ.SEIZOMAKER_NK, MAF.KIKIHINMEI_NK, MAF.KIKIKATASHIKI, MAE.PLANT_ID ");
        sql.append("FROM MCM.MCM_MA_KIKIKOSEI MAE ");
        sql.append("INNER JOIN MCM.MCM_MA_KIKIMEISAI MAF ON MAE.KIKIKOSEI_ID = MAF.KIKIKOSEI_ID ");
        sql.append("INNER JOIN MCM.MCM_MA_ATSUKAIKIKI MAH ON MAF.ATSUKAIKIKI_ID = MAH.ATSUKAIKIKI_ID ");
        sql.append("INNER JOIN MCM.MCM_MA_SEIZOMAKER MAJ ON MAH.SEIZOMAKER_ID = MAJ.SEIZOMAKER_ID ");
        sql.append("WHERE 1=1 ");

        List<Object> params = new ArrayList<>();
        appendExistsConditions(sql, params, form, "MAE");

        sql.append("GROUP BY MAJ.SEIZOMAKER_NK, MAF.KIKIHINMEI_NK, MAF.KIKIKATASHIKI, ");
        sql.append("MAE.PLANT_ID, MAJ.HYOJIJUN, MAH.HYOJIJUN ");
        sql.append("ORDER BY MAH.HYOJIJUN ASC, MAJ.HYOJIJUN ASC");

        List<Map<String, Object>> rows = jdbcTemplate.queryForList(sql.toString(), params.toArray());
        List<Mcm0010uKikiKoseiDto> result = new ArrayList<>();
        for (Map<String, Object> row : rows) {
            Mcm0010uKikiKoseiDto dto = new Mcm0010uKikiKoseiDto();
            dto.setPlantId(toBigDecimal(row.get("PLANT_ID")));
            dto.setSeizomakerNk(toStr(row.get("SEIZOMAKER_NK")));
            dto.setKikihinmeiNk(toStr(row.get("KIKIHINMEI_NK")));
            dto.setKikikatashiki(toStr(row.get("KIKIKATASHIKI")));
            dto.setSuryo(toBigDecimal(row.get("SURYO")));
            result.add(dto);
        }
        return result;
    }

    /**
     * プラント行削除
     * 元VB: McmDBUtility.DeletePlant(plantId)
     * 関連テーブルの削除含む
     */
    // VB MCM_DEL_PAC.SP_ALL の対象限定削除は下記 deletePlants に実装。

    /**
     * 指定プラントIDにブランド構成データが存在するか
     * 元VB: Me.Mcm0010uDataSet.MCM_MA_BRAND_KOSEI.Select("PLANT_ID = " & plantId).Length
     */
    public boolean hasBrandKosei(BigDecimal plantId) {
        Integer count = jdbcTemplate.queryForObject(
            "SELECT COUNT(*) FROM MCM.MCM_MA_BRAND_KOSEI WHERE PLANT_ID = ?",
            Integer.class, plantId);
        return count != null && count > 0;
    }

    // ===== EXISTS条件共通（ブランド構成・機器構成用） =====
    private void appendExistsConditions(StringBuilder sql,List<Object> args,Mcm0010uForm form,String target) {
        sql.append(" AND EXISTS (SELECT 1 FROM MCM.MCM_MA_PLANT P JOIN MCM.MCM_MA_NONYUSAKI MAA ON MAA.NONYUSAKI_ID=P.NONYUSAKI_ID WHERE P.PLANT_ID=").append(target).append(".PLANT_ID ");
        appendNonyusaki(sql,args,form);
        if(!isBlank(form.getSupportId())){sql.append(" AND P.SUPPORT_ID LIKE ?");args.add(prefix(form.getSupportId()));}
        if(!isBlank(form.getPlantNk())){sql.append(" AND UPPER(P.PLANT_NK) LIKE UPPER(?)");args.add(contains(form.getPlantNk()));}
        sql.append(") ");
    }
    private void appendNonyusaki(StringBuilder sql,List<Object> args,Mcm0010uForm form) {
        if(!isBlank(form.getNonyusakiCd())) {sql.append(" AND MAA.NONYUSAKI_CD LIKE ?");args.add(prefix(form.getNonyusakiCd()));}
        if(!isBlank(form.getNonyusakiNk())) {
            sql.append(" AND CHARINDEX(UPPER(?),UPPER(CONCAT(COALESCE(MAA.NONYUSAKI_NK,''),' ',COALESCE(MAA.KYUNONYUSAKI_NK,''),' ',COALESCE(MAA.NONYUSAKIKOJO_NK,''),' ',COALESCE(MAA.NONYUSAKIKANA_KN,''),' ',COALESCE(MAA.NONYUSAKIEIMEI_EN,'')))) > 0");args.add(nvl(form.getNonyusakiNk()));
        }
        if(!isBlank(form.getJusyo())) {
            sql.append(" AND CHARINDEX(UPPER(?),UPPER(CONCAT(COALESCE(MAA.YUBIN_NO,''),' ',COALESCE(MAA.JUSYO1_NK,''),' ',COALESCE(MAA.JUSYO2_NK,''),' ',COALESCE(MAA.KUNI_NK,''),' ',COALESCE(MAA.TEL_NO,''),' ',COALESCE(MAA.FAX_NO,''),' ',COALESCE(MAA.BIKO,'')))) > 0");args.add(nvl(form.getJusyo()));
        }
    }
    private String prefix(String value){return nvl(value)+"%";}
    private String contains(String value){return "%"+nvl(value)+"%";}

    public boolean canView(String user,String screen){String a=permissions.getAuthority(user,screen);return "1".equals(a)||"2".equals(a);}
    public boolean canUpdate(String user,String screen){return "2".equals(permissions.getAuthority(user,screen));}
    public boolean canDelete(String user){return canUpdate(user,"MCM0010U")&&permissions.isMaintenance(user);}
    public NonyusakiEntity findNonyusaki(BigDecimal id){
        validId(id); var rows=jdbcTemplate.query("SELECT * FROM MCM.MCM_MA_NONYUSAKI WHERE NONYUSAKI_ID=?",new BeanPropertyRowMapper<>(NonyusakiEntity.class),id);
        if(rows.size()!=1)throw new IllegalStateException("納入先が見つかりません。検索し直してください。");return rows.get(0);
    }
    public PlantEntity findPlant(BigDecimal id){
        validId(id);var rows=jdbcTemplate.query("SELECT * FROM MCM.MCM_MA_PLANT WHERE PLANT_ID=?",new BeanPropertyRowMapper<>(PlantEntity.class),id);
        if(rows.size()!=1)throw new IllegalStateException("プラントが見つかりません。検索し直してください。");return rows.get(0);
    }
    public static void validId(BigDecimal id){if(id==null||id.signum()<=0||id.stripTrailingZeros().scale()>0||id.compareTo(BigDecimal.valueOf(Long.MAX_VALUE))>0)throw new IllegalStateException("対象を選択し直してください。");}
    public static String plantVersion(PlantEntity p){return Objects.toString(p.getLastupdateDt(),"")+"/"+Objects.toString(p.getNonyusakiId(),"")+"/"+Objects.toString(p.getSupportId(),"")+"/"+Objects.toString(p.getPlantNk(),"");}
    /** 画面表示済みのプラントをVB SP_ALLと同じ範囲で一括削除。途中失敗時はすべて戻す。 */
    @Transactional(isolation=Isolation.SERIALIZABLE)
    public void deletePlants(List<BigDecimal> ids,Map<String,String> versions,String user) {
        if(!canDelete(user))throw new IllegalStateException("削除権限がありません。");
        if(ids==null||ids.isEmpty())throw new IllegalStateException("削除するプラントを選択してください。");
        var unique=new LinkedHashSet<String>();
        for(var id:ids){validId(id);if(!unique.add(id.toBigIntegerExact().toString()))throw new IllegalStateException("対象を選択し直してください。");}
        for(var id:ids){
            var locked=jdbcTemplate.query("SELECT * FROM MCM.MCM_MA_PLANT WITH (UPDLOCK,HOLDLOCK) WHERE PLANT_ID=?",new BeanPropertyRowMapper<>(PlantEntity.class),id);
            if(locked.size()!=1||!Objects.equals(versions.get(id.toBigIntegerExact().toString()),plantVersion(locked.get(0))))throw new IllegalStateException("他のユーザがデータを変更した可能性があります。検索し直してください。");
        }
        // 契約全体を削除するVB仕様により、別プラントへ波及する契約は事前に止める。
        for(var id:ids)checkSharedContracts(id,unique);
        for(var id:ids){
            var quoteIds=jdbcTemplate.queryForList("SELECT UM_KIHON_MITSUMORI_ID FROM MCM.MCM_UM_KIHON_MITSUMORI WHERE PLANT_ID=? ORDER BY UM_KIHON_MITSUMORI_ID",BigDecimal.class,id);
            for(var quote:quoteIds) {
                cancelUnsentMail(quote);
                // SQL ServerではUM_TANKAがUM_MITSUMORIを参照するため、子を先に削除する。
                jdbcTemplate.update("DELETE FROM MCM.MCM_UM_TANKA WHERE UM_KIKIMEISAI_ID IN (SELECT D.UM_KIKIMEISAI_ID FROM MCM.MCM_UM_KIKIMEISAI D JOIN MCM.MCM_UM_KIKIKOSEI C ON C.UM_KIKIKOSEI_ID=D.UM_KIKIKOSEI_ID JOIN MCM.MCM_UM_KIHON_BRAND B ON B.UM_KIHON_BRAND_ID=C.UM_KIHON_BRAND_ID WHERE B.UM_KIHON_MITSUMORI_ID=?)",quote);
                userQuotes.deleteUser(quote,user);
            }
        }
        for(var id:ids){
            var quoteIds=jdbcTemplate.queryForList("SELECT TM_IRAI_ID FROM MCM.MCM_TM_MITSUMORI WHERE PLANT_ID=? ORDER BY TM_IRAI_ID",BigDecimal.class,id);
            for(var quote:quoteIds)deleteSupplier(quote);
        }
        for(var id:ids)deleteMaster(id);
    }
    /** 04のメール送信待ちキューと整合させる。送信済みの履歴は保持する。 */
    private void cancelUnsentMail(BigDecimal quote) {
        Boolean installed=jdbcTemplate.execute((org.springframework.jdbc.core.ConnectionCallback<Boolean>)c->{
            try(var tables=c.getMetaData().getTables(c.getCatalog(),"MCM","MCM_WEB_MAIL_OUTBOX",null)){return tables.next();}
        });
        if(!Boolean.TRUE.equals(installed))return;
        var states=jdbcTemplate.queryForList("SELECT STATUS FROM MCM.MCM_WEB_MAIL_OUTBOX WITH (UPDLOCK,HOLDLOCK) WHERE UM_KIHON_MITSUMORI_ID=?",String.class,quote);
        if(states.contains("SENDING"))throw new IllegalStateException("申請メールの送信処理中です。送信結果を確認してから削除してください。");
        jdbcTemplate.update("UPDATE MCM.MCM_WEB_MAIL_OUTBOX SET STATUS='SKIPPED',LAST_ERROR=N'見積削除により送信中止',UPDATED_DT=CURRENT_TIMESTAMP WHERE UM_KIHON_MITSUMORI_ID=? AND STATUS IN ('PENDING','FAILED')",quote);
    }
    private static final String TK_CONTRACTS="SELECT DISTINCT B.TK_KEIYAKU_ID FROM MCM.MCM_TK_KIKAN B JOIN MCM.MCM_TK_KIKIKOSEI C ON C.TK_KIKAN_ID=B.TK_KIKAN_ID JOIN MCM.MCM_TK_KIKIMEISAI D ON D.TK_KIKIKOSEI_ID=C.TK_KIKIKOSEI_ID JOIN MCM.MCM_TK_TANKA F ON F.TK_KIKIMEISAI_ID=D.TK_KIKIMEISAI_ID JOIN MCM.MCM_TM_TANKA TF ON TF.TM_TANKA_ID=F.TM_TANKA_ID JOIN MCM.MCM_TM_KIKIMEISAI TD ON TD.TM_KIKIMEISAI_ID=TF.TM_KIKIMEISAI_ID JOIN MCM.MCM_TM_KIKIKOSEI TC ON TC.TM_KIKIKOSEI_ID=TD.TM_KIKIKOSEI_ID WHERE TC.TM_IRAI_ID=?";
    private void checkSharedContracts(BigDecimal plant,java.util.Set<String> selected) {
        String uk="SELECT DISTINCT M.PLANT_ID FROM MCM.MCM_UM_MITSUMORI Q JOIN MCM.MCM_UM_KIHON_MITSUMORI M ON M.UM_KIHON_MITSUMORI_ID=Q.UM_KIHON_MITSUMORI_ID JOIN MCM.MCM_UK_MITSUMORI_BRAND L ON L.UM_MITSUMORI_ID=Q.UM_MITSUMORI_ID JOIN MCM.MCM_UK_BRAND B ON B.UK_BRAND_ID=L.UK_BRAND_ID JOIN MCM.MCM_UK_KIKAN K ON K.UK_KIKAN_ID=B.UK_KIKAN_ID WHERE K.UK_KEIYAKU_ID IN (SELECT K2.UK_KEIYAKU_ID FROM MCM.MCM_UK_KIKAN K2 JOIN MCM.MCM_UK_BRAND B2 ON B2.UK_KIKAN_ID=K2.UK_KIKAN_ID JOIN MCM.MCM_UK_MITSUMORI_BRAND L2 ON L2.UK_BRAND_ID=B2.UK_BRAND_ID JOIN MCM.MCM_UM_MITSUMORI Q2 ON Q2.UM_MITSUMORI_ID=L2.UM_MITSUMORI_ID JOIN MCM.MCM_UM_KIHON_MITSUMORI M2 ON M2.UM_KIHON_MITSUMORI_ID=Q2.UM_KIHON_MITSUMORI_ID WHERE M2.PLANT_ID=?)";
        requireSelected(jdbcTemplate.queryForList(uk,BigDecimal.class,plant),selected);
        // 見積へのリンクが欠けた過去期間も、契約のPLANT_IDで対象外への波及を防止する。
        String ukPeriods="SELECT DISTINCT P.PLANT_ID FROM MCM.MCM_UK_KIKAN P WHERE P.UK_KEIYAKU_ID IN (SELECT K.UK_KEIYAKU_ID FROM MCM.MCM_UK_KIKAN K JOIN MCM.MCM_UK_BRAND B ON B.UK_KIKAN_ID=K.UK_KIKAN_ID JOIN MCM.MCM_UK_MITSUMORI_BRAND L ON L.UK_BRAND_ID=B.UK_BRAND_ID JOIN MCM.MCM_UM_MITSUMORI Q ON Q.UM_MITSUMORI_ID=L.UM_MITSUMORI_ID JOIN MCM.MCM_UM_KIHON_MITSUMORI M ON M.UM_KIHON_MITSUMORI_ID=Q.UM_KIHON_MITSUMORI_ID WHERE M.PLANT_ID=?)";
        requireSelected(jdbcTemplate.queryForList(ukPeriods,BigDecimal.class,plant),selected);
        var quotes=jdbcTemplate.queryForList("SELECT TM_IRAI_ID FROM MCM.MCM_TM_MITSUMORI WHERE PLANT_ID=?",BigDecimal.class,plant);
        for(var quote:quotes)for(var contract:jdbcTemplate.queryForList(TK_CONTRACTS,BigDecimal.class,quote)) {
            String tk="SELECT DISTINCT M.PLANT_ID FROM MCM.MCM_TK_KIKAN B JOIN MCM.MCM_TK_KIKIKOSEI C ON C.TK_KIKAN_ID=B.TK_KIKAN_ID JOIN MCM.MCM_TK_KIKIMEISAI D ON D.TK_KIKIKOSEI_ID=C.TK_KIKIKOSEI_ID JOIN MCM.MCM_TK_TANKA F ON F.TK_KIKIMEISAI_ID=D.TK_KIKIMEISAI_ID JOIN MCM.MCM_TM_TANKA TF ON TF.TM_TANKA_ID=F.TM_TANKA_ID JOIN MCM.MCM_TM_KIKIMEISAI TD ON TD.TM_KIKIMEISAI_ID=TF.TM_KIKIMEISAI_ID JOIN MCM.MCM_TM_KIKIKOSEI TC ON TC.TM_KIKIKOSEI_ID=TD.TM_KIKIKOSEI_ID JOIN MCM.MCM_TM_MITSUMORI M ON M.TM_IRAI_ID=TC.TM_IRAI_ID WHERE B.TK_KEIYAKU_ID=?";
            requireSelected(jdbcTemplate.queryForList(tk,BigDecimal.class,contract),selected);
            requireSelected(jdbcTemplate.queryForList("SELECT DISTINCT PLANT_ID FROM MCM.MCM_TK_KIKAN WHERE TK_KEIYAKU_ID=?",BigDecimal.class,contract),selected);
        }
    }
    private void requireSelected(List<BigDecimal> plants,java.util.Set<String> selected){for(var p:plants)if(p==null||!selected.contains(p.toBigIntegerExact().toString()))throw new IllegalStateException("別のプラントと契約を共有しているため、削除できません。");}
    private void deleteSupplier(BigDecimal quote) {
        Integer used=jdbcTemplate.queryForObject("SELECT COUNT(*) FROM MCM.MCM_UM_TANKA U JOIN MCM.MCM_TM_TANKA F ON F.TM_TANKA_ID=U.TM_TANKA_ID JOIN MCM.MCM_TM_KIKIMEISAI D ON D.TM_KIKIMEISAI_ID=F.TM_KIKIMEISAI_ID JOIN MCM.MCM_TM_KIKIKOSEI C ON C.TM_KIKIKOSEI_ID=D.TM_KIKIKOSEI_ID WHERE C.TM_IRAI_ID=?",Integer.class,quote);
        if(used!=null&&used>0)throw new IllegalStateException("別のユーザ見積で使用されているため、削除できません。");
        for(var contract:jdbcTemplate.queryForList(TK_CONTRACTS,BigDecimal.class,quote))deleteSupplierContract(contract);
        for(String sql:DELETE_TM)jdbcTemplate.update(sql,quote);
    }
    private void deleteSupplierContract(BigDecimal contract){for(String sql:DELETE_TK)jdbcTemplate.update(sql,contract);}
    private void deleteMaster(BigDecimal plant){for(String sql:DELETE_MA)jdbcTemplate.update(sql,plant);}

    private static final String[] DELETE_TK = {
        "DELETE TKK FROM MCM.MCM_TK_SIHARAIMEISAI TKK WHERE EXISTS (SELECT * FROM MCM.MCM_TK_KIKAN TKB INNER JOIN MCM.MCM_TK_SIHARAI TKJ ON TKJ.TK_KIKAN_ID = TKB.TK_KIKAN_ID WHERE TKJ.TK_SIHARAI_ID = TKK.TK_SIHARAI_ID AND TKB.TK_KEIYAKU_ID = ?)",
        "DELETE TKJ FROM MCM.MCM_TK_SIHARAI TKJ WHERE EXISTS (SELECT * FROM MCM.MCM_TK_KIKAN TKB WHERE TKB.TK_KIKAN_ID = TKJ.TK_KIKAN_ID AND TKB.TK_KEIYAKU_ID = ?)",
        "DELETE TKI FROM MCM.MCM_TK_TENKENMEISAI TKI WHERE EXISTS (SELECT * FROM MCM.MCM_TK_KIKAN TKB INNER JOIN MCM.MCM_TK_KIKIKOSEI TKC ON TKC.TK_KIKAN_ID = TKB.TK_KIKAN_ID INNER JOIN MCM.MCM_TK_TENKEN TKH ON TKH.TK_KIKIKOSEI_ID = TKC.TK_KIKIKOSEI_ID WHERE TKH.TK_TENKEN_ID = TKI.TK_TENKEN_ID AND TKB.TK_KEIYAKU_ID = ?)",
        "DELETE TKH FROM MCM.MCM_TK_TENKEN TKH WHERE EXISTS (SELECT * FROM MCM.MCM_TK_KIKAN TKB INNER JOIN MCM.MCM_TK_KIKIKOSEI TKC ON TKC.TK_KIKAN_ID = TKB.TK_KIKAN_ID WHERE TKC.TK_KIKIKOSEI_ID = TKH.TK_KIKIKOSEI_ID AND TKB.TK_KEIYAKU_ID = ?)",
        "DELETE TKG FROM MCM.MCM_TK_TENPU TKG WHERE EXISTS (SELECT * FROM MCM.MCM_TK_KIKAN TKB WHERE TKB.TK_KIKAN_ID = TKG.TK_KIKAN_ID AND TKB.TK_KEIYAKU_ID = ?)",
        "DELETE TKF FROM MCM.MCM_TK_TANKA TKF WHERE EXISTS (SELECT * FROM MCM.MCM_TK_KIKAN TKB INNER JOIN MCM.MCM_TK_KIKIKOSEI TKC ON TKC.TK_KIKAN_ID = TKB.TK_KIKAN_ID INNER JOIN MCM.MCM_TK_KIKIMEISAI TKD ON TKD.TK_KIKIKOSEI_ID = TKC.TK_KIKIKOSEI_ID WHERE TKD.TK_KIKIMEISAI_ID = TKF.TK_KIKIMEISAI_ID AND TKB.TK_KEIYAKU_ID = ?)",
        "DELETE TKE FROM MCM.MCM_TK_KOTAIMEISAI TKE WHERE EXISTS (SELECT * FROM MCM.MCM_TK_KIKAN TKB INNER JOIN MCM.MCM_TK_KIKIKOSEI TKC ON TKC.TK_KIKAN_ID = TKB.TK_KIKAN_ID INNER JOIN MCM.MCM_TK_KIKIMEISAI TKD ON TKD.TK_KIKIKOSEI_ID = TKC.TK_KIKIKOSEI_ID WHERE TKD.TK_KIKIMEISAI_ID = TKE.TK_KIKIMEISAI_ID AND TKB.TK_KEIYAKU_ID = ?)",
        "DELETE TKD FROM MCM.MCM_TK_KIKIMEISAI TKD WHERE EXISTS (SELECT * FROM MCM.MCM_TK_KIKAN TKB INNER JOIN MCM.MCM_TK_KIKIKOSEI TKC ON TKC.TK_KIKAN_ID = TKB.TK_KIKAN_ID WHERE TKC.TK_KIKIKOSEI_ID = TKD.TK_KIKIKOSEI_ID AND TKB.TK_KEIYAKU_ID = ?)",
        "DELETE TKC FROM MCM.MCM_TK_KIKIKOSEI TKC WHERE EXISTS (SELECT * FROM MCM.MCM_TK_KIKAN TKB WHERE TKB.TK_KIKAN_ID = TKC.TK_KIKAN_ID AND TKB.TK_KEIYAKU_ID = ?)",
        "DELETE TKB FROM MCM.MCM_TK_KIKAN TKB WHERE TKB.TK_KEIYAKU_ID = ?",
        "DELETE TKA FROM MCM.MCM_TK_KEIYAKU TKA WHERE TKA.TK_KEIYAKU_ID = ?"
    };

    private static final String[] DELETE_TM = {
        "DELETE TMH FROM MCM.MCM_TM_TENKEN TMH WHERE EXISTS (SELECT * FROM MCM.MCM_TM_KIKIKOSEI TMC WHERE TMC.TM_KIKIKOSEI_ID = TMH.TM_KIKIKOSEI_ID AND TMC.TM_IRAI_ID = ?)",
        "DELETE TMF FROM MCM.MCM_TM_TANKA TMF WHERE EXISTS (SELECT * FROM MCM.MCM_TM_KIKIKOSEI TMC INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD ON TMD.TM_KIKIKOSEI_ID = TMC.TM_KIKIKOSEI_ID WHERE TMD.TM_KIKIMEISAI_ID = TMF.TM_KIKIMEISAI_ID AND TMC.TM_IRAI_ID = ?)",
        "DELETE TMG FROM MCM.MCM_TM_KIKAN TMG WHERE EXISTS (SELECT * FROM MCM.MCM_TM_KEIYAKUJIKAN TMB WHERE TMB.TM_KEIYAKUJIKAN_ID = TMG.TM_KEIYAKUJIKAN_ID AND TMB.TM_IRAI_ID = ?)",
        "DELETE TME FROM MCM.MCM_TM_KOTAIMEISAI TME WHERE EXISTS (SELECT * FROM MCM.MCM_TM_KIKIKOSEI TMC INNER JOIN MCM.MCM_TM_KIKIMEISAI TMD ON TMD.TM_KIKIKOSEI_ID = TMC.TM_KIKIKOSEI_ID WHERE TMD.TM_KIKIMEISAI_ID = TME.TM_KIKIMEISAI_ID AND TMC.TM_IRAI_ID = ?)",
        "DELETE TMD FROM MCM.MCM_TM_KIKIMEISAI TMD WHERE EXISTS (SELECT * FROM MCM.MCM_TM_KIKIKOSEI TMC WHERE TMC.TM_KIKIKOSEI_ID = TMD.TM_KIKIKOSEI_ID AND TMC.TM_IRAI_ID = ?)",
        "DELETE TMC FROM MCM.MCM_TM_KIKIKOSEI TMC WHERE TMC.TM_IRAI_ID = ?",
        "DELETE TMB FROM MCM.MCM_TM_KEIYAKUJIKAN TMB WHERE TMB.TM_IRAI_ID = ?",
        "DELETE TMA FROM MCM.MCM_TM_MITSUMORI TMA WHERE TMA.TM_IRAI_ID = ?"
    };

    private static final String[] DELETE_MA = {
        "DELETE MAG FROM MCM.MCM_MA_KIKIKOTAIKANRI MAG WHERE EXISTS (SELECT * FROM MCM.MCM_MA_KIKIKOSEI MAE WHERE MAE.KIKIKOSEI_ID = MAG.KIKIKOSEI_ID AND MAE.PLANT_ID = ?)",
        "DELETE MAF FROM MCM.MCM_MA_KIKIMEISAI MAF WHERE EXISTS (SELECT * FROM MCM.MCM_MA_KIKIKOSEI MAE WHERE MAE.KIKIKOSEI_ID = MAF.KIKIKOSEI_ID AND MAE.PLANT_ID = ?)",
        "DELETE MAE FROM MCM.MCM_MA_KIKIKOSEI MAE WHERE MAE.PLANT_ID = ?",
        "DELETE MAD FROM MCM.MCM_MA_BRAND_KOSEI MAD WHERE MAD.PLANT_ID = ?",
        "DELETE MAB FROM MCM.MCM_MA_PLANT MAB WHERE MAB.PLANT_ID = ?"
    };

    // ===== ユーティリティ =====
    private boolean isBlank(String s) {
        return s == null || s.trim().isEmpty();
    }

    private String nvl(String s) {
        return s == null ? "" : s.trim();
    }

    private String toStr(Object o) {
        return o == null ? null : o.toString();
    }

    private BigDecimal toBigDecimal(Object o) {
        if (o == null) return null;
        if (o instanceof BigDecimal) return (BigDecimal) o;
        return new BigDecimal(o.toString());
    }

    private Long toLong(Object o) {
        if (o == null) return null;
        if (o instanceof Long) return (Long) o;
        return new BigDecimal(o.toString()).longValueExact();
    }
}
