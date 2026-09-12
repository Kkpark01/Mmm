package com.daifuku.mcm.repository;

import com.daifuku.mcm.dto.Mcm2004uRowDto;
import com.daifuku.mcm.form.Mcm2004uForm;
import org.springframework.jdbc.core.namedparam.MapSqlParameterSource;
import org.springframework.jdbc.core.namedparam.NamedParameterJdbcTemplate;
import org.springframework.stereotype.Repository;

import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

/**
 * 【変換元】Mcm2004uScreen.vb / Mcm2004uDataSet.Designer.vb / Mcm2004u_MCM_2004_VTableAdapter.xml
 * MCM2004U（カスタマー見積・契約一覧）JdbcTemplate Repository
 *
 * <p>Oracle → SQL Server 変換済み
 *   （INSTR→CHARINDEX, TO_DATE→CAST, ||→+, NVL→ISNULL 等）</p>
 *
 * @author MCM Migration Tool
 */
@Repository
public class Mcm2004uJdbcRepository {

    private final NamedParameterJdbcTemplate jdbc;

    public Mcm2004uJdbcRepository(NamedParameterJdbcTemplate jdbc) {
        this.jdbc = jdbc;
    }

    /** VB McmLogin_AuthorityInfo.xml: function-specific authority, not the user's maximum. */
    public String findAuthority(String loginId,String functionId) {
        String sql="SELECT MAX(C.RIYOKENGEN_KBN) FROM MCM.MCM_MO_TANTO A " +
            "JOIN MCM.MCM_MO_TANTOKENGEN B ON B.TANTO_ID=A.TANTO_ID " +
            "JOIN MCM.MCM_MO_KENGENKOSEI C ON C.KENGENBUNRUI_ID=B.KENGENBUNRUI_ID AND C.RIYOKENGEN_KBN=B.RIYOKENGEN_KBN " +
            "JOIN MCM.MCM_MO_KENGENBUNRUI D ON D.KENGENBUNRUI_ID=C.KENGENBUNRUI_ID " +
            "JOIN MCM.MCM_MO_KINO E ON E.KINO_ID=C.KINO_ID WHERE A.LOGIN_ID=:loginId AND E.KINO_ID=:functionId";
        return jdbc.queryForObject(sql,new MapSqlParameterSource("loginId",loginId).addValue("functionId",functionId),String.class);
    }
    public boolean findMaintenance(String loginId) {
        var flags=jdbc.queryForList("SELECT MAINTENANCE_FLG FROM MCM.MCM_MO_TANTO WHERE LOGIN_ID=:loginId",new MapSqlParameterSource("loginId",loginId),BigDecimal.class);
        return flags.size()==1&&flags.get(0)!=null&&BigDecimal.ONE.compareTo(flags.get(0))==0;
    }

    // ================================================================
    // SELECT列定義（共通）
    // 【変換元】Mcm2004u_MCM_2004_VTableAdapter.xml <SELECT>
    // ================================================================
    private static final String SELECT_COLUMNS =
        "UVA.UM_KIHON_MITSUMORI_ID, " +
        "UVA.JOTAI, " +
        "UVA.UM_MITSUMORI_NO, " +
        "UVA.KEIYAKU_NO, " +
        "UVA.UK_KEIYAKU_ID, " +
        "UVA.FUSEI, " +
        "UVA.NONYUSAKI_ID, " +
        "UVA.NONYUSAKI_CD, " +
        "UVA.NONYUSAKI_NK, " +
        "UVA.KYUNONYUSAKI_NK, " +
        "UVA.NONYUSAKIKOJO_NK, " +
        "UVA.NONYUSAKIKANA_KN, " +
        "UVA.NONYUSAKIEIMEI_EN, " +
        "UVA.PLANT_ID, " +
        "UVA.SUPPORT_ID, " +
        "UVA.PLANT_NK, " +
        "UVA.KEIYAKUJIKANTAI, " +
        "UVA.HOSYUHOHO, " +
        "UVA.MITSUMORI_DT, " +
        "UVA.MITSUMORISAKUSEISYA_NK, " +
        "UVA.SOFUTENPO_ID, " +
        "UVA.SOFUMEISHO1_NK, " +
        "UVA.SOFUMEISHO2_NK, " +
        "UVA.SOFUMEISHO3_NK, " +
        "UVA.SOFUMEISHO4_NK, " +
        "UVA.SOFUTENPORYAKU_NK, " +
        "UVA.SOFUTANTO_NK, " +
        "UVA.IRAITENPO_ID, " +
        "UVA.IRAIMEISHO1_NK, " +
        "UVA.IRAIMEISHO2_NK, " +
        "UVA.IRAIMEISHO3_NK, " +
        "UVA.IRAIMEISHO4_NK, " +
        "UVA.IRAITENPORYAKU_NK, " +
        "UVA.IRAITANTO_NK, " +
        "UVA.KEIYAKU_DT, " +
        "UVA.SYOUNIN_JOTAI, " +
        "UVA.MITSUMORI_UPD, " +
        "UVA.MITSUMORI_COPY, " +
        "UVA.MITSUMORI_DEL, " +
        "UVA.KEIYAKU_KEIYAKU, " +
        "UVA.KEIYAKU_DEL";

    // ================================================================
    // RowMapper
    // ================================================================
    private Mcm2004uRowDto mapRow(ResultSet rs, int rowNum) throws SQLException {
        Mcm2004uRowDto dto = new Mcm2004uRowDto();
        dto.setUmKihonMitsumoriId(rs.getBigDecimal("UM_KIHON_MITSUMORI_ID"));
        dto.setJotai(rs.getString("JOTAI"));
        dto.setUmMitsumoriNo(rs.getString("UM_MITSUMORI_NO"));
        dto.setKeiyakuNo(rs.getString("KEIYAKU_NO"));
        dto.setUkKeiyakuId(rs.getBigDecimal("UK_KEIYAKU_ID"));
        dto.setFusei(rs.getString("FUSEI"));
        dto.setNonyusakiId(rs.getBigDecimal("NONYUSAKI_ID"));
        dto.setNonyusakiCd(rs.getString("NONYUSAKI_CD"));
        dto.setNonyusakiNk(rs.getString("NONYUSAKI_NK"));
        dto.setKyunonyusakiNk(rs.getString("KYUNONYUSAKI_NK"));
        dto.setNonyusakikojoNk(rs.getString("NONYUSAKIKOJO_NK"));
        dto.setNonyusakikanaKn(rs.getString("NONYUSAKIKANA_KN"));
        dto.setNonyusakieimeiEn(rs.getString("NONYUSAKIEIMEI_EN"));
        dto.setPlantId(rs.getBigDecimal("PLANT_ID"));
        dto.setSupportId(rs.getString("SUPPORT_ID"));
        dto.setPlantNk(rs.getString("PLANT_NK"));
        dto.setKeiyakujikantai(rs.getString("KEIYAKUJIKANTAI"));
        dto.setHosyuhoho(rs.getString("HOSYUHOHO"));
        java.sql.Date mDt = rs.getDate("MITSUMORI_DT");
        dto.setMitsumoriDt(mDt != null ? mDt.toLocalDate() : null);
        dto.setMitsumorisakuseisyaNk(rs.getString("MITSUMORISAKUSEISYA_NK"));
        dto.setSofutenpoId(rs.getBigDecimal("SOFUTENPO_ID"));
        dto.setSofumeisho1Nk(rs.getString("SOFUMEISHO1_NK"));
        dto.setSofumeisho2Nk(rs.getString("SOFUMEISHO2_NK"));
        dto.setSofumeisho3Nk(rs.getString("SOFUMEISHO3_NK"));
        dto.setSofumeisho4Nk(rs.getString("SOFUMEISHO4_NK"));
        dto.setSofutenporyakuNk(rs.getString("SOFUTENPORYAKU_NK"));
        dto.setSofutantoNk(rs.getString("SOFUTANTO_NK"));
        dto.setIraitenpoId(rs.getBigDecimal("IRAITENPO_ID"));
        dto.setIraimeisho1Nk(rs.getString("IRAIMEISHO1_NK"));
        dto.setIraimeisho2Nk(rs.getString("IRAIMEISHO2_NK"));
        dto.setIraimeisho3Nk(rs.getString("IRAIMEISHO3_NK"));
        dto.setIraimeisho4Nk(rs.getString("IRAIMEISHO4_NK"));
        dto.setIraitenporyakuNk(rs.getString("IRAITENPORYAKU_NK"));
        dto.setIraitantoNk(rs.getString("IRAITANTO_NK"));
        java.sql.Date kDt = rs.getDate("KEIYAKU_DT");
        dto.setKeiyakuDt(kDt != null ? kDt.toLocalDate() : null);
        dto.setSyouninJotai(rs.getString("SYOUNIN_JOTAI"));
        dto.setMitsumoriUpd(rs.getString("MITSUMORI_UPD"));
        dto.setMitsumoriCopy(rs.getString("MITSUMORI_COPY"));
        dto.setMitsumoriDel(rs.getString("MITSUMORI_DEL"));
        dto.setKeiyakuKeiyaku(rs.getString("KEIYAKU_KEIYAKU"));
        dto.setKeiyakuDel(rs.getString("KEIYAKU_DEL"));
        return dto;
    }

    // ================================================================
    // 1. カスタマー基本見積ID指定検索
    // 【変換元】Mcm2004uScreen.vb - Search(UmKihonMitsumoriId)
    //   元コード: Me.MCM_2004_VTableAdapter.Fill(DataSet, umKihonMitsumoriId)
    // ================================================================
    public List<Mcm2004uRowDto> searchByUmKihonMitsumoriId(BigDecimal umKihonMitsumoriId) {

        String sql =
            "SELECT " + SELECT_COLUMNS + " " +
            "FROM MCM.MCM_2004_V UVA " +
            "WHERE UVA.UM_KIHON_MITSUMORI_ID = :umKihonMitsumoriId " +
            "ORDER BY UVA.UM_MITSUMORI_NO ASC, UVA.KEIYAKU_NO ASC";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("umKihonMitsumoriId", umKihonMitsumoriId);
        return jdbc.query(sql, params, this::mapRow);
    }

    // ================================================================
    // 2. 動的条件検索
    // 【変換元】Mcm2004u_MCM_2004_VTableAdapter.xml（11条件動的WHERE）
    //   Oracle → SQL Server 変換:
    //     TO_DATE(:{2},'YYYY/MM/DD') → CAST(:param AS DATE)
    //     INSTR(a, b) > 0           → CHARINDEX(b, a) > 0
    //     || (文字列結合)            → +
    //     UPPER()                    → UPPER() (同一)
    //     LIKE UPPER(:{0}) SUFFIX="%"→ LIKE UPPER(:param) + '%'
    // ================================================================
    public List<Mcm2004uRowDto> searchByConditions(Mcm2004uForm form) {

        StringBuilder sql = new StringBuilder();
        sql.append("SELECT ").append(SELECT_COLUMNS).append(" ");
        sql.append("FROM MCM.MCM_2004_V UVA ");
        sql.append("WHERE 1 = 1 ");

        MapSqlParameterSource params = new MapSqlParameterSource();

        // --- 条件1: 見積NO（前方一致） ---
        // 【変換元】<CONDITION SUFFIX="%">UPPER(UVA.UM_MITSUMORI_NO) LIKE UPPER(:{0})</CONDITION>
        if (isNotEmpty(form.getUmMitsumoriNo())) {
            sql.append("AND UPPER(UVA.UM_MITSUMORI_NO) LIKE UPPER(:umMitsumoriNo) + '%' ");
            params.addValue("umMitsumoriNo", form.getUmMitsumoriNo());
        }

        // --- 条件2: 契約NO（前方一致） ---
        // 【変換元】<CONDITION SUFFIX="%">UPPER(UVA.KEIYAKU_NO) LIKE UPPER(:{1})</CONDITION>
        if (isNotEmpty(form.getKeiyakuNo())) {
            sql.append("AND UPPER(UVA.KEIYAKU_NO) LIKE UPPER(:keiyakuNo) + '%' ");
            params.addValue("keiyakuNo", form.getKeiyakuNo());
        }

        // --- 条件3: 見積作成日From ---
        // 【変換元】<CONDITION>UVA.MITSUMORI_DT >= TO_DATE(:{2},'YYYY/MM/DD')</CONDITION>
        // → SQL Server: CAST(:mitsumoriDtStart AS DATE)
        if (isNotEmpty(form.getMitsumoriDtStart())) {
            sql.append("AND UVA.MITSUMORI_DT >= CAST(:mitsumoriDtStart AS DATE) ");
            params.addValue("mitsumoriDtStart", form.getMitsumoriDtStart());
        }

        // --- 条件4: 見積作成日To ---
        // 【変換元】<CONDITION>UVA.MITSUMORI_DT <= TO_DATE(:{3},'YYYY/MM/DD')</CONDITION>
        if (isNotEmpty(form.getMitsumoriDtEnd())) {
            sql.append("AND UVA.MITSUMORI_DT <= CAST(:mitsumoriDtEnd AS DATE) ");
            params.addValue("mitsumoriDtEnd", form.getMitsumoriDtEnd());
        }

        // --- 条件5: 送付先事業所 ---
        // 【変換元】<CONDITION>UVA.SOFUTENPO_ID = :{4}</CONDITION>
        if (isNotEmpty(form.getSofutenpoId())) {
            sql.append("AND UVA.SOFUTENPO_ID = :sofutenpoId ");
            params.addValue("sofutenpoId", new BigDecimal(form.getSofutenpoId()));
        }

        // --- 条件6: 納入先コード（部分一致） ---
        // 【変換元】<CONDITION>INSTR(UVA.NONYUSAKI_CD, :{5}) > 0</CONDITION>
        // → SQL Server: CHARINDEX(:param, col) > 0
        if (isNotEmpty(form.getNonyusakiCd())) {
            sql.append("AND CHARINDEX(:nonyusakiCd, UVA.NONYUSAKI_CD) > 0 ");
            params.addValue("nonyusakiCd", form.getNonyusakiCd());
        }

        // --- 条件7: 納入先名（複合部分一致） ---
        // 【変換元】<CONDITION>INSTR(UPPER(UVA.NONYUSAKI_NK || ' ' || UVA.KYUNONYUSAKI_NK
        //           || ' ' || UVA.NONYUSAKIKANA_KN || '' || UVA.NONYUSAKIEIMEI_EN), UPPER(:{6})) > 0
        // → SQL Server: CHARINDEX(UPPER(:param),
        //     UPPER(UVA.NONYUSAKI_NK + ' ' + UVA.KYUNONYUSAKI_NK + ' ' + ...)) > 0
        if (isNotEmpty(form.getNonyusakiNk())) {
            sql.append("AND CHARINDEX(UPPER(:nonyusakiNk), ");
            sql.append("UPPER(ISNULL(UVA.NONYUSAKI_NK,'') + ' ' + ");
            sql.append("ISNULL(UVA.KYUNONYUSAKI_NK,'') + ' ' + ");
            sql.append("ISNULL(UVA.NONYUSAKIKANA_KN,'') + ");
            sql.append("ISNULL(UVA.NONYUSAKIEIMEI_EN,''))) > 0 ");
            params.addValue("nonyusakiNk", form.getNonyusakiNk());
        }

        // --- 条件8: サポートID（部分一致） ---
        // 【変換元】<CONDITION>INSTR(UVA.SUPPORT_ID, :{7}) > 0</CONDITION>
        if (isNotEmpty(form.getSupportId())) {
            sql.append("AND CHARINDEX(:supportId, UVA.SUPPORT_ID) > 0 ");
            params.addValue("supportId", form.getSupportId());
        }

        // --- 条件9: プラント名（部分一致） ---
        // 【変換元】<CONDITION>INSTR(UVA.PLANT_NK, :{8}) > 0</CONDITION>
        if (isNotEmpty(form.getPlantNk())) {
            sql.append("AND CHARINDEX(:plantNk, UVA.PLANT_NK) > 0 ");
            params.addValue("plantNk", form.getPlantNk());
        }

        // --- 条件10: 状態（複数値IN相当） ---
        // 【変換元】<CONDITION>INSTR(:{9}, UVA.JOTAI) > 0</CONDITION>
        // → SQL Server: CHARINDEX(UVA.JOTAI, :jotaiList) > 0
        String jotaiList = buildJotaiList(form);
        if (isNotEmpty(jotaiList)) {
            sql.append("AND CHARINDEX(UVA.JOTAI, :jotaiList) > 0 ");
            params.addValue("jotaiList", jotaiList);
        }

        // --- 条件11: 承認状態（複数値IN相当） ---
        // 【変換元】<CONDITION>INSTR(:{10}, UVA.SYOUNIN_JOTAI) > 0</CONDITION>
        String syouninJotaiList = buildSyouninJotaiList(form);
        if (isNotEmpty(syouninJotaiList)) {
            sql.append("AND CHARINDEX(UVA.SYOUNIN_JOTAI, :syouninJotaiList) > 0 ");
            params.addValue("syouninJotaiList", syouninJotaiList);
        }

        // --- ORDER BY ---
        sql.append("ORDER BY UVA.UM_MITSUMORI_NO ASC, UVA.KEIYAKU_NO ASC");

        return jdbc.query(sql.toString(), params, this::mapRow);
    }

    // ================================================================
    // 3. 見積破棄（JOTAI更新）
    // 【変換元】Mcm2004uScreen.vb - CellContentClick() 見積破棄分岐
    //   元コード:
    //     Mcm2004uDataSet.MCM_UM_KIHON_MITSUMORI(0).JOTAI = McmConstant.JOTAI_HAKI_CD
    //     Me.DaoContainer.Update(Mcm2004uDataSet, "MCM_UM_KIHON_MITSUMORI")
    // ================================================================
    public int updateMitsumoriJotai(BigDecimal umKihonMitsumoriId,
                                     String jotai,
                                     String loginUserId) {

        String sql =
            "UPDATE MCM.MCM_UM_KIHON_MITSUMORI " +
            "SET JOTAI = :jotai, " +
            "    LASTUPDATE_BY = :loginUserId, " +
            "    LASTUPDATE_DT = GETDATE() " +
            "WHERE UM_KIHON_MITSUMORI_ID = :umKihonMitsumoriId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("jotai", jotai);
        params.addValue("loginUserId", loginUserId);
        params.addValue("umKihonMitsumoriId", umKihonMitsumoriId);
        return jdbc.update(sql, params);
    }

    // ================================================================
    // 4. 契約破棄（JOTAI更新）— MCM_UK_KEIYAKU テーブル
    // 【変換元】Mcm2004uScreen.vb - CellContentClick() 契約破棄分岐
    //   元コード:
    //     Mcm2004uDataSet.MCM_UK_KEIYAKU(0).JOTAI = McmConstant.JOTAI_HAKI_CD
    //     Me.DaoContainer.Update(Mcm2004uDataSet, "MCM_UK_KEIYAKU")
    // ================================================================
    public int updateUkKeiyakuJotai(BigDecimal ukKeiyakuId,
                                     String jotai,
                                     String loginUserId) {

        String sql =
            "UPDATE MCM.MCM_UK_KEIYAKU " +
            "SET JOTAI = :jotai, " +
            "    LASTUPDATE_BY = :loginUserId, " +
            "    LASTUPDATE_DT = GETDATE() " +
            "WHERE UK_KEIYAKU_ID = :ukKeiyakuId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("jotai", jotai);
        params.addValue("loginUserId", loginUserId);
        params.addValue("ukKeiyakuId", ukKeiyakuId);
        return jdbc.update(sql, params);
    }

    // ================================================================
    // 5. 契約破棄（JOTAI更新）— MCM_UM_KIHON_MITSUMORI テーブル
    //    契約破棄時に見積側のJOTAIも更新するケース
    // 【変換元】CellContentClick() 契約破棄 → 基本見積のJOTAIも戻す
    // ================================================================
    public int updateMitsumoriJotaiByUkKeiyaku(BigDecimal umKihonMitsumoriId,
                                                String jotai,
                                                String loginUserId) {
        return updateMitsumoriJotai(umKihonMitsumoriId, jotai, loginUserId);
    }

    // ================================================================
    // 6. ユーザ契約存在チェック
    // 【変換元】CellContentClick() 内の存在チェック
    // ================================================================
    public int countUkKeiyaku(BigDecimal ukKeiyakuId) {

        String sql =
            "SELECT COUNT(*) FROM MCM.MCM_UK_KEIYAKU " +
            "WHERE UK_KEIYAKU_ID = :ukKeiyakuId";

        MapSqlParameterSource params = new MapSqlParameterSource();
        params.addValue("ukKeiyakuId", ukKeiyakuId);
        Integer count = jdbc.queryForObject(sql, params, Integer.class);
        return count != null ? count : 0;
    }

    // ================================================================
    // 7. 行削除（McmDBUtility.DeleteUser 相当）
    // 【変換元】Mcm2004uScreen.vb - RowDeleteButton_Click()
    //   元コード: McmDBUtility.DeleteUser(umKihonMitsumoriId)
    // TODO: McmDBUtility.DeleteUser の完全な実装に合わせて調整すること
    // ================================================================
    /** McmDBUtility.DeleteUser -> MCM_DEL_PAC.SP_UM, all 18 statements in VB order. */
    public int deleteUser(BigDecimal id, String loginUserId) {
        var params = new MapSqlParameterSource("id",id);
        var contracts = jdbc.queryForList("SELECT DISTINCT UKB.UK_KEIYAKU_ID FROM MCM.MCM_UK_KIKAN UKB INNER JOIN MCM.MCM_UK_BRAND UKC ON UKC.UK_KIKAN_ID = UKB.UK_KIKAN_ID INNER JOIN MCM.MCM_UK_MITSUMORI_BRAND UKG ON UKG.UK_BRAND_ID = UKC.UK_BRAND_ID INNER JOIN MCM.MCM_UM_MITSUMORI UMG ON UMG.UM_MITSUMORI_ID = UKG.UM_MITSUMORI_ID WHERE UMG.UM_KIHON_MITSUMORI_ID = :id",params,BigDecimal.class);
        for (var contract:contracts) {
            params.addValue("contract",contract);
            jdbc.update("DELETE UKI FROM MCM.MCM_UK_SEIBAN UKI WHERE UKI.UK_KEIYAKU_ID = :contract",params);
            jdbc.update("DELETE UKH FROM MCM.MCM_UK_TENKEN UKH WHERE EXISTS (SELECT * FROM MCM.MCM_UK_KIKAN UKB WHERE UKB.UK_KIKAN_ID = UKH.UK_KIKAN_ID AND UKB.UK_KEIYAKU_ID = :contract)",params);
            jdbc.update("DELETE UKG FROM MCM.MCM_UK_MITSUMORI_BRAND UKG WHERE EXISTS (SELECT * FROM MCM.MCM_UK_KIKAN UKB INNER JOIN MCM.MCM_UK_BRAND UKC ON UKC.UK_KIKAN_ID = UKB.UK_KIKAN_ID WHERE UKC.UK_BRAND_ID = UKG.UK_BRAND_ID AND UKB.UK_KEIYAKU_ID = :contract)",params);
            jdbc.update("DELETE UKF FROM MCM.MCM_UK_KOTAIMEISAI UKF WHERE EXISTS (SELECT * FROM MCM.MCM_UK_KIKAN UKB INNER JOIN MCM.MCM_UK_BRAND UKC ON UKC.UK_KIKAN_ID = UKB.UK_KIKAN_ID INNER JOIN MCM.MCM_UK_KIKIKOSEI UKD ON UKD.UK_BRAND_ID = UKC.UK_BRAND_ID INNER JOIN MCM.MCM_UK_KIKIMEISAI UKE ON UKE.UK_KIKIKOSEI_ID = UKD.UK_KIKIKOSEI_ID WHERE UKE.UK_KIKIMEISAI_ID = UKF.UK_KIKIMEISAI_ID AND UKB.UK_KEIYAKU_ID = :contract)",params);
            jdbc.update("DELETE UKE FROM MCM.MCM_UK_KIKIMEISAI UKE WHERE EXISTS (SELECT * FROM MCM.MCM_UK_KIKAN UKB INNER JOIN MCM.MCM_UK_BRAND UKC ON UKC.UK_KIKAN_ID = UKB.UK_KIKAN_ID INNER JOIN MCM.MCM_UK_KIKIKOSEI UKD ON UKD.UK_BRAND_ID = UKC.UK_BRAND_ID WHERE UKD.UK_KIKIKOSEI_ID = UKE.UK_KIKIKOSEI_ID AND UKB.UK_KEIYAKU_ID = :contract)",params);
            jdbc.update("DELETE UKD FROM MCM.MCM_UK_KIKIKOSEI UKD WHERE EXISTS (SELECT * FROM MCM.MCM_UK_KIKAN UKB INNER JOIN MCM.MCM_UK_BRAND UKC ON UKC.UK_KIKAN_ID = UKB.UK_KIKAN_ID WHERE UKC.UK_BRAND_ID = UKD.UK_BRAND_ID AND UKB.UK_KEIYAKU_ID = :contract)",params);
            jdbc.update("DELETE UKC FROM MCM.MCM_UK_BRAND UKC WHERE EXISTS (SELECT * FROM MCM.MCM_UK_KIKAN UKB WHERE UKB.UK_KIKAN_ID = UKC.UK_KIKAN_ID AND UKB.UK_KEIYAKU_ID = :contract)",params);
            jdbc.update("DELETE UKB FROM MCM.MCM_UK_KIKAN UKB WHERE UKB.UK_KEIYAKU_ID = :contract",params);
            jdbc.update("DELETE UKA FROM MCM.MCM_UK_KEIYAKU UKA WHERE UKA.UK_KEIYAKU_ID = :contract",params);
        }
        jdbc.update("DELETE UMI FROM MCM.MCM_UM_TENPU UMI WHERE UMI.UM_KIHON_MITSUMORI_ID = :id",params);
        jdbc.update("DELETE UMH FROM MCM.MCM_UM_BRAND UMH WHERE EXISTS (SELECT * FROM MCM.MCM_UM_KIHON_BRAND UMB WHERE UMB.UM_KIHON_BRAND_ID = UMH.UM_KIHON_BRAND_ID AND UMB.UM_KIHON_MITSUMORI_ID = :id)",params);
        jdbc.update("DELETE UMG FROM MCM.MCM_UM_MITSUMORI UMG WHERE UMG.UM_KIHON_MITSUMORI_ID = :id",params);
        jdbc.update("DELETE UMF FROM MCM.MCM_UM_TANKA UMF WHERE EXISTS (SELECT * FROM MCM.MCM_UM_KIHON_BRAND UMB INNER JOIN MCM.MCM_UM_KIKIKOSEI UMC ON UMC.UM_KIHON_BRAND_ID = UMB.UM_KIHON_BRAND_ID INNER JOIN MCM.MCM_UM_KIKIMEISAI UMD ON UMD.UM_KIKIKOSEI_ID = UMC.UM_KIKIKOSEI_ID WHERE UMD.UM_KIKIMEISAI_ID = UMF.UM_KIKIMEISAI_ID AND UMB.UM_KIHON_MITSUMORI_ID = :id)",params);
        jdbc.update("DELETE UME FROM MCM.MCM_UM_KOTAIMEISAI UME WHERE EXISTS (SELECT * FROM MCM.MCM_UM_KIHON_BRAND UMB INNER JOIN MCM.MCM_UM_KIKIKOSEI UMC ON UMC.UM_KIHON_BRAND_ID = UMB.UM_KIHON_BRAND_ID INNER JOIN MCM.MCM_UM_KIKIMEISAI UMD ON UMD.UM_KIKIKOSEI_ID = UMC.UM_KIKIKOSEI_ID WHERE UMD.UM_KIKIMEISAI_ID = UME.UM_KIKIMEISAI_ID AND UMB.UM_KIHON_MITSUMORI_ID = :id)",params);
        jdbc.update("DELETE UMD FROM MCM.MCM_UM_KIKIMEISAI UMD WHERE EXISTS (SELECT * FROM MCM.MCM_UM_KIHON_BRAND UMB INNER JOIN MCM.MCM_UM_KIKIKOSEI UMC ON UMC.UM_KIHON_BRAND_ID = UMB.UM_KIHON_BRAND_ID WHERE UMC.UM_KIKIKOSEI_ID = UMD.UM_KIKIKOSEI_ID AND UMB.UM_KIHON_MITSUMORI_ID = :id)",params);
        jdbc.update("DELETE UMC FROM MCM.MCM_UM_KIKIKOSEI UMC WHERE EXISTS (SELECT * FROM MCM.MCM_UM_KIHON_BRAND UMB WHERE UMB.UM_KIHON_BRAND_ID = UMC.UM_KIHON_BRAND_ID AND UMB.UM_KIHON_MITSUMORI_ID = :id)",params);
        jdbc.update("DELETE UMB FROM MCM.MCM_UM_KIHON_BRAND UMB WHERE UMB.UM_KIHON_MITSUMORI_ID = :id",params);
        return jdbc.update("DELETE UMA FROM MCM.MCM_UM_KIHON_MITSUMORI UMA WHERE UMA.UM_KIHON_MITSUMORI_ID = :id",params);
    }

    // ================================================================
    // ヘルパーメソッド
    // ================================================================

    /**
     * 状態チェックボックスからJOTAI値リスト文字列を構築する。
     * 【変換元】SEARCHButton_Click() 内の状態チェックボックス値構築
     *   元コード:
     *     If JOTAI_MITSUMORICheckBox.Checked Then jotaiCheck = McmConstant.JOTAI_MITSUMORI_NK + " "
     *     If JOTAI_KEIYAKUCheckBox.Checked Then jotaiCheck = jotaiCheck + McmConstant.JOTAI_KEIYAKU_NK + " "
     *     ...
     *   MCM_2004_V ビューの JOTAI 列は漢字名称（"見積", "契約" 等）で格納されており、
     *   CHARINDEX(UVA.JOTAI, :jotaiList) > 0 で検索するため、漢字＋スペース区切りが必要。
     */
    private String buildJotaiList(Mcm2004uForm form) {
        StringBuilder sb = new StringBuilder();
        if (form.isJotaiMitsumori())  sb.append("見積 ");  // McmConstant.JOTAI_MITSUMORI_NK + " "
        if (form.isJotaiKeiyaku())    sb.append("契約 ");  // McmConstant.JOTAI_KEIYAKU_NK + " "
        if (form.isJotaiKaiyaku())    sb.append("解約 ");  // McmConstant.JOTAI_KAIYAKU_NK + " "
        if (form.isJotaiHaki())       sb.append("破棄 ");  // McmConstant.JOTAI_HAKI_NK + " "
        return sb.toString();
    }

    /**
     * 承認状態チェックボックスからSYOUNIN_JOTAI値リスト文字列を構築する。
     * 【変換元】SEARCHButton_Click() 内の承認状態チェックボックス値構築
     *   元コード:
     *     If SYOUNIN_JOTAI_SAKUSEICHUCheckBox.Checked Then syouninCheck = McmConstant.SHONINJOTAI_SAKUSEICHU_NK + " "
     *     ...
     */
    private String buildSyouninJotaiList(Mcm2004uForm form) {
        StringBuilder sb = new StringBuilder();
        if (form.isSyouninJotaiSakuseichu())     sb.append("作成中 ");  // SHONINJOTAI_SAKUSEICHU_NK + " "
        if (form.isSyouninJotaiSinsachu())       sb.append("審査中 ");  // SHONINJOTAI_SHINSACHU_NK + " "
        if (form.isSyouninJotaiShoninchu())      sb.append("承認中 ");  // SHONINJOTAI_SHONINCHU_NK + " "
        if (form.isSyouninJotaiShoninzumi())     sb.append("承認済 ");  // SHONINJOTAI_SHONINZUMI_NK + " "
        if (form.isSyouninJotaiSashimodoshichu()) sb.append("差戻中 "); // SHONINJOTAI_SASHIMODOSHI_NK + " "
        return sb.toString();
    }

    private boolean isNotEmpty(String s) {
        return s != null && !s.isBlank();
    }

    // ================================================================
    // 8. 送付先事業所一覧取得（コンボボックス用）
    // 【変換元】McmSohumeisyo4NkDataTable_Sql.xml
    //   元SQL:
    //     SELECT DISTINCT UMA.SOFUTENPO_ID, MAR.MEISHO4_NK, MAR.HYOJIJUN
    //     FROM MCM_UM_KIHON_MITSUMORI UMA
    //     LEFT OUTER JOIN MCM_MA_TENPO MAR ON MAR.TENPO_ID = UMA.SOFUTENPO_ID
    //     WHERE UMA.SOFUMEISHO4_NK IS NOT NULL
    //     ORDER BY MAR.HYOJIJUN
    // ================================================================
    public List<java.util.Map<String, Object>> findDistinctSofutenpo() {
        String sql =
            "SELECT DISTINCT UMA.SOFUTENPO_ID, MAR.MEISHO4_NK, MAR.HYOJIJUN " +
            "FROM MCM.MCM_UM_KIHON_MITSUMORI UMA " +
            "LEFT OUTER JOIN MCM.MCM_MA_TENPO MAR ON MAR.TENPO_ID = UMA.SOFUTENPO_ID " +
            "WHERE UMA.SOFUMEISHO4_NK IS NOT NULL " +
            "ORDER BY MAR.HYOJIJUN";
        return jdbc.queryForList(sql, new MapSqlParameterSource());
    }
}
