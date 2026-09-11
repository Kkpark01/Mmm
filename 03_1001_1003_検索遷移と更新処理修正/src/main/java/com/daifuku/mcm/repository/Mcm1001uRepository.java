/**
 * 【変換元】Mcm1001uScreen.vb - SearchNonyusaki() / SearchPlant()
 *           Mcm1001uDataSet.Designer.vb - SelectCommand (XML内SQL)
 *   元ファイル行数: 約4,296行（Screen.vb全体）
 *   MCM1001U（取引先見積依頼作成検索）JdbcTemplate専用リポジトリ
 *
 *   新パターン: MCM0021Uと同様に、動的WHERE + EXISTS副問合せのため
 *   JdbcTemplateで動的SQL組み立てを行う専用リポジトリ。
 *
 * @author MCM Migration Tool
 * @since v8
 */
package com.daifuku.mcm.repository;

import com.daifuku.mcm.dto.Mcm1001uNonyusakiDto;
import com.daifuku.mcm.dto.Mcm1001uPlantDto;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.RowMapper;
import org.springframework.stereotype.Repository;

import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

/**
 * MCM1001U 取引先見積依頼作成検索 リポジトリ.
 *
 * 【変換元】Mcm1001uScreen.vb + Mcm1001uDataSet.Designer.vb
 *   - Oracle SQL → SQL Server 変換済み
 *   - INSTR(A,B)>0 → CHARINDEX(B,A)>0 (引数順序逆転に注意)
 *   - || → +
 *   - NVL → ISNULL
 *   - TO_CHAR(date, format) → FORMAT(date, format)
 *   - :param → ? (PreparedStatement)
 */
@Repository
public class Mcm1001uRepository {

    private final JdbcTemplate jdbcTemplate;

    public Mcm1001uRepository(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    // ========================================================
    // 納入先検索
    // 【変換元】Mcm1001uScreen.vb - SearchNonyusaki()
    //   元コード: Me.DaoContainer.Fill(Me.Mcm1001uDataSet, "MCM_MA_NONYUSAKI")
    //   条件: {0}=LIKE NONYUSAKI_CD, {1}=INSTR name, {2}=INSTR etc,
    //         {3}=EXISTS+LIKE SUPPORT_ID, {4}=EXISTS+LIKE PLANT_NK
    // ========================================================

    public List<Mcm1001uNonyusakiDto> searchNonyusaki(
            String nonyusakiCd, String supportId,
            String nonyusakiNk, String plantNk, String etc) {

        StringBuilder sql = new StringBuilder();
        List<Object> params = new ArrayList<>();

        sql.append("SELECT ");
        sql.append("  MAA.NONYUSAKI_ID, ");
        sql.append("  MAA.NONYUSAKI_CD, ");
        sql.append("  MAA.NONYUSAKI_NK, ");
        sql.append("  MAA.NONYUSAKIKANA_KN, ");
        sql.append("  MAA.NONYUSAKIEIMEI_EN, ");
        sql.append("  MAA.NONYUSAKIKOJO_NK, ");
        sql.append("  MAA.KYUNONYUSAKI_NK ");
        sql.append("FROM MCM.MCM_MA_NONYUSAKI MAA ");
        sql.append("WHERE 1 = 1 ");

        // 条件{0}: 納入先コード（LIKE）
        // 【変換元】Oracle: UPPER(MAA.NONYUSAKI_CD) LIKE UPPER(:NONYUSAKI_CD)
        if (isNotEmpty(nonyusakiCd)) {
            sql.append("  AND UPPER(MAA.NONYUSAKI_CD) LIKE UPPER(?) ");
            params.add(nonyusakiCd);
        }

        // 条件{1}: 納入先名フリーワード（CHARINDEX）
        // 【変換元】Oracle: INSTR(UPPER(MAA.NONYUSAKI_NK||...), UPPER(:NK))>0
        if (isNotEmpty(nonyusakiNk)) {
            sql.append("  AND CHARINDEX(UPPER(?), UPPER( ");
            sql.append("    ISNULL(MAA.NONYUSAKI_NK, '') + ' ' + ");
            sql.append("    ISNULL(MAA.NONYUSAKIKANA_KN, '') + ' ' + ");
            sql.append("    ISNULL(MAA.NONYUSAKIEIMEI_EN, '') + ' ' + ");
            sql.append("    ISNULL(MAA.NONYUSAKIKOJO_NK, '') + ' ' + ");
            sql.append("    ISNULL(MAA.KYUNONYUSAKI_NK, '') ");
            sql.append("  )) > 0 ");
            params.add(nonyusakiNk.trim());
        }

        // 条件{2}: その他（住所/郵便番号/TEL/FAX/備考 CHARINDEX）
        // 【変換元】Oracle: INSTR(UPPER(MAA.YUBIN_NO||...),UPPER(:ETC))>0
        if (isNotEmpty(etc)) {
            sql.append("  AND CHARINDEX(UPPER(?), UPPER( ");
            sql.append("    ISNULL(MAA.YUBIN_NO, '') + ' ' + ");
            sql.append("    ISNULL(MAA.JUSYO1_NK, '') + ' ' + ");
            sql.append("    ISNULL(MAA.JUSYO2_NK, '') + ' ' + ");
            sql.append("    ISNULL(MAA.KUNI_NK, '') + ' ' + ");
            sql.append("    ISNULL(MAA.TEL_NO, '') + ' ' + ");
            sql.append("    ISNULL(MAA.FAX_NO, '') + ' ' + ");
            sql.append("    ISNULL(MAA.BIKO, '') ");
            sql.append("  )) > 0 ");
            params.add(etc.trim());
        }

        // 条件{3}: サポートID（EXISTS + LIKE）
        if (isNotEmpty(supportId)) {
            sql.append("  AND EXISTS ( ");
            sql.append("    SELECT 1 FROM MCM.MCM_MA_PLANT MAP ");
            sql.append("    WHERE MAP.NONYUSAKI_ID = MAA.NONYUSAKI_ID ");
            sql.append("      AND UPPER(MAP.SUPPORT_ID) LIKE UPPER(?) ");
            sql.append("  ) ");
            params.add(supportId);
        }

        // 条件{4}: プラント名（EXISTS + LIKE）
        if (isNotEmpty(plantNk)) {
            sql.append("  AND EXISTS ( ");
            sql.append("    SELECT 1 FROM MCM.MCM_MA_PLANT MAP ");
            sql.append("    WHERE MAP.NONYUSAKI_ID = MAA.NONYUSAKI_ID ");
            sql.append("      AND UPPER(MAP.PLANT_NK) LIKE UPPER(?) ");
            sql.append("  ) ");
            params.add(plantNk);
        }

        sql.append("ORDER BY MAA.NONYUSAKI_CD ");

        return jdbcTemplate.query(sql.toString(), params.toArray(), new NonyusakiRowMapper());
    }

    // ========================================================
    // プラント検索（納入先ID指定）
    // 【変換元】Mcm1001uScreen.vb - DataGridView_NONYUSAKI_SelectionChanged
    //   Oracle: TO_CHAR → SQL Server: FORMAT
    // ========================================================

    public List<Mcm1001uPlantDto> searchPlant(BigDecimal nonyusakiId) {

        String sql =
            "SELECT " +
            "  MAP.PLANT_ID, " +
            "  MAP.NONYUSAKI_ID, " +
            "  MAP.SUPPORT_ID, " +
            "  MAP.PLANT_NK, " +
            "  FORMAT(MAP.NONYU_DT, 'yyyy/MM/dd') AS NONYU_DT, " +
            "  FORMAT(MAP.HOSYUSYUSOKU_DT, 'yyyy/MM/dd') AS HOSYUSYUSOKU_DT, " +
            "  MAP.SYSTEMKADOYOUBI, " +
            "  MAP.SYSTEMKADONISSU, " +
            "  MAP.SYSTEMKADOJIKAN, " +
            "  MAA.NONYUSAKI_CD, " +
            "  MAA.NONYUSAKI_NK " +
            "FROM MCM.MCM_MA_PLANT MAP " +
            "INNER JOIN MCM.MCM_MA_NONYUSAKI MAA " +
            "  ON MAA.NONYUSAKI_ID = MAP.NONYUSAKI_ID " +
            "WHERE MAP.NONYUSAKI_ID = ? " +
            "ORDER BY MAP.SUPPORT_ID";

        return jdbcTemplate.query(sql, new Object[]{nonyusakiId}, new PlantRowMapper());
    }

    // ========================================================
    // 全プラント検索（検索条件付き）
    // 【変換元】Mcm1001uScreen.vb - 検索実行時
    // ========================================================

    public List<Mcm1001uPlantDto> searchPlantWithConditions(
            String nonyusakiCd, String supportId,
            String nonyusakiNk, String plantNk, String etc) {

        StringBuilder sql = new StringBuilder();
        List<Object> params = new ArrayList<>();

        sql.append("SELECT ");
        sql.append("  MAP.PLANT_ID, ");
        sql.append("  MAP.NONYUSAKI_ID, ");
        sql.append("  MAP.SUPPORT_ID, ");
        sql.append("  MAP.PLANT_NK, ");
        sql.append("  FORMAT(MAP.NONYU_DT, 'yyyy/MM/dd') AS NONYU_DT, ");
        sql.append("  FORMAT(MAP.HOSYUSYUSOKU_DT, 'yyyy/MM/dd') AS HOSYUSYUSOKU_DT, ");
        sql.append("  MAP.SYSTEMKADOYOUBI, ");
        sql.append("  MAP.SYSTEMKADONISSU, ");
        sql.append("  MAP.SYSTEMKADOJIKAN, ");
        sql.append("  MAA.NONYUSAKI_CD, ");
        sql.append("  MAA.NONYUSAKI_NK ");
        sql.append("FROM MCM.MCM_MA_PLANT MAP ");
        sql.append("INNER JOIN MCM.MCM_MA_NONYUSAKI MAA ");
        sql.append("  ON MAA.NONYUSAKI_ID = MAP.NONYUSAKI_ID ");
        sql.append("WHERE 1 = 1 ");

        if (isNotEmpty(nonyusakiCd)) {
            sql.append("  AND UPPER(MAA.NONYUSAKI_CD) LIKE UPPER(?) ");
            params.add(nonyusakiCd);
        }

        if (isNotEmpty(nonyusakiNk)) {
            sql.append("  AND CHARINDEX(UPPER(?), UPPER( ");
            sql.append("    ISNULL(MAA.NONYUSAKI_NK, '') + ' ' + ");
            sql.append("    ISNULL(MAA.NONYUSAKIKANA_KN, '') + ' ' + ");
            sql.append("    ISNULL(MAA.NONYUSAKIEIMEI_EN, '') + ' ' + ");
            sql.append("    ISNULL(MAA.NONYUSAKIKOJO_NK, '') + ' ' + ");
            sql.append("    ISNULL(MAA.KYUNONYUSAKI_NK, '') ");
            sql.append("  )) > 0 ");
            params.add(nonyusakiNk.trim());
        }

        if (isNotEmpty(etc)) {
            sql.append("  AND CHARINDEX(UPPER(?), UPPER( ");
            sql.append("    ISNULL(MAA.YUBIN_NO, '') + ' ' + ");
            sql.append("    ISNULL(MAA.JUSYO1_NK, '') + ' ' + ");
            sql.append("    ISNULL(MAA.JUSYO2_NK, '') + ' ' + ");
            sql.append("    ISNULL(MAA.KUNI_NK, '') + ' ' + ");
            sql.append("    ISNULL(MAA.TEL_NO, '') + ' ' + ");
            sql.append("    ISNULL(MAA.FAX_NO, '') + ' ' + ");
            sql.append("    ISNULL(MAA.BIKO, '') ");
            sql.append("  )) > 0 ");
            params.add(etc.trim());
        }

        if (isNotEmpty(supportId)) {
            sql.append("  AND UPPER(MAP.SUPPORT_ID) LIKE UPPER(?) ");
            params.add(supportId);
        }

        if (isNotEmpty(plantNk)) {
            sql.append("  AND UPPER(MAP.PLANT_NK) LIKE UPPER(?) ");
            params.add(plantNk);
        }

        sql.append("ORDER BY MAA.NONYUSAKI_CD, MAP.SUPPORT_ID ");

        return jdbcTemplate.query(sql.toString(), params.toArray(), new PlantRowMapper());
    }

    // ========================================================
    // RowMapper
    // ========================================================

    private static class NonyusakiRowMapper implements RowMapper<Mcm1001uNonyusakiDto> {
        @Override
        public Mcm1001uNonyusakiDto mapRow(ResultSet rs, int rowNum) throws SQLException {
            Mcm1001uNonyusakiDto dto = new Mcm1001uNonyusakiDto();
            dto.setNonyusakiId(rs.getBigDecimal("NONYUSAKI_ID"));
            dto.setNonyusakiCd(rs.getString("NONYUSAKI_CD"));
            dto.setNonyusakiNk(rs.getString("NONYUSAKI_NK"));
            dto.setNonyusakikanaKn(rs.getString("NONYUSAKIKANA_KN"));
            dto.setNonyusakieimeiEn(rs.getString("NONYUSAKIEIMEI_EN"));
            dto.setNonyusakikojoNk(rs.getString("NONYUSAKIKOJO_NK"));
            dto.setKyunonyusakiNk(rs.getString("KYUNONYUSAKI_NK"));
            return dto;
        }
    }

    private static class PlantRowMapper implements RowMapper<Mcm1001uPlantDto> {
        @Override
        public Mcm1001uPlantDto mapRow(ResultSet rs, int rowNum) throws SQLException {
            Mcm1001uPlantDto dto = new Mcm1001uPlantDto();
            dto.setPlantId(rs.getBigDecimal("PLANT_ID"));
            dto.setNonyusakiId(rs.getBigDecimal("NONYUSAKI_ID"));
            dto.setSupportId(rs.getString("SUPPORT_ID"));
            dto.setPlantNk(rs.getString("PLANT_NK"));
            dto.setNonyuDt(rs.getString("NONYU_DT"));
            dto.setHosyusyusokuDt(rs.getString("HOSYUSYUSOKU_DT"));
            dto.setSystemkadoyoubi(rs.getString("SYSTEMKADOYOUBI"));
            dto.setSystemkadonissu(rs.getBigDecimal("SYSTEMKADONISSU"));
            dto.setSystemkadojikan(rs.getBigDecimal("SYSTEMKADOJIKAN"));
            dto.setNonyusakiCd(rs.getString("NONYUSAKI_CD"));
            dto.setNonyusakiNk(rs.getString("NONYUSAKI_NK"));
            return dto;
        }
    }

    private static boolean isNotEmpty(String str) {
        return str != null && !str.trim().isEmpty();
    }
}
