package com.daifuku.mcm.repository;

import com.daifuku.mcm.dto.Mcm1003uRowDto;
import com.daifuku.mcm.form.Mcm1003uForm;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.RowMapper;
import org.springframework.jdbc.core.namedparam.MapSqlParameterSource;
import org.springframework.jdbc.core.namedparam.NamedParameterJdbcTemplate;
import org.springframework.stereotype.Repository;

import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

/**
 * 【変換元】Mcm1003uDataSet.Designer.vb - UVATableAdapter + Mcm1003u_UVATableAdapter.xml
 * 【変換元行数】約672,000行（DataSet.Designer.vb全体）
 * 【作成者】T.Kajimura (2008/02/14)
 * 【更新者】T.Iwasawa (2009/08/12), M.Ohsuka (2009/12/02, 2010/01/15)
 *
 * MCM_1003_V ビューに対する動的検索を実行するRepository。
 *
 * 元コード:
 *   UVATableAdapter.xml で SELECT/FROM/WHERE/ORDERBY を XML定義し、
 *   CPBaseTableAdapter が動的にSQLを構築 → ODP.NETでOracle実行。
 *   各CONDITION要素の PREFIX/SUFFIX属性で LIKE パターンを制御。
 *
 * 変換後:
 *   NamedParameterJdbcTemplate で SQL Server 用の動的SQLを構築。
 *   Oracle構文 → SQL Server構文への変換を実施済み。
 *
 * 【SQL変換】
 *   TRUNC(date, 'DD')      → CAST(date AS DATE)
 *   TO_DATE(:p, 'YYYY/MM/DD') → CAST(@p AS DATE)
 *   INSTR(:p, col) > 0     → CHARINDEX(col, @p) > 0
 *   ||                     → +
 *   TO_NUMBER(col)          → CAST(col AS DECIMAL)
 */
@Repository
public class Mcm1003uViewRepository {

    private static final Logger log = LoggerFactory.getLogger(Mcm1003uViewRepository.class);

    private final NamedParameterJdbcTemplate jdbcTemplate;

    public Mcm1003uViewRepository(NamedParameterJdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    /**
     * 取引先見積・契約一覧を動的条件で検索する。
     *
     * 【変換元】Mcm1003u_UVATableAdapter.xml - SELECT/WHERE/ORDERBY
     *   元コード: CPBaseTableAdapter.Fill() → XML定義のSQL動的構築 → Oracle実行
     *
     * @param form 検索フォーム（null/空のフィールドは条件に含めない）
     * @return 検索結果リスト
     */
    public List<Mcm1003uRowDto> search(Mcm1003uForm form) {
        return query(form, null);
    }

    public List<Mcm1003uRowDto> findByKeiyakujikanId(BigDecimal id) {
        if (id == null || id.signum() <= 0) return List.of();
        return query(null, id);
    }

    private List<Mcm1003uRowDto> query(Mcm1003uForm form, BigDecimal id) {

        MapSqlParameterSource params = new MapSqlParameterSource();

        /*
         * 【変換元】Mcm1003u_UVATableAdapter.xml - SELECT句
         *   元コード:
         *     CASE WHEN TVA.FUSEI IS NOT NULL
         *       THEN TVA.KEIYAKU_NO || ' ' || TVA.FUSEI
         *       ELSE TVA.KEIYAKU_NO END KEIYAKU_NO
         *   → SQL Server: || を + に変換
         */
        StringBuilder sql = new StringBuilder();
        sql.append("SELECT ");
        sql.append("  TVA.TM_IRAI_ID, ");
        sql.append("  TVA.TM_KEIYAKUJIKAN_ID, ");
        sql.append("  TVA.JOTAI, ");
        sql.append("  TVA.TM_IRAI_NO, ");
        sql.append("  TVA.MITSUMORI_DT, ");
        sql.append("  TVA.IRAITANTOSYA, ");
        sql.append("  TVA.KEIYAKUJIKANTAI, ");
        sql.append("  CASE WHEN TVA.FUSEI IS NOT NULL ");
        sql.append("    THEN TVA.KEIYAKU_NO + ' ' + TVA.FUSEI ");
        sql.append("    ELSE TVA.KEIYAKU_NO END AS KEIYAKU_NO, ");
        sql.append("  TVA.FUSEI, ");
        sql.append("  TVA.TORIHIKISAKI_ID, ");
        sql.append("  TVA.TORIHIKISAKI_CD, ");
        sql.append("  TVA.TORIHIKISAKI_NK, ");
        sql.append("  TVA.NONYUSAKI_ID, ");
        sql.append("  TVA.NONYUSAKI_CD, ");
        sql.append("  TVA.NONYUSAKI_NK, ");
        sql.append("  TVA.KYUNONYUSAKI_NK, ");
        sql.append("  TVA.NONYUSAKIKOJO_NK, ");
        sql.append("  TVA.PLANT_ID, ");
        sql.append("  TVA.SUPPORT_ID, ");
        sql.append("  TVA.PLANT_NK, ");
        sql.append("  TVA.HOSYUHOHO, ");
        sql.append("  TVA.TENKENUMU, ");
        sql.append("  TVA.TENKENKANOYOBI, ");
        sql.append("  TVA.YAKANTAIOUMU, ");
        sql.append("  TVA.MITSUMORI_IRAI, ");
        sql.append("  TVA.MITSUMORI_COPY, ");
        sql.append("  TVA.MITSUMORI_DEL, ");
        sql.append("  TVA.KEIYAKU_KEIYAKU, ");
        sql.append("  TVA.KEIYAKU_DEL, ");
        sql.append("  TVA.TK_KEIYAKU_ID, ");
        sql.append("  TVA.SHONINJOTAI, ");
        sql.append("  TVA.KAITOKIZITSU_DT ");
        sql.append("FROM MCM.MCM_1003_V TVA ");
        sql.append("WHERE 1 = 1 ");
        if (id != null) {
            sql.append("AND TVA.TM_KEIYAKUJIKAN_ID = :id");
            return jdbcTemplate.query(sql.toString(), new MapSqlParameterSource("id", id), new Mcm1003uRowMapper());
        }

        /*
         * 【変換元】Mcm1003u_UVATableAdapter.xml - WHERE句 CONDITION[0]
         *   元コード: UPPER(TVA.TM_IRAI_NO) LIKE UPPER(:p0)  SUFFIX="%"
         *   → LIKE @param + '%'（前方一致）
         */
        if (isNotEmpty(form.getMitsumoriIraiNo())) {
            sql.append("AND UPPER(TVA.TM_IRAI_NO) LIKE UPPER(:mitsumoriIraiNo) ");
            params.addValue("mitsumoriIraiNo", form.getMitsumoriIraiNo() + "%");
        }

        /*
         * 【変換元】CONDITION[1]
         *   元コード: UPPER(TVA.KEIYAKU_NO) LIKE UPPER(:p1)  PREFIX="%" SUFFIX="%"
         *   → LIKE '%' + @param + '%'（部分一致）
         */
        if (isNotEmpty(form.getKeiyakuNo())) {
            sql.append("AND UPPER(TVA.KEIYAKU_NO) LIKE UPPER(:keiyakuNo) ");
            params.addValue("keiyakuNo", "%" + form.getKeiyakuNo() + "%");
        }

        /*
         * 【変換元】CONDITION[2]
         *   元コード: TRUNC(TVA.MITSUMORI_DT, 'DD') >= TO_DATE(:p2, 'YYYY/MM/DD')
         *   → SQL Server: CAST(TVA.MITSUMORI_DT AS DATE) >= CAST(:iraibi1 AS DATE)
         */
        if (isNotEmpty(form.getIraibi1())) {
            sql.append("AND CAST(TVA.MITSUMORI_DT AS DATE) >= CAST(:iraibi1 AS DATE) ");
            params.addValue("iraibi1", java.sql.Date.valueOf(java.time.LocalDate.parse(form.getIraibi1(), java.time.format.DateTimeFormatter.ofPattern("uuuu/MM/dd").withResolverStyle(java.time.format.ResolverStyle.STRICT))));
        }

        /*
         * 【変換元】CONDITION[3]
         *   元コード: TO_DATE(:p3, 'YYYY/MM/DD') >= TRUNC(TVA.MITSUMORI_DT, 'DD')
         */
        if (isNotEmpty(form.getIraibi2())) {
            sql.append("AND CAST(:iraibi2 AS DATE) >= CAST(TVA.MITSUMORI_DT AS DATE) ");
            params.addValue("iraibi2", java.sql.Date.valueOf(java.time.LocalDate.parse(form.getIraibi2(), java.time.format.DateTimeFormatter.ofPattern("uuuu/MM/dd").withResolverStyle(java.time.format.ResolverStyle.STRICT))));
        }

        /*
         * 【変換元】CONDITION[4]
         *   元コード: UPPER(TVA.IRAITANTOSYA) LIKE UPPER(:p4)  PREFIX="%" SUFFIX="%"
         */
        if (isNotEmpty(form.getIraisya())) {
            sql.append("AND UPPER(TVA.IRAITANTOSYA) LIKE UPPER(:iraisya) ");
            params.addValue("iraisya", "%" + form.getIraisya() + "%");
        }

        /*
         * 【変換元】CONDITION[5]
         *   元コード: TVA.TORIHIKISAKI_CD LIKE :p5  SUFFIX="%"
         */
        if (isNotEmpty(form.getTorihikisakiCd())) {
            sql.append("AND TVA.TORIHIKISAKI_CD LIKE :torihikisakiCd ");
            params.addValue("torihikisakiCd", form.getTorihikisakiCd() + "%");
        }

        /*
         * 【変換元】CONDITION[6]
         *   元コード: UPPER(TVA.TORIHIKISAKI_NK) LIKE UPPER(:p6)  PREFIX="%" SUFFIX="%"
         */
        if (isNotEmpty(form.getTorihikisakimei())) {
            sql.append("AND UPPER(TVA.TORIHIKISAKI_NK) LIKE UPPER(:torihikisakimei) ");
            params.addValue("torihikisakimei", "%" + form.getTorihikisakimei() + "%");
        }

        /*
         * 【変換元】CONDITION[7]
         *   元コード: TVA.NONYUSAKI_CD LIKE :p7  SUFFIX="%"
         */
        if (isNotEmpty(form.getNonyusakiCd())) {
            sql.append("AND TVA.NONYUSAKI_CD LIKE :nonyusakiCd ");
            params.addValue("nonyusakiCd", form.getNonyusakiCd() + "%");
        }

        /*
         * 【変換元】CONDITION[8]
         *   元コード: UPPER(TVA.NONYUSAKI_NK || ' ' || TVA.KYUNONYUSAKI_NK || ' '
         *             || TVA.NONYUSAKIKANA_KN || ' ' || TVA.NONYUSAKIEIMEI_EN) LIKE UPPER(:p8)
         *   PREFIX="%" SUFFIX="%"
         *   → SQL Server: || を + に変換
         */
        if (isNotEmpty(form.getNonyusakimei())) {
            sql.append("AND UPPER(ISNULL(TVA.NONYUSAKI_NK,'') + ' ' + ISNULL(TVA.KYUNONYUSAKI_NK,'') ");
            sql.append("  + ' ' + ISNULL(TVA.NONYUSAKIKANA_KN,'') + ' ' + ISNULL(TVA.NONYUSAKIEIMEI_EN,'')) ");
            sql.append("  LIKE UPPER(:nonyusakimei) ");
            params.addValue("nonyusakimei", "%" + form.getNonyusakimei() + "%");
        }

        /*
         * 【変換元】CONDITION[9]
         *   元コード: TVA.SUPPORT_ID LIKE :p9  SUFFIX="%"
         */
        if (isNotEmpty(form.getSupportId())) {
            sql.append("AND TVA.SUPPORT_ID LIKE :supportId ");
            params.addValue("supportId", form.getSupportId() + "%");
        }

        /*
         * 【変換元】CONDITION[10]
         *   元コード: UPPER(TVA.PLANT_NK) LIKE UPPER(:p10)  PREFIX="%" SUFFIX="%"
         */
        if (isNotEmpty(form.getPlantmei())) {
            sql.append("AND UPPER(TVA.PLANT_NK) LIKE UPPER(:plantmei) ");
            params.addValue("plantmei", "%" + form.getPlantmei() + "%");
        }

        /*
         * 【変換元】CONDITION[11]
         *   元コード: INSTR(:p11, TVA.JOTAI) > 0
         *   → SQL Server: CHARINDEX(TVA.JOTAI, :jotaiList) > 0
         *
         *   元コード（Screen.vb SearchButton_Click）:
         *     状態チェックボックスのチェック状態を連結して渡す。
         *     例: "作成中,依頼,見積,契約,"
         */
        String jotaiList = buildJotaiList(form);
        if (isNotEmpty(jotaiList)) {
            sql.append("AND CHARINDEX(TVA.JOTAI, :jotaiList) > 0 ");
            params.addValue("jotaiList", jotaiList);
        }

        /*
         * 【変換元】CONDITION[12]
         *   元コード: INSTR(:p12, TVA.SHONINJOTAI) > 0
         *   → SQL Server: CHARINDEX(TVA.SHONINJOTAI, :shoninList) > 0
         */
        String shoninList = buildShoninList(form);
        if (isNotEmpty(shoninList)) {
            sql.append("AND CHARINDEX(TVA.SHONINJOTAI, :shoninList) > 0 ");
            params.addValue("shoninList", shoninList);
        }

        /*
         * 【変換元】Mcm1003u_UVATableAdapter.xml - ORDERBY句
         *   元コード: TVA.TM_IRAI_NO DESC, TO_NUMBER(TVA.KEIYAKUJIKANTAI) ASC
         *   → SQL Server: TRY_CAST(TVA.KEIYAKUJIKANTAI AS DECIMAL) ASC (空文字エラー対策)
         */
        sql.append("ORDER BY TVA.TM_IRAI_NO DESC, TRY_CAST(TVA.KEIYAKUJIKANTAI AS DECIMAL) ASC ");

        log.debug("MCM1003U 検索SQL: {}", sql);
        return jdbcTemplate.query(sql.toString(), params, new Mcm1003uRowMapper());
    }

    // ===================================================================
    // 状態チェックボックスリスト構築
    // ===================================================================

    /**
     * 状態チェックボックスの選択状態から検索用文字列を構築する。
     *
     * 【変換元】Mcm1003uScreen.vb - SearchButton_Click()
     *   元コード:
     *     If Me.SagyochuCheckBox.Checked Then strJotai &= "作成中,"
     *     If Me.IraiCheckBox.Checked     Then strJotai &= "依頼,"
     *     If Me.MitsumoriCheckBox.Checked Then strJotai &= "見積,"
     *     If Me.KeiyakuCheckBox.Checked   Then strJotai &= "契約,"
     *     If Me.HakiCheckBox.Checked      Then strJotai &= "破棄,"
     *     If Me.KaiyakuCheckBox.Checked   Then strJotai &= "解約,"
     */
    private String buildJotaiList(Mcm1003uForm form) {
        StringBuilder sb = new StringBuilder();
        if (form.isSagyochu())  sb.append("作成中,");
        if (form.isIrai())      sb.append("依頼,");
        if (form.isMitsumori()) sb.append("見積,");
        if (form.isKeiyaku())   sb.append("契約,");
        if (form.isHaki())      sb.append("破棄,");
        if (form.isKaiyaku())   sb.append("解約,");
        return sb.toString();
    }

    /**
     * 承認状態チェックボックスの選択状態から検索用文字列を構築する。
     *
     * 【変換元】Mcm1003uScreen.vb - SearchButton_Click()
     *   元コード:
     *     If Me.ShoninSakuseichuCheckBox.Checked   Then strShonin &= "作成中,"
     *     If Me.ShoninSinsachuCheckBox.Checked      Then strShonin &= "審査中,"
     *     If Me.ShoninShoninchuCheckBox.Checked      Then strShonin &= "承認中,"
     *     If Me.ShoninShoninzumiCheckBox.Checked     Then strShonin &= "承認済,"
     *     If Me.ShoninSashimodoshiCheckBox.Checked   Then strShonin &= "差し戻し,"
     */
    private String buildShoninList(Mcm1003uForm form) {
        StringBuilder sb = new StringBuilder();
        if (form.isShoninSakuseichu())   sb.append("作成中,");
        if (form.isShoninSinsachu())     sb.append("審査中,");
        if (form.isShoninShoninchu())    sb.append("承認中,");
        if (form.isShoninShoninzumi())   sb.append("承認済,");
        if (form.isShoninSashimodoshi()) sb.append("差し戻し,");
        return sb.toString();
    }

    /**
     * 文字列が空でないかチェックする。
     */
    private boolean isNotEmpty(String value) {
        return value != null && !value.trim().isEmpty();
    }

    // ===================================================================
    // RowMapper
    // ===================================================================

    /**
     * MCM_1003_V ビューの検索結果を Mcm1003uRowDto にマッピングする。
     *
     * 【変換元】UVADataTable カラム定義からマッピング
     */
    private static class Mcm1003uRowMapper implements RowMapper<Mcm1003uRowDto> {
        @Override
        public Mcm1003uRowDto mapRow(ResultSet rs, int rowNum) throws SQLException {
            Mcm1003uRowDto dto = new Mcm1003uRowDto();

            // === ID系（BigDecimal） ===
            dto.setTmIraiId(rs.getBigDecimal("TM_IRAI_ID"));
            dto.setTmKeiyakujikanId(rs.getBigDecimal("TM_KEIYAKUJIKAN_ID"));
            dto.setTorihikisakiId(rs.getBigDecimal("TORIHIKISAKI_ID"));
            dto.setNonyusakiId(rs.getBigDecimal("NONYUSAKI_ID"));
            dto.setPlantId(rs.getBigDecimal("PLANT_ID"));
            dto.setTkKeiyakuId(rs.getBigDecimal("TK_KEIYAKU_ID"));

            // === 文字列系 ===
            dto.setJotai(rs.getString("JOTAI"));
            dto.setTmIraiNo(rs.getString("TM_IRAI_NO"));
            dto.setKeiyakujikantai(rs.getString("KEIYAKUJIKANTAI"));
            dto.setKeiyakuNo(rs.getString("KEIYAKU_NO"));
            dto.setFusei(rs.getString("FUSEI"));
            dto.setTorihikisakiCd(rs.getString("TORIHIKISAKI_CD"));
            dto.setTorihikisakiNk(rs.getString("TORIHIKISAKI_NK"));
            dto.setNonyusakiCd(rs.getString("NONYUSAKI_CD"));
            dto.setNonyusakiNk(rs.getString("NONYUSAKI_NK"));
            dto.setKyunonyusakiNk(rs.getString("KYUNONYUSAKI_NK"));
            dto.setNonyusakikojoNk(rs.getString("NONYUSAKIKOJO_NK"));
            dto.setSupportId(rs.getString("SUPPORT_ID"));
            dto.setPlantNk(rs.getString("PLANT_NK"));
            dto.setHosyuhoho(rs.getString("HOSYUHOHO"));
            dto.setTenkenumu(rs.getString("TENKENUMU"));
            dto.setTenkenkanoyobi(rs.getString("TENKENKANOYOBI"));
            dto.setYakantaioumu(rs.getString("YAKANTAIOUMU"));
            dto.setMitsumoriIrai(rs.getString("MITSUMORI_IRAI"));
            dto.setMitsumoriCopy(rs.getString("MITSUMORI_COPY"));
            dto.setMitsumoriDel(rs.getString("MITSUMORI_DEL"));
            dto.setKeiyakuKeiyaku(rs.getString("KEIYAKU_KEIYAKU"));
            dto.setKeiyakuDel(rs.getString("KEIYAKU_DEL"));
            dto.setShoninjotai(rs.getString("SHONINJOTAI"));
            dto.setIraitantosya(rs.getString("IRAITANTOSYA"));
            dto.setKaitokizitsuDt(rs.getString("KAITOKIZITSU_DT"));

            // === 日時系 ===
            Timestamp mitsumoriTs = rs.getTimestamp("MITSUMORI_DT");
            if (mitsumoriTs != null) {
                dto.setMitsumoriDt(mitsumoriTs.toLocalDateTime());
            }

            return dto;
        }
    }
}
