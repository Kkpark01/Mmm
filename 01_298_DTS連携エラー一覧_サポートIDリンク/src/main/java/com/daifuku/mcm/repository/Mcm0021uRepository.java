/**
 * 【変換元】Mcm0021uDataSet.Designer.vb（15,756行）
 *           Mcm0021u_MCM_MA_*.xml（5ファイル）
 *           Mcm0021uScreen.vb - getTkKeiyakuList() / getUkKeiyakuList()
 *   MCM0021U プラント付替え 専用リポジトリ
 *   元コード: 各TableAdapterのFill/GetData/スカラクエリ + 画面内の生SQL
 *
 *   Oracle → SQL Server 変換:
 *     INSTR(a, b) → CHARINDEX(b, a)
 *     || → + (文字列結合)
 *     NVL(a, b) → ISNULL(a, b)
 *     SYSDATE → GETDATE()
 *     TO_DATE(SYSDATE, 'YYYY/MM/DD') → CAST(GETDATE() AS DATE)
 *     :param → ? (JdbcTemplate positional)
 */
package com.daifuku.mcm.repository;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.RowMapper;
import org.springframework.stereotype.Repository;

import com.daifuku.mcm.dto.Mcm0021uKikikoseiDto;
import com.daifuku.mcm.dto.Mcm0021uKikimeisaiSearchDto;
import com.daifuku.mcm.dto.Mcm0021uMotoRowDto;
import com.daifuku.mcm.dto.Mcm0021uNonyusakiDto;
import com.daifuku.mcm.dto.Mcm0021uSakiRowDto;

/**
 * MCM0021U（プラント付替え）画面専用リポジトリ
 * <p>JdbcTemplateを使用した動的SQL組み立て + 複雑な結合クエリを実装する。</p>
 */
@Repository
public class Mcm0021uRepository {

    private final JdbcTemplate jdbcTemplate;

    public Mcm0021uRepository(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    // =================================================================
    // RowMapper定義
    // =================================================================

    /**
     * 納入先検索結果 RowMapper
     * 【変換元】MCM_MA_NONYUSAKIDataTable の列定義
     */
    private static final RowMapper<Mcm0021uNonyusakiDto> NONYUSAKI_ROW_MAPPER = (rs, rowNum) -> {
        Mcm0021uNonyusakiDto dto = new Mcm0021uNonyusakiDto();
        dto.setNonyusakiId(rs.getBigDecimal("NONYUSAKI_ID"));
        dto.setNonyusakiCd(rs.getString("NONYUSAKI_CD"));
        dto.setNonyusakiNk(rs.getString("NONYUSAKI_NK"));
        dto.setNonyusakikojoNk(rs.getString("NONYUSAKIKOJO_NK"));
        dto.setPlantId(rs.getBigDecimal("PLANT_ID"));
        dto.setSupportId(rs.getString("SUPPORT_ID"));
        dto.setPlantNk(rs.getString("PLANT_NK"));
        return dto;
    };

    /**
     * 機器構成+ブランド検索結果 RowMapper
     * 【変換元】MCM_MA_KIKIKOSEIDataTable の列定義
     */
    private static final RowMapper<Mcm0021uKikikoseiDto> KIKIKOSEI_ROW_MAPPER = (rs, rowNum) -> {
        Mcm0021uKikikoseiDto dto = new Mcm0021uKikikoseiDto();
        dto.setPlantId(rs.getBigDecimal("PLANT_ID"));
        dto.setKikikoseiId(rs.getBigDecimal("KIKIKOSEI_ID"));
        dto.setKikikoseiNk(rs.getString("KIKIKOSEI_NK"));
        dto.setSetNm(rs.getBigDecimal("SET_NM"));
        dto.setTehaiseiban(rs.getString("TEHAISEIBAN"));
        dto.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));
        dto.setBrandsyosaiNk(rs.getString("BRANDSYOSAI_NK"));
        dto.setBrandId(rs.getBigDecimal("BRAND_ID"));
        dto.setIkosaki(BigDecimal.ZERO);
        return dto;
    };

    /**
     * 機器明細検索結果 RowMapper
     * 【変換元】MCM_MA_KIKIMEISAIDataTable の列定義
     */
    private static final RowMapper<Mcm0021uKikimeisaiSearchDto> KIKIMEISAI_ROW_MAPPER = (rs, rowNum) -> {
        Mcm0021uKikimeisaiSearchDto dto = new Mcm0021uKikimeisaiSearchDto();
        dto.setSeizomakerNk(rs.getString("SEIZOMAKER_NK"));
        dto.setKikimeisaiId(rs.getBigDecimal("KIKIMEISAI_ID"));
        dto.setKikihinmeiNk(rs.getString("KIKIHINMEI_NK"));
        dto.setKikikatashiki(rs.getString("KIKIKATASHIKI"));
        dto.setSuryoNm(rs.getBigDecimal("SURYO_NM"));
        dto.setKikikoseiId(rs.getBigDecimal("KIKIKOSEI_ID"));
        dto.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));
        return dto;
    };

    /**
     * 付替え元 RowMapper
     * 【変換元】MCM_MA_KIKIKOSEI_MOTODataTable の列定義
     */
    private static final RowMapper<Mcm0021uMotoRowDto> MOTO_ROW_MAPPER = (rs, rowNum) -> {
        Mcm0021uMotoRowDto dto = new Mcm0021uMotoRowDto();
        dto.setNonyusakiId(rs.getBigDecimal("NONYUSAKI_ID"));
        dto.setNonyusakiCd(rs.getString("NONYUSAKI_CD"));
        dto.setNonyusakiNk(rs.getString("NONYUSAKI_NK"));
        dto.setNonyusakikojoNk(rs.getString("NONYUSAKIKOJO_NK"));
        dto.setPlantId(rs.getBigDecimal("PLANT_ID"));
        dto.setSupportId(rs.getString("SUPPORT_ID"));
        dto.setPlantNk(rs.getString("PLANT_NK"));
        dto.setKikikoseiId(rs.getBigDecimal("KIKIKOSEI_ID"));
        dto.setKikikoseiNk(rs.getString("KIKIKOSEI_NK"));
        dto.setSetNm(rs.getBigDecimal("SET_NM"));
        dto.setTehaiseiban(rs.getString("TEHAISEIBAN"));
        dto.setBrandsyosaiNk(rs.getString("BRANDSYOSAI_NK"));
        dto.setIdoCheck(1); // 初期値: チェックON
        dto.setHyojijun(rs.getBigDecimal("HYOJIJUN"));
        dto.setHyojijun1(rs.getBigDecimal("HYOJIJUN1"));
        dto.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));
        dto.setIkosaki(null); // 初期値: 未選択
        return dto;
    };

    /**
     * 付替え先 RowMapper
     * 【変換元】MCM_MA_KIKIKOSEI_SAKIDataTable の列定義
     */
    private static final RowMapper<Mcm0021uSakiRowDto> SAKI_ROW_MAPPER = (rs, rowNum) -> {
        Mcm0021uSakiRowDto dto = new Mcm0021uSakiRowDto();
        dto.setNonyusakiId(rs.getBigDecimal("NONYUSAKI_ID"));
        dto.setNonyusakiCd(rs.getString("NONYUSAKI_CD"));
        dto.setNonyusakiNk(rs.getString("NONYUSAKI_NK"));
        dto.setNonyusakikojoNk(rs.getString("NONYUSAKIKOJO_NK"));
        dto.setPlantId(rs.getBigDecimal("PLANT_ID"));
        dto.setSupportId(rs.getString("SUPPORT_ID"));
        dto.setPlantNk(rs.getString("PLANT_NK"));
        dto.setKikikoseiId(rs.getBigDecimal("KIKIKOSEI_ID"));
        dto.setKikikoseiNk(rs.getString("KIKIKOSEI_NK"));
        dto.setSetNm(rs.getBigDecimal("SET_NM"));
        dto.setTehaiseiban(rs.getString("TEHAISEIBAN"));
        dto.setBrandId(rs.getBigDecimal("BRAND_ID"));
        dto.setBrandsyosaiNk(rs.getString("BRANDSYOSAI_NK"));
        dto.setHyojijun(rs.getBigDecimal("HYOJIJUN"));
        dto.setHyojijun1(rs.getBigDecimal("HYOJIJUN1"));
        dto.setIdoCheck(1); // 付替え先は表示専用（固定値1）
        return dto;
    };

    // =================================================================
    // 1. 納入先検索（動的SQL）
    // 【変換元】Mcm0021u_MCM_MA_NONYUSAKITableAdapter.xml
    //   SearchButton_Click() → Me.Fill(MCM_MA_NONYUSAKI, ...)
    //   4つの動的条件: NONYUSAKI_CD, NONYUSAKI_NK, SUPPORT_ID, PLANT_NK
    // =================================================================
    /**
     * 納入先を動的条件で検索する
     * @param nonyusakiCd 納入先コード（部分一致、大文字小文字無視）
     * @param nonyusakiNk 納入先名（部分一致、旧名・カナ・英名含む、大文字小文字無視）
     * @param supportId サポートID（部分一致、大文字小文字無視）
     * @param plantNk プラント名（部分一致、大文字小文字無視）
     * @return 納入先+プラント結合リスト
     */
    public List<Mcm0021uNonyusakiDto> searchNonyusaki(
            String nonyusakiCd, String nonyusakiNk, String supportId, String plantNk) {

        StringBuilder sql = new StringBuilder();
        List<Object> params = new ArrayList<>();

        sql.append("SELECT MAA.NONYUSAKI_ID, MAA.NONYUSAKI_CD, MAA.NONYUSAKI_NK, ");
        sql.append("       MAA.NONYUSAKIKOJO_NK, MAB.PLANT_ID, MAB.SUPPORT_ID, MAB.PLANT_NK ");
        sql.append("FROM MCM.MCM_MA_NONYUSAKI MAA ");
        sql.append("INNER JOIN MCM.MCM_MA_PLANT MAB ON MAA.NONYUSAKI_ID = MAB.NONYUSAKI_ID ");
        sql.append("WHERE 1=1 ");

        /*
         * 【変換元】<CONDITION SUFFIX="%"> MAA.NONYUSAKI_CD LIKE :{0}
         */
        if (!isEmpty(nonyusakiCd)) {
            sql.append("AND UPPER(MAA.NONYUSAKI_CD) LIKE UPPER(?) ");
            params.add("%" + nonyusakiCd + "%");
        }
        /*
         * 【変換元】INSTR(UPPER(MAA.NONYUSAKI_NK || ' ' || MAA.KYUNONYUSAKI_NK
         *           || ' ' || MAA.NONYUSAKIKOJO_NK || ' ' || MAA.NONYUSAKIKANA_KN
         *           || ' ' || MAA.NONYUSAKIEIMEI_EN), UPPER(:{1})) > 0
         *   Oracle INSTR(a, b) → SQL Server CHARINDEX(b, a)
         *   Oracle || → SQL Server +
         *   NVL未使用のため ISNULL で NULL安全に変換
         */
        if (!isEmpty(nonyusakiNk)) {
            sql.append("AND CHARINDEX(UPPER(?), UPPER( ");
            sql.append("  ISNULL(MAA.NONYUSAKI_NK, '') + ' ' + ");
            sql.append("  ISNULL(MAA.KYUNONYUSAKI_NK, '') + ' ' + ");
            sql.append("  ISNULL(MAA.NONYUSAKIKOJO_NK, '') + ' ' + ");
            sql.append("  ISNULL(MAA.NONYUSAKIKANA_KN, '') + ' ' + ");
            sql.append("  ISNULL(MAA.NONYUSAKIEIMEI_EN, '') ");
            sql.append(")) > 0 ");
            params.add(nonyusakiNk);
        }
        /*
         * 【変換元】<CONDITION SUFFIX="%"> MAB.SUPPORT_ID LIKE :{2}
         */
        if (!isEmpty(supportId)) {
            sql.append("AND UPPER(MAB.SUPPORT_ID) LIKE UPPER(?) ");
            params.add("%" + supportId + "%");
        }
        /*
         * 【変換元】<CONDITION PREFIX="%" SUFFIX="%"> UPPER(MAB.PLANT_NK) LIKE UPPER(:{3})
         */
        if (!isEmpty(plantNk)) {
            sql.append("AND UPPER(MAB.PLANT_NK) LIKE UPPER(?) ");
            params.add("%" + plantNk + "%");
        }

        sql.append("ORDER BY MAA.NONYUSAKI_CD ASC, MAB.SUPPORT_ID ASC ");

        return jdbcTemplate.query(sql.toString(), NONYUSAKI_ROW_MAPPER, params.toArray());
    }

    // =================================================================
    // 2. 機器構成+ブランド検索（動的SQL）
    // 【変換元】Mcm0021u_MCM_MA_KIKIKOSEITableAdapter.xml
    //   SearchButton_Click() → Me.Fill(MCM_MA_KIKIKOSEI, ...)
    //   SET_NM = MAX((SELECT COUNT(S.KOTAIKANRI_ID) FROM MCM_MA_KIKIKOTAIKANRI S ...))
    //   EXISTS副問い合わせによる動的WHERE
    // =================================================================
    /**
     * 機器構成+ブランドを動的条件で検索する
     */
    public List<Mcm0021uKikikoseiDto> searchKikikosei(
            String nonyusakiCd, String nonyusakiNk, String supportId, String plantNk) {

        StringBuilder sql = new StringBuilder();
        List<Object> params = new ArrayList<>();

        sql.append("SELECT MAE.PLANT_ID, MAE.KIKIKOSEI_ID, MAE.KIKIKOSEI_NK, ");
        sql.append("  MAX(KC.SET_NM) AS SET_NM, ");
        sql.append("  MAE.TEHAISEIBAN, MAD.BRANDKOSEI_ID, ");
        sql.append("  MAD.BRANDSYOSAI_NK, MAC.BRAND_ID ");
        sql.append("FROM MCM.MCM_MA_KIKIKOSEI MAE ");
        sql.append("INNER JOIN MCM.MCM_MA_KIKIKOTAIKANRI MAG ON MAG.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID ");
        sql.append("INNER JOIN MCM.MCM_MA_BRAND_KOSEI MAD ON MAD.BRANDKOSEI_ID = MAG.BRANDKOSEI_ID ");
        sql.append("INNER JOIN MCM.MCM_MA_BRAND MAC ON MAC.BRAND_ID = MAD.BRAND_ID ");
        sql.append("OUTER APPLY (SELECT COUNT(S.KOTAIKANRI_ID) AS SET_NM ");
        sql.append("  FROM MCM.MCM_MA_KIKIKOTAIKANRI S ");
        sql.append("  WHERE S.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID AND S.BRANDKOSEI_ID = MAD.BRANDKOSEI_ID ");
        sql.append("  AND S.ATSUKAIKIKIKOSEI_ID = MAG.ATSUKAIKIKIKOSEI_ID) KC ");
        sql.append("WHERE 1=1 ");

        /* 動的WHERE: 各検索条件に対応するEXISTS副問い合わせ */
        if (!isEmpty(nonyusakiCd)) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ");
            sql.append("  INNER JOIN MCM.MCM_MA_PLANT MAB ON MAB.NONYUSAKI_ID = MAA.NONYUSAKI_ID ");
            sql.append("  WHERE MAB.PLANT_ID = MAE.PLANT_ID AND UPPER(MAA.NONYUSAKI_CD) LIKE UPPER(?)) ");
            params.add("%" + nonyusakiCd + "%");
        }
        if (!isEmpty(nonyusakiNk)) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ");
            sql.append("  INNER JOIN MCM.MCM_MA_PLANT MAB ON MAB.NONYUSAKI_ID = MAA.NONYUSAKI_ID ");
            sql.append("  WHERE MAB.PLANT_ID = MAE.PLANT_ID ");
            sql.append("  AND CHARINDEX(UPPER(?), UPPER( ");
            sql.append("    ISNULL(MAA.NONYUSAKI_NK,'') + ' ' + ISNULL(MAA.KYUNONYUSAKI_NK,'') + ' ' + ");
            sql.append("    ISNULL(MAA.NONYUSAKIKOJO_NK,'') + ' ' + ISNULL(MAA.NONYUSAKIKANA_KN,'') + ' ' + ");
            sql.append("    ISNULL(MAA.NONYUSAKIEIMEI_EN,''))) > 0) ");
            params.add(nonyusakiNk);
        }
        if (!isEmpty(supportId)) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ");
            sql.append("  INNER JOIN MCM.MCM_MA_PLANT MAB ON MAB.NONYUSAKI_ID = MAA.NONYUSAKI_ID ");
            sql.append("  WHERE MAB.PLANT_ID = MAE.PLANT_ID AND UPPER(MAB.SUPPORT_ID) LIKE UPPER(?)) ");
            params.add("%" + supportId + "%");
        }
        if (!isEmpty(plantNk)) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ");
            sql.append("  INNER JOIN MCM.MCM_MA_PLANT MAB ON MAB.NONYUSAKI_ID = MAA.NONYUSAKI_ID ");
            sql.append("  WHERE MAB.PLANT_ID = MAE.PLANT_ID AND UPPER(MAB.PLANT_NK) LIKE UPPER(?)) ");
            params.add("%" + plantNk + "%");
        }

        sql.append("GROUP BY MAE.PLANT_ID, MAE.KIKIKOSEI_ID, MAE.KIKIKOSEI_NK, ");
        sql.append("  MAE.TEHAISEIBAN, MAE.HYOJIJUN, MAD.BRANDKOSEI_ID, ");
        sql.append("  MAD.BRANDSYOSAI_NK, MAC.BRAND_ID, MAD.HYOJIJUN ");
        sql.append("ORDER BY MAE.PLANT_ID ASC, MAE.HYOJIJUN ASC, MAD.HYOJIJUN ASC ");

        return jdbcTemplate.query(sql.toString(), KIKIKOSEI_ROW_MAPPER, params.toArray());
    }

    // =================================================================
    // 3. 機器明細検索（動的SQL、MCM_MA_KOTAIMEISAI_V ビュー使用）
    // 【変換元】Mcm0021u_MCM_MA_KIKIMEISAITableAdapter.xml
    //   MCM_MA_KOTAIMEISAI_V ビューに対するEXISTS副問い合わせ
    // =================================================================
    /**
     * 機器明細を動的条件で検索する（MCM_MA_KOTAIMEISAI_Vビュー使用）
     */
    public List<Mcm0021uKikimeisaiSearchDto> searchKikimeisai(
            String nonyusakiCd, String nonyusakiNk, String supportId, String plantNk) {

        StringBuilder sql = new StringBuilder();
        List<Object> params = new ArrayList<>();

        sql.append("SELECT DISTINCT ");
        sql.append("  MAV.MAE_PLANT_ID AS PLANT_ID, ");
        sql.append("  MAV.MAF_HYOJIJUN AS HYOJIJUN, ");
        sql.append("  MAV.MAJ_SEIZOMAKER_NK AS SEIZOMAKER_NK, ");
        sql.append("  MAV.MAF_KIKIMEISAI_ID AS KIKIMEISAI_ID, ");
        sql.append("  MAV.MAF_KIKIHINMEI_NK AS KIKIHINMEI_NK, ");
        sql.append("  MAV.MAF_KIKIKATASHIKI AS KIKIKATASHIKI, ");
        sql.append("  MAV.MAF_SURYO_NM AS SURYO_NM, ");
        sql.append("  MAV.MAE_KIKIKOSEI_ID AS KIKIKOSEI_ID, ");
        sql.append("  MAV.MAD_BRANDKOSEI_ID AS BRANDKOSEI_ID ");
        sql.append("FROM MCM.MCM_MA_KOTAIMEISAI_V MAV ");
        sql.append("WHERE 1=1 ");

        if (!isEmpty(nonyusakiCd)) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ");
            sql.append("  INNER JOIN MCM.MCM_MA_PLANT MAB ON MAB.NONYUSAKI_ID = MAA.NONYUSAKI_ID ");
            sql.append("  WHERE MAB.PLANT_ID = MAV.MAE_PLANT_ID AND UPPER(MAA.NONYUSAKI_CD) LIKE UPPER(?)) ");
            params.add("%" + nonyusakiCd + "%");
        }
        if (!isEmpty(nonyusakiNk)) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ");
            sql.append("  INNER JOIN MCM.MCM_MA_PLANT MAB ON MAB.NONYUSAKI_ID = MAA.NONYUSAKI_ID ");
            sql.append("  WHERE MAB.PLANT_ID = MAV.MAE_PLANT_ID ");
            sql.append("  AND CHARINDEX(UPPER(?), UPPER( ");
            sql.append("    ISNULL(MAA.NONYUSAKI_NK,'') + ' ' + ISNULL(MAA.KYUNONYUSAKI_NK,'') + ' ' + ");
            sql.append("    ISNULL(MAA.NONYUSAKIKOJO_NK,'') + ' ' + ISNULL(MAA.NONYUSAKIKANA_KN,'') + ' ' + ");
            sql.append("    ISNULL(MAA.NONYUSAKIEIMEI_EN,''))) > 0) ");
            params.add(nonyusakiNk);
        }
        if (!isEmpty(supportId)) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ");
            sql.append("  INNER JOIN MCM.MCM_MA_PLANT MAB ON MAB.NONYUSAKI_ID = MAA.NONYUSAKI_ID ");
            sql.append("  WHERE MAB.PLANT_ID = MAV.MAE_PLANT_ID AND UPPER(MAB.SUPPORT_ID) LIKE UPPER(?)) ");
            params.add("%" + supportId + "%");
        }
        if (!isEmpty(plantNk)) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ");
            sql.append("  INNER JOIN MCM.MCM_MA_PLANT MAB ON MAB.NONYUSAKI_ID = MAA.NONYUSAKI_ID ");
            sql.append("  WHERE MAB.PLANT_ID = MAV.MAE_PLANT_ID AND UPPER(MAB.PLANT_NK) LIKE UPPER(?)) ");
            params.add("%" + plantNk + "%");
        }

        sql.append("ORDER BY MAV.MAE_PLANT_ID ASC, MAV.MAE_KIKIKOSEI_ID ASC, MAV.MAF_HYOJIJUN ASC ");

        return jdbcTemplate.query(sql.toString(), KIKIMEISAI_ROW_MAPPER, params.toArray());
    }

    // =================================================================
    // 4. 付替え元データ取得
    // 【変換元】Mcm0021u_MCM_MA_KIKIKOSEI_MOTOTableAdapter.xml
    //   TsukekaemotoButton_Click() → Me.Fill(MCM_MA_KIKIKOSEI_MOTO, plantId)
    //   6テーブル結合: KIKIKOSEI + KIKIKOTAIKANRI + BRAND_KOSEI + BRAND + PLANT + NONYUSAKI
    //   SET_NM = MAX((SELECT COUNT(S.KOTAIKANRI_ID) ...))
    //   初期値: 1 AS IDO_CHECK, '' AS IKOSAKI
    // =================================================================
    /**
     * 付替え元の機器構成データを取得する
     */
    /** #298: 個体管理と結合せず、受け取ったプラントIDに一致する基本情報を取得する。 */
    public Mcm0021uNonyusakiDto findPlantInfoById(BigDecimal plantId) {
        String sql = "SELECT MAA.NONYUSAKI_ID, MAA.NONYUSAKI_CD, MAA.NONYUSAKI_NK, "
                + "MAA.NONYUSAKIKOJO_NK, MAB.PLANT_ID, MAB.SUPPORT_ID, MAB.PLANT_NK "
                + "FROM MCM.MCM_MA_PLANT MAB "
                + "INNER JOIN MCM.MCM_MA_NONYUSAKI MAA ON MAA.NONYUSAKI_ID = MAB.NONYUSAKI_ID "
                + "WHERE MAB.PLANT_ID = ?";
        return jdbcTemplate.query(sql, NONYUSAKI_ROW_MAPPER, plantId).stream().findFirst().orElse(null);
    }

    public List<Mcm0021uMotoRowDto> findMotoByPlantId(BigDecimal plantId) {
        String sql =
            "SELECT " +
            "  MAA.NONYUSAKI_ID, MAA.NONYUSAKI_CD, MAA.NONYUSAKI_NK, " +
            "  MAA.NONYUSAKIKOJO_NK, MAE.PLANT_ID, MAB.SUPPORT_ID, MAB.PLANT_NK, " +
            "  MAE.KIKIKOSEI_ID, MAE.KIKIKOSEI_NK, " +
            "  MAX(KC.SET_NM) AS SET_NM, " +
            "  MAE.TEHAISEIBAN, MAD.BRANDKOSEI_ID, MAD.BRANDSYOSAI_NK, " +
            "  MAE.HYOJIJUN, MAD.HYOJIJUN AS HYOJIJUN1 " +
            "FROM MCM.MCM_MA_KIKIKOSEI MAE " +
            "INNER JOIN MCM.MCM_MA_KIKIKOTAIKANRI MAG ON MAG.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID " +
            "INNER JOIN MCM.MCM_MA_BRAND_KOSEI MAD ON MAD.BRANDKOSEI_ID = MAG.BRANDKOSEI_ID " +
            "INNER JOIN MCM.MCM_MA_BRAND MAC ON MAC.BRAND_ID = MAD.BRAND_ID " +
            "INNER JOIN MCM.MCM_MA_PLANT MAB ON MAB.PLANT_ID = MAD.PLANT_ID " +
            "INNER JOIN MCM.MCM_MA_NONYUSAKI MAA ON MAA.NONYUSAKI_ID = MAB.NONYUSAKI_ID " +
            "OUTER APPLY (SELECT COUNT(S.KOTAIKANRI_ID) AS SET_NM " +
            "  FROM MCM.MCM_MA_KIKIKOTAIKANRI S " +
            "  WHERE S.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID AND S.BRANDKOSEI_ID = MAD.BRANDKOSEI_ID " +
            "  AND S.ATSUKAIKIKIKOSEI_ID = MAG.ATSUKAIKIKIKOSEI_ID) KC " +
            "WHERE MAB.PLANT_ID = ? " +
            "GROUP BY MAA.NONYUSAKI_ID, MAA.NONYUSAKI_CD, MAA.NONYUSAKI_NK, " +
            "  MAA.NONYUSAKIKOJO_NK, MAE.PLANT_ID, MAB.SUPPORT_ID, MAB.PLANT_NK, " +
            "  MAE.KIKIKOSEI_ID, MAE.KIKIKOSEI_NK, MAE.TEHAISEIBAN, " +
            "  MAD.BRANDKOSEI_ID, MAD.BRANDSYOSAI_NK, MAE.HYOJIJUN, MAD.HYOJIJUN " +
            "ORDER BY MAE.HYOJIJUN ASC, MAD.HYOJIJUN ASC";

        return jdbcTemplate.query(sql, MOTO_ROW_MAPPER, plantId);
    }


    // =================================================================
    // 5. 付替え先データ取得
    // 【変換元】Mcm0021u_MCM_MA_KIKIKOSEI_SAKITableAdapter.xml
    //   TsukekaesakiButton_Click() → Me.Fill(MCM_MA_KIKIKOSEI_SAKI, plantId)
    //   付替え元と同構造だが BRAND_ID を取得（付替え先はブランドIDで表示）
    // =================================================================
    /**
     * 付替え先の機器構成データを取得する
     */
    public List<Mcm0021uSakiRowDto> findSakiByPlantId(BigDecimal plantId) {
        String sql =
            "SELECT " +
            "  MAA.NONYUSAKI_ID, MAA.NONYUSAKI_CD, MAA.NONYUSAKI_NK, " +
            "  MAA.NONYUSAKIKOJO_NK, MAE.PLANT_ID, MAB.SUPPORT_ID, MAB.PLANT_NK, " +
            "  MAE.KIKIKOSEI_ID, MAE.KIKIKOSEI_NK, " +
            "  MAX(KC.SET_NM) AS SET_NM, " +
            "  MAE.TEHAISEIBAN, MAD.BRAND_ID, MAD.BRANDSYOSAI_NK, " +
            "  MAE.HYOJIJUN, MAD.HYOJIJUN AS HYOJIJUN1 " +
            "FROM MCM.MCM_MA_KIKIKOSEI MAE " +
            "INNER JOIN MCM.MCM_MA_KIKIKOTAIKANRI MAG ON MAG.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID " +
            "INNER JOIN MCM.MCM_MA_BRAND_KOSEI MAD ON MAD.BRANDKOSEI_ID = MAG.BRANDKOSEI_ID " +
            "INNER JOIN MCM.MCM_MA_BRAND MAC ON MAC.BRAND_ID = MAD.BRAND_ID " +
            "INNER JOIN MCM.MCM_MA_PLANT MAB ON MAB.PLANT_ID = MAD.PLANT_ID " +
            "INNER JOIN MCM.MCM_MA_NONYUSAKI MAA ON MAA.NONYUSAKI_ID = MAB.NONYUSAKI_ID " +
            "OUTER APPLY (SELECT COUNT(S.KOTAIKANRI_ID) AS SET_NM " +
            "  FROM MCM.MCM_MA_KIKIKOTAIKANRI S " +
            "  WHERE S.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID AND S.BRANDKOSEI_ID = MAD.BRANDKOSEI_ID " +
            "  AND S.ATSUKAIKIKIKOSEI_ID = MAG.ATSUKAIKIKIKOSEI_ID) KC " +
            "WHERE MAB.PLANT_ID = ? " +
            "GROUP BY MAA.NONYUSAKI_ID, MAA.NONYUSAKI_CD, MAA.NONYUSAKI_NK, " +
            "  MAA.NONYUSAKIKOJO_NK, MAE.PLANT_ID, MAB.SUPPORT_ID, MAB.PLANT_NK, " +
            "  MAE.KIKIKOSEI_ID, MAE.KIKIKOSEI_NK, MAE.TEHAISEIBAN, " +
            "  MAD.BRAND_ID, MAD.BRANDSYOSAI_NK, MAE.HYOJIJUN, MAD.HYOJIJUN " +
            "ORDER BY MAE.HYOJIJUN ASC, MAD.HYOJIJUN ASC";

        return jdbcTemplate.query(sql, SAKI_ROW_MAPPER, plantId);
    }

    // =================================================================
    // 6. ブランド構成コンボ用データ取得
    // 【変換元】MCM_MA_BRAND_KOSEI_FOR_COMBOTableAdapter
    //   SELECT DISTINCT BRANDKOSEI_ID, BRANDSYOSAI_NK
    //   FROM MCM_MA_BRAND_KOSEI WHERE PLANT_ID = :plantId
    // =================================================================
    /**
     * 付替え先プラントのブランド構成リストを取得する（コンボボックス用）
     * @param plantId 付替え先プラントID
     * @return [BRANDKOSEI_ID, BRANDSYOSAI_NK] のリスト
     */
    public List<Object[]> findBrandKoseiForCombo(BigDecimal plantId) {
        String sql = "SELECT DISTINCT BRANDKOSEI_ID, BRANDSYOSAI_NK " +
                     "FROM MCM.MCM_MA_BRAND_KOSEI WHERE PLANT_ID = ? " +
                     "ORDER BY BRANDSYOSAI_NK ASC";
        return jdbcTemplate.query(sql, (rs, rowNum) ->
            new Object[] { rs.getBigDecimal("BRANDKOSEI_ID"), rs.getString("BRANDSYOSAI_NK") },
            plantId);
    }

    // =================================================================
    // 7. スカラクエリ群
    // 【変換元】MCM_MA_KIKIKOSEITableAdapter.MaxValueId() / MaxValueNo()
    //           MCM_MA_KIKIMEISAITableAdapter.MaxValueId() / MaxValueNo()
    //   Oracle NVL(MAX(...), 0) → SQL Server ISNULL(MAX(...), 0)
    // =================================================================

    /**
     * 機器構成IDの最大値を取得
     * 【変換元】MCM_MA_KIKIKOSEITableAdapter.MaxValueId()
     *   元コード: SELECT NVL(MAX(KIKIKOSEI_ID), 0) FROM MCM_MA_KIKIKOSEI
     */
    public BigDecimal getMaxKikikoseiId() {
        return jdbcTemplate.queryForObject(
            "SELECT ISNULL(MAX(KIKIKOSEI_ID), 0) FROM MCM.MCM_MA_KIKIKOSEI WITH (UPDLOCK,HOLDLOCK)",
            BigDecimal.class);
    }

    /**
     * 指定プラントの機器構成表示順の最大値を取得
     * 【変換元】MCM_MA_KIKIKOSEITableAdapter.MaxValueNo()
     *   元コード: SELECT NVL(MAX(HYOJIJUN), 0) FROM MCM_MA_KIKIKOSEI WHERE PLANT_ID = :plantId
     */
    public BigDecimal getMaxKikikoseiHyojijun(BigDecimal plantId) {
        return jdbcTemplate.queryForObject(
            "SELECT ISNULL(MAX(HYOJIJUN), 0) FROM MCM.MCM_MA_KIKIKOSEI WHERE PLANT_ID = ?",
            BigDecimal.class, plantId);
    }

    /**
     * 機器明細IDの最大値を取得
     * 【変換元】MCM_MA_KIKIMEISAITableAdapter.MaxValueId()
     *   元コード: SELECT NVL(MAX(KIKIMEISAI_ID), 0) FROM MCM_MA_KIKIMEISAI
     */
    public BigDecimal getMaxKikimeisaiId() {
        return jdbcTemplate.queryForObject(
            "SELECT ISNULL(MAX(KIKIMEISAI_ID), 0) FROM MCM.MCM_MA_KIKIMEISAI WITH (UPDLOCK,HOLDLOCK)",
            BigDecimal.class);
    }

    /**
     * 指定プラント+機器構成の機器明細表示順の最大値を取得
     * 【変換元】MCM_MA_KIKIMEISAITableAdapter.MaxValueNo()
     */
    public BigDecimal getMaxKikimeisaiHyojijun(BigDecimal plantId, BigDecimal kikikoseiId) {
        return jdbcTemplate.queryForObject(
            "SELECT ISNULL(MAX(MAF.HYOJIJUN), 0) " +
            "FROM MCM.MCM_MA_KIKIMEISAI MAF " +
            "INNER JOIN MCM.MCM_MA_KIKIKOSEI MAE ON MAF.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID " +
            "WHERE MAE.PLANT_ID = ? AND MAE.KIKIKOSEI_ID = ?",
            BigDecimal.class, plantId, kikikoseiId);
    }

    // =================================================================
    // 8. セット数取得（MCM_MA_KOTAIMEISAI_V ビュー）
    // 【変換元】MCM_MA_KOTAIMEISAI_V_DATA_MAXTableAdapter
    //   SELECT MAF_ATSUKAIKIKI_ID, COUNT(*) AS DATA_MAX
    //   FROM MCM_MA_KOTAIMEISAI_V
    //   WHERE MAE_KIKIKOSEI_ID = :kikikoseiId AND MAG_BRANDKOSEI_ID = :brandkoseiId
    //     AND MAM_KOTAIKANRI_FLG = 1
    //   GROUP BY MAF_ATSUKAIKIKI_ID
    // =================================================================
    /**
     * セット数を取得する（MCM_MA_KOTAIMEISAI_V ビュー使用）
     * @return [ATSUKAIKIKI_ID, DATA_MAX(件数)] のリスト
     */
    public List<Object[]> findKotaimeisaiVDataMax(BigDecimal kikikoseiId, BigDecimal brandkoseiId) {
        String sql =
            "SELECT MAF_ATSUKAIKIKI_ID, COUNT(*) AS DATA_MAX " +
            "FROM MCM.MCM_MA_KOTAIMEISAI_V " +
            "WHERE MAE_KIKIKOSEI_ID = ? AND MAG_BRANDKOSEI_ID = ? AND MAM_KOTAIKANRI_FLG = 1 " +
            "GROUP BY MAF_ATSUKAIKIKI_ID";
        return jdbcTemplate.query(sql, (rs, rowNum) ->
            new Object[] { rs.getBigDecimal("MAF_ATSUKAIKIKI_ID"), rs.getInt("DATA_MAX") },
            kikikoseiId, brandkoseiId);
    }

    // =================================================================
    // 9. 契約チェック：取引先契約期間ID取得
    // 【変換元】Mcm0021uScreen.vb - getTkKeiyakuList()
    //   Oracle: TO_DATE(SYSDATE, 'YYYY/MM/DD') → SQL Server: CAST(GETDATE() AS DATE)
    //   Oracle: NVL(TO_DATE(TKB.SYURYO_DT, 'YYYY/MM/DD'), TO_DATE('9999/12/31', ...))
    //         → SQL Server: ISNULL(CAST(TKV.TKB_SYURYO_DT AS DATE), '9999-12-31')
    // =================================================================
    /**
     * 取引先契約期間IDリストを取得する（契約不整合チェック用）
     * <p>指定されたブランド構成+機器構成に紐づく有効な取引先契約期間を検索する。</p>
     */
    public List<BigDecimal> getTkKeiyakuList(BigDecimal brandKoseiId, BigDecimal kikiKoseiId) {
        String sql =
            "SELECT DISTINCT TKV.TKB_TK_KIKAN_ID " +
            "FROM MCM.MCM_TK_KOTAIMEISAI_V TKV " +
            "INNER JOIN MCM.MCM_TK_KEIYAKU TKA ON TKV.TKB_TK_KEIYAKU_ID = TKA.TK_KEIYAKU_ID " +
            "WHERE CAST(GETDATE() AS DATE) < ISNULL(CAST(TKV.TKB_SYURYO_DT AS DATE), '9999-12-31') " +
            "AND TKA.JOTAI = 2 " +
            "AND TKV.TKE_BRANDKOSEI_ID = ? " +
            "AND TKV.TKC_KIKIKOSEI_ID = ?";
        return jdbcTemplate.query(sql, (rs, rowNum) ->
            rs.getBigDecimal("TKB_TK_KIKAN_ID"), brandKoseiId, kikiKoseiId);
    }

    // =================================================================
    // 10. 契約チェック：店舗契約期間ID取得
    // 【変換元】Mcm0021uScreen.vb - getUkKeiyakuList()
    //   Oracle → SQL Server 同上の変換ルール
    // =================================================================
    /**
     * 店舗契約期間IDリストを取得する（契約不整合チェック用）
     * <p>指定されたブランド構成+機器構成に紐づく有効な店舗契約期間を検索する。</p>
     */
    public List<BigDecimal> getUkKeiyakuList(BigDecimal brandKoseiId, BigDecimal kikiKoseiId) {
        String sql =
            "SELECT DISTINCT UKB.UK_KIKAN_ID " +
            "FROM MCM.MCM_UK_KEIYAKU UKA " +
            "INNER JOIN MCM.MCM_UK_KIKAN UKB ON UKA.UK_KEIYAKU_ID = UKB.UK_KEIYAKU_ID " +
            "INNER JOIN MCM.MCM_UK_BRAND UKC ON UKB.UK_KIKAN_ID = UKC.UK_KIKAN_ID " +
            "INNER JOIN MCM.MCM_UK_KIKIKOSEI UKD ON UKC.UK_BRAND_ID = UKD.UK_BRAND_ID " +
            "WHERE CAST(GETDATE() AS DATE) < ISNULL(CAST(UKB.SYURYO_DT AS DATE), '9999-12-31') " +
            "AND UKA.JOTAI = 2 " +
            "AND UKC.BRANDKOSEI_ID = ? " +
            "AND UKD.KIKIKOSEI_ID = ?";
        return jdbcTemplate.query(sql, (rs, rowNum) ->
            rs.getBigDecimal("UK_KIKAN_ID"), brandKoseiId, kikiKoseiId);
    }

    // =================================================================
    // 11. プラント数カウント（契約不整合チェック用）
    // 【変換元】MCM_MA_KIKIKOSEITableAdapter.CountPlantId()
    //   SELECT COUNT(DISTINCT PLANT_ID) FROM MCM_MA_KIKIKOSEI WHERE KIKIKOSEI_ID = :kikikoseiId
    // =================================================================
    /**
     * 指定機器構成IDに紐づくプラント数をカウントする
     * <p>付替え後にプラントが複数となるかチェックするために使用。</p>
     */
    /** #265: 全件移動・一部分割の両方で個体の関連IDを同じSQLで更新する。 */
    public int moveIndividuals(BigDecimal sourceComposition, BigDecimal sourceBrand,
                               BigDecimal destinationBrand, BigDecimal destinationComposition,
                               String user, java.time.LocalDateTime now) {
        return jdbcTemplate.update("UPDATE MCM.MCM_MA_KIKIKOTAIKANRI "
                + "SET BRANDKOSEI_ID = ?, KIKIKOSEI_ID = ?, LASTUPDATE_BY = ?, LASTUPDATE_DT = ? "
                + "WHERE KIKIKOSEI_ID = ? AND BRANDKOSEI_ID = ?",
                destinationBrand, destinationComposition, user, now, sourceComposition, sourceBrand);
    }

    public int countPlantIdByKikikoseiId(BigDecimal kikikoseiId) {
        Integer count = jdbcTemplate.queryForObject(
            "SELECT COUNT(DISTINCT PLANT_ID) FROM MCM.MCM_MA_KIKIKOSEI WHERE KIKIKOSEI_ID = ?",
            Integer.class, kikikoseiId);
        return count != null ? count : 0;
    }

    // =================================================================
    // ユーティリティ
    // =================================================================
    /**
     * 文字列が空またはnullかチェック
     */
    private static boolean isEmpty(String s) {
        return s == null || s.trim().isEmpty();
    }
}
