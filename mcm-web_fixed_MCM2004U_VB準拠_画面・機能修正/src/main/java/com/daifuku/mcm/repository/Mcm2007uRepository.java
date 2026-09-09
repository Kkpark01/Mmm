package com.daifuku.mcm.repository;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Repository;

import com.daifuku.mcm.form.Mcm2007uForm;
import com.daifuku.mcm.form.Mcm2007uForm.KihonBrandRowForm;
import com.daifuku.mcm.form.Mcm2007uForm.KoseiRowForm;
import com.daifuku.mcm.form.Mcm2007uForm.KotaiRowForm;
import com.daifuku.mcm.form.Mcm2007uForm.MeisaiRowForm;
import com.daifuku.mcm.form.Mcm2007uForm.MitsumoriRowForm;

/**
 * 【変換元】Mcm2007uDataSet.Designer.vb / Mcm2007uScreen1.vb / Mcm2007uScreen2.vb
 * MCM2007U ユーザ契約内容変更リポジトリ（SELECT専用）
 *
 * Step1: 店舗見積（MCM_UM_MITSUMORI）+ 基本ブランド（MCM_UM_KIHON_BRAND）一覧取得
 * Step2: 機器構成マスタ（MCM_MA_KIKIKOSEI）+ 機器明細マスタ + 個体管理マスタ取得
 */
@Repository
public class Mcm2007uRepository {

    @Autowired
    private JdbcTemplate jdbc;

    private static final DateTimeFormatter DT_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");

    /** VB getUmMitsumoriId: 当期を優先、存在しなければ最も近い未来の開始日。 */
    public BigDecimal findInitialMitsumoriId(BigDecimal estimateId) {
        var ids = jdbc.queryForList(
            "SELECT TOP (1) G.UM_MITSUMORI_ID FROM MCM.MCM_UM_MITSUMORI G " +
            "WHERE G.UM_KIHON_MITSUMORI_ID = ? AND (" +
            "CAST(GETDATE() AS date) BETWEEN G.KAISI_DT AND ISNULL(G.SYURYO_DT,CONVERT(date,'99991231',112)) " +
            "OR G.KAISI_DT > CAST(GETDATE() AS date)) " +
            "ORDER BY CASE WHEN G.KAISI_DT <= CAST(GETDATE() AS date) THEN 0 ELSE 1 END, G.KAISI_DT, G.UM_MITSUMORI_ID",
            BigDecimal.class, estimateId);
        return ids.isEmpty() ? null : ids.get(0);
    }

    private String fmtDt(Date d) {
        if (d == null) return "";
        return d.toLocalDate().format(DT_FMT);
    }

    // ===================================================================
    // Step1: 店舗見積一覧 (MCM_UM_MITSUMORI + MCM_UM_KIHON_MITSUMORI)
    // 条件: 承認済(SYOUNIN_JOTAI=3) かつ 廃棄でない(JOTAI<>3)
    // 【変換元】Mcm2007uScreen1.vb — DataGrid1バインド
    // ===================================================================

    public List<MitsumoriRowForm> findMitsumoriList(BigDecimal plantId) {
        return jdbc.query(
            "SELECT G.UM_MITSUMORI_ID, G.UM_MITSUMORI_NO, G.KAISI_DT, G.SYURYO_DT, G.JOTAI " +
            "FROM MCM.MCM_UM_MITSUMORI G " +
            "INNER JOIN MCM.MCM_UM_KIHON_MITSUMORI A ON A.UM_KIHON_MITSUMORI_ID = G.UM_KIHON_MITSUMORI_ID " +
            "WHERE A.PLANT_ID = ? " +
            "  AND G.SYOUNIN_JOTAI = '3' " +
            "  AND G.JOTAI <> '3' " +
            "ORDER BY G.KAISI_DT DESC",
            new Object[]{plantId},
            (rs, rn) -> {
                MitsumoriRowForm r = new MitsumoriRowForm();
                r.setUmMitsumoriId(rs.getBigDecimal("UM_MITSUMORI_ID"));
                r.setUmMitsumoriNo(rs.getString("UM_MITSUMORI_NO"));
                r.setKaisiDt(fmtDt(rs.getDate("KAISI_DT")));
                r.setSyuryoDt(fmtDt(rs.getDate("SYURYO_DT")));
                r.setJotai(rs.getString("JOTAI"));
                return r;
            });
    }

    // ===================================================================
    // Step1: 基本ブランド一覧 (MCM_UM_KIHON_BRAND + MCM_MA_BRAND_KOSEI)
    // 指定見積IDに紐づくブランド行をまとめて取得
    // 【変換元】Mcm2007uScreen1.vb — DataGrid2バインド
    // ===================================================================

    public List<KihonBrandRowForm> findKihonBrandList(List<BigDecimal> umMitsumoriIds) {
        if (umMitsumoriIds == null || umMitsumoriIds.isEmpty()) {
            return new ArrayList<>();
        }
        String placeholders = umMitsumoriIds.stream().map(id -> "?").collect(Collectors.joining(","));
        String sql =
            "SELECT B.UM_KIHON_BRAND_ID, G.UM_MITSUMORI_ID, B.BRANDKOSEI_ID, " +
            "       MBC.BRAND_NK, MBC.BRANDSYOSAI_NK, " +
            "       B.KEIYAKUJIKANTAI, B.HOSYUHOHO " +
            "FROM MCM.MCM_UM_KIHON_BRAND B " +
            "INNER JOIN MCM.MCM_UM_MITSUMORI G ON G.UM_KIHON_MITSUMORI_ID = B.UM_KIHON_MITSUMORI_ID " +
            "INNER JOIN MCM.MCM_MA_BRAND_KOSEI MBC ON MBC.BRANDKOSEI_ID = B.BRANDKOSEI_ID " +
            "WHERE G.UM_MITSUMORI_ID IN (" + placeholders + ") " +
            "ORDER BY G.UM_MITSUMORI_ID, MBC.HYOJIJUN";
        return jdbc.query(sql, umMitsumoriIds.toArray(), (rs, rn) -> {
            KihonBrandRowForm r = new KihonBrandRowForm();
            r.setUmKihonBrandId(rs.getBigDecimal("UM_KIHON_BRAND_ID"));
            r.setUmMitsumoriId(rs.getBigDecimal("UM_MITSUMORI_ID"));
            r.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));
            r.setBrandNk(rs.getString("BRAND_NK"));
            r.setBrandsyosaiNk(rs.getString("BRANDSYOSAI_NK"));
            r.setKeiyakujikantai(rs.getString("KEIYAKUJIKANTAI"));
            r.setHosyuhoho(rs.getString("HOSYUHOHO"));
            return r;
        });
    }

    // ===================================================================
    // Step2: 機器構成マスタ (MCM_MA_KIKIKOSEI)
    // 選択済み見積IDに紐づく構成を取得
    // CHECK_FLG: MCM_UM_KIKIKOSEIに存在する場合=true
    // 【変換元】Mcm2007uScreen2.vb DataGridBrand, DataGridKikosei
    // ===================================================================

    public List<KoseiRowForm> findKoseiList(List<BigDecimal> umMitsumoriIds, BigDecimal plantId) {
        if (umMitsumoriIds == null || umMitsumoriIds.isEmpty()) {
            return new ArrayList<>();
        }
        String placeholders = umMitsumoriIds.stream().map(id -> "?").collect(Collectors.joining(","));
        // パラメータ順: サブクエリのIN句, 外側WHERE句のplantId
        List<Object> allParams = new ArrayList<>(umMitsumoriIds);
        allParams.add(plantId);

        String sql2 =
            "SELECT " +
            "  MAE.KIKIKOSEI_ID, MAD.BRANDKOSEI_ID, " +
            "  MAE.KIKIKOSEI_NK, MAE.SET_NM, MAE.TANI, MAE.TEHAISEIBAN, " +
            "  MAE.CONTROLLER_FLG, MAE.HOSYUHOHO, MAD.HYOJIJUN, " +
            "  CASE WHEN SUB.KIKIKOSEI_ID IS NOT NULL THEN 1 ELSE 0 END AS CHECK_FLG, " +
            "  SUB.UM_KIHON_BRAND_ID " +
            "FROM MCM.MCM_MA_KIKIKOSEI MAE " +
            "INNER JOIN MCM.MCM_MA_BRAND_KOSEI MAD ON MAD.BRANDKOSEI_ID = MAE.BRANDKOSEI_ID " +
            "LEFT JOIN ( " +
            "  SELECT DISTINCT UMB.UM_KIHON_BRAND_ID, UMC.KIKIKOSEI_ID " +
            "  FROM MCM.MCM_UM_KIHON_BRAND UMB " +
            "  INNER JOIN MCM.MCM_UM_KIKIKOSEI UMC ON UMC.UM_KIHON_BRAND_ID = UMB.UM_KIHON_BRAND_ID " +
            "  INNER JOIN MCM.MCM_UM_MITSUMORI UMG ON UMG.UM_KIHON_MITSUMORI_ID = UMB.UM_KIHON_MITSUMORI_ID " +
            "  WHERE UMG.UM_MITSUMORI_ID IN (" + placeholders + ") " +
            ") SUB ON SUB.UM_KIHON_BRAND_ID = MAD.UM_KIHON_BRAND_ID " +
            "       AND SUB.KIKIKOSEI_ID = MAE.KIKIKOSEI_ID " +
            "WHERE MAE.PLANT_ID = ? " +
            "ORDER BY MAD.HYOJIJUN, MAE.HYOJIJUN";

        return jdbc.query(sql2, allParams.toArray(), (rs, rn) -> {
            KoseiRowForm r = new KoseiRowForm();
            r.setKikikoseiId(rs.getBigDecimal("KIKIKOSEI_ID"));
            r.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));
            r.setUmKihonBrandId(rs.getBigDecimal("UM_KIHON_BRAND_ID"));
            r.setKikikoseiNk(rs.getString("KIKIKOSEI_NK"));
            r.setSetNm(rs.getString("SET_NM"));
            r.setTani(rs.getString("TANI"));
            r.setTehaiseiban(rs.getString("TEHAISEIBAN"));
            r.setControllerFlg(rs.getString("CONTROLLER_FLG"));
            r.setHosyuhoho(rs.getString("HOSYUHOHO"));
            r.setHyojijun(rs.getInt("HYOJIJUN"));
            r.setCheckFlg(rs.getInt("CHECK_FLG") == 1);
            return r;
        });
    }

    // ===================================================================
    // Step2: 機器明細マスタ (MCM_MA_KIKIMEISAI)
    // 選択済み見積IDに紐づく明細を取得
    // 【変換元】Mcm2007uScreen2.vb DataGridMeisai
    // ===================================================================

    public List<MeisaiRowForm> findMeisaiList(List<BigDecimal> umMitsumoriIds, BigDecimal plantId) {
        if (umMitsumoriIds == null || umMitsumoriIds.isEmpty()) {
            return new ArrayList<>();
        }
        String placeholders = umMitsumoriIds.stream().map(id -> "?").collect(Collectors.joining(","));
        List<Object> params = new ArrayList<>();
        params.addAll(umMitsumoriIds);
        params.add(plantId);

        String sql =
            "SELECT " +
            "  M.KIKIMEISAI_ID, M.KIKIKOSEI_ID, MAD.BRANDKOSEI_ID, " +
            "  M.SEIZOMAKER_ID, M.SEIZOMAKER_NK, M.KIKIHINMEI_NK, M.KIKIKATASHIKI, " +
            "  M.SURYO_NM, M.KEIYAKUNAIYO, M.KEIYAKU_NO, M.SERVICEKEITAI, " +
            "  M.TORIHOSYUJIKAN_ID, M.DAIFUKUHOSYUJIKAN_ID, " +
            "  M.TENKENKAISU, M.TENKENYOBI, M.HOSYUHOHO, " +
            "  MAD.HYOJIJUN AS MAD_HYOJIJUN, M.HYOJIJUN, " +
            "  CASE WHEN SUB.KIKIMEISAI_ID IS NOT NULL THEN 1 ELSE 0 END AS CHECK_FLG " +
            "FROM MCM.MCM_MA_KIKIMEISAI M " +
            "INNER JOIN MCM.MCM_MA_KIKIKOSEI MAE ON MAE.KIKIKOSEI_ID = M.KIKIKOSEI_ID " +
            "INNER JOIN MCM.MCM_MA_BRAND_KOSEI MAD ON MAD.BRANDKOSEI_ID = MAE.BRANDKOSEI_ID " +
            "LEFT JOIN ( " +
            "  SELECT DISTINCT UM.KIKIMEISAI_ID " +
            "  FROM MCM.MCM_UM_KIKIMEISAI UM " +
            "  INNER JOIN MCM.MCM_UM_KIKIKOSEI UMC ON UMC.UM_KIKIKOSEI_ID = UM.UM_KIKIKOSEI_ID " +
            "  INNER JOIN MCM.MCM_UM_KIHON_BRAND UMB ON UMB.UM_KIHON_BRAND_ID = UMC.UM_KIHON_BRAND_ID " +
            "  INNER JOIN MCM.MCM_UM_MITSUMORI UMG ON UMG.UM_KIHON_MITSUMORI_ID = UMB.UM_KIHON_MITSUMORI_ID " +
            "  WHERE UMG.UM_MITSUMORI_ID IN (" + placeholders + ") " +
            ") SUB ON SUB.KIKIMEISAI_ID = M.KIKIMEISAI_ID " +
            "WHERE MAE.PLANT_ID = ? " +
            "ORDER BY MAD.HYOJIJUN, M.HYOJIJUN";

        return jdbc.query(sql, params.toArray(), (rs, rn) -> {
            MeisaiRowForm r = new MeisaiRowForm();
            r.setKikimeisaiId(rs.getBigDecimal("KIKIMEISAI_ID"));
            r.setKikikoseiId(rs.getBigDecimal("KIKIKOSEI_ID"));
            r.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));
            r.setSeizomakerId(rs.getBigDecimal("SEIZOMAKER_ID"));
            r.setSeizomankerNk(rs.getString("SEIZOMAKER_NK"));
            r.setKikihinmeiNk(rs.getString("KIKIHINMEI_NK"));
            r.setKikikatashiki(rs.getString("KIKIKATASHIKI"));
            r.setSuryoNm(rs.getString("SURYO_NM"));
            r.setKeiyakunaiyo(rs.getString("KEIYAKUNAIYO"));
            r.setKeiyakuNo(rs.getString("KEIYAKU_NO"));
            r.setServicekeitai(rs.getString("SERVICEKEITAI"));
            r.setTorihosyujikanId(rs.getBigDecimal("TORIHOSYUJIKAN_ID"));
            r.setDaifukuhosyujikanId(rs.getBigDecimal("DAIFUKUHOSYUJIKAN_ID"));
            r.setTenkenkaisu(rs.getString("TENKENKAISU"));
            r.setTenkenyobi(rs.getString("TENKENYOBI"));
            r.setHosyuhoho(rs.getString("HOSYUHOHO"));
            r.setMaeHyojijun(rs.getInt("MAD_HYOJIJUN"));
            r.setHyojijun(rs.getInt("HYOJIJUN"));
            r.setCheckFlg(rs.getInt("CHECK_FLG") == 1);
            return r;
        });
    }

    // ===================================================================
    // Step2: 個体管理マスタ (MCM_MA_KIKIKOTAIKANRI)
    // 【変換元】Mcm2007uScreen2.vb DataGridKotai
    // ===================================================================

    public List<KotaiRowForm> findKotaiList(List<BigDecimal> umMitsumoriIds, BigDecimal plantId) {
        if (umMitsumoriIds == null || umMitsumoriIds.isEmpty()) {
            return new ArrayList<>();
        }
        String placeholders = umMitsumoriIds.stream().map(id -> "?").collect(Collectors.joining(","));
        List<Object> params = new ArrayList<>();
        params.addAll(umMitsumoriIds);
        params.add(plantId);

        String sql =
            "SELECT " +
            "  KT.KOTAIKANRI_ID, KT.KIKIMEISAI_ID, KT.KIKIKOSEI_ID, MAD.BRANDKOSEI_ID, " +
            "  KT.KOTAI_NK, KT.SERIAL_NO, KT.ITIZINONNYUU_DT, KT.SETCHIBASYO, " +
            "  CASE WHEN SUB.KOTAIKANRI_ID IS NOT NULL THEN 1 ELSE 0 END AS CHECK_FLG " +
            "FROM MCM.MCM_MA_KIKIKOTAIKANRI KT " +
            "INNER JOIN MCM.MCM_MA_KIKIMEISAI M ON M.KIKIMEISAI_ID = KT.KIKIMEISAI_ID " +
            "INNER JOIN MCM.MCM_MA_KIKIKOSEI MAE ON MAE.KIKIKOSEI_ID = M.KIKIKOSEI_ID " +
            "INNER JOIN MCM.MCM_MA_BRAND_KOSEI MAD ON MAD.BRANDKOSEI_ID = MAE.BRANDKOSEI_ID " +
            "LEFT JOIN ( " +
            "  SELECT DISTINCT UMK.KOTAIKANRI_ID " +
            "  FROM MCM.MCM_UM_KOTAIKANRI UMK " +
            "  INNER JOIN MCM.MCM_UM_KIKIMEISAI UM ON UM.UM_KIKIMEISAI_ID = UMK.UM_KIKIMEISAI_ID " +
            "  INNER JOIN MCM.MCM_UM_KIKIKOSEI UMC ON UMC.UM_KIKIKOSEI_ID = UM.UM_KIKIKOSEI_ID " +
            "  INNER JOIN MCM.MCM_UM_KIHON_BRAND UMB ON UMB.UM_KIHON_BRAND_ID = UMC.UM_KIHON_BRAND_ID " +
            "  INNER JOIN MCM.MCM_UM_MITSUMORI UMG ON UMG.UM_KIHON_MITSUMORI_ID = UMB.UM_KIHON_MITSUMORI_ID " +
            "  WHERE UMG.UM_MITSUMORI_ID IN (" + placeholders + ") " +
            ") SUB ON SUB.KOTAIKANRI_ID = KT.KOTAIKANRI_ID " +
            "WHERE MAE.PLANT_ID = ? " +
            "ORDER BY MAD.BRANDKOSEI_ID, KT.KIKIKOSEI_ID, KT.KOTAI_NK";

        return jdbc.query(sql, params.toArray(), (rs, rn) -> {
            KotaiRowForm r = new KotaiRowForm();
            r.setKotaikanriId(rs.getBigDecimal("KOTAIKANRI_ID"));
            r.setKikimeisaiId(rs.getBigDecimal("KIKIMEISAI_ID"));
            r.setKikikoseiId(rs.getBigDecimal("KIKIKOSEI_ID"));
            r.setBrandkoseiId(rs.getBigDecimal("BRANDKOSEI_ID"));
            r.setKotaiNk(rs.getString("KOTAI_NK"));
            r.setSerialNo(rs.getString("SERIAL_NO"));
            r.setItizinonnyuDt(fmtDt(rs.getDate("ITIZINONNYUU_DT")));
            r.setSetchibasyo(rs.getString("SETCHIBASYO"));
            r.setCheckFlg(rs.getInt("CHECK_FLG") == 1);
            return r;
        });
    }

    // ===================================================================
    // 重複個体チェック (既存契約に同じ個体が含まれていないか)
    // 【変換元】Mcm2007uScreen1.vb — NextButton_Click バリデーション
    // ===================================================================

    public List<String> findDuplicateKotai(List<BigDecimal> kotaikanriIds, BigDecimal ukKeiyakuId) {
        if (kotaikanriIds == null || kotaikanriIds.isEmpty()) {
            return Collections.emptyList();
        }
        String placeholders = kotaikanriIds.stream().map(id -> "?").collect(Collectors.joining(","));
        List<Object> params = new ArrayList<>(kotaikanriIds);
        params.add(ukKeiyakuId);

        String sql =
            "SELECT KT.KOTAI_NK " +
            "FROM MCM.MCM_UK_KOTAIMEISAI KO " +
            "INNER JOIN MCM.MCM_MA_KIKIKOTAIKANRI KT ON KT.KOTAIKANRI_ID = KO.KOTAIKANRI_ID " +
            "INNER JOIN MCM.MCM_UK_KIKIMEISAI M ON M.UK_KIKIMEISAI_ID = KO.UK_KIKIMEISAI_ID " +
            "INNER JOIN MCM.MCM_UK_KIKIKOSEI K ON K.UK_KIKIKOSEI_ID = M.UK_KIKIKOSEI_ID " +
            "INNER JOIN MCM.MCM_UK_BRAND B ON B.UK_BRAND_ID = K.UK_BRAND_ID " +
            "INNER JOIN MCM.MCM_UK_KIKAN KK ON KK.UK_KIKAN_ID = B.UK_KIKAN_ID " +
            "WHERE KO.KOTAIKANRI_ID IN (" + placeholders + ") " +
            "  AND KK.UK_KEIYAKU_ID = ?";

        return jdbc.query(sql, params.toArray(),
            (rs, rn) -> rs.getString("KOTAI_NK"));
    }
}
