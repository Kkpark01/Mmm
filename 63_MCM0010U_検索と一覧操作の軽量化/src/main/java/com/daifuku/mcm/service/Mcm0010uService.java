/**
 * 【変換元】Mcm0010uScreen.vb
 * 【説明】納入先マスタ検索画面のサービスクラス
 * 元メソッド対応:
 *   SEARCHButton_Click → search()
 *   RowDeleteButton_Click → deletePlant()
 *   DataRelation連動 → getPlantsByNonyusakiId(), getBrandsByPlantId(), getKikiByPlantId()
 */
package com.daifuku.mcm.service;

import com.daifuku.mcm.common.AppConstants;
import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.dto.Mcm0010uBrandKoseiDto;
import com.daifuku.mcm.dto.Mcm0010uKikiKoseiDto;
import com.daifuku.mcm.entity.NonyusakiEntity;
import com.daifuku.mcm.entity.PlantEntity;
import com.daifuku.mcm.form.Mcm0010uForm;
import com.daifuku.mcm.repository.NonyusakiRepository;
import com.daifuku.mcm.repository.PlantRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.HashSet;
import java.util.Set;
import java.util.Collections;

@Service
public class Mcm0010uService {

    @Autowired
    private NonyusakiRepository nonyusakiRepository;

    @Autowired
    private PlantRepository plantRepository;

    @Autowired
    private JdbcTemplate jdbcTemplate;

    /**
     * 【移植】担当者の実効権限区分（"UPDATE"/"INSPECTION"）を取得する。
     * 元VB: CPCoreUserControl.GetAuthorityDivision() → CPUserInfo.GetAuthorityDivision(functionId)
     *   McmLogin_AuthorityInfo.xml のSQL（MCM_MO_TANTOKENGEN×MCM_MO_KENGENKOSEI×
     *   MCM_MO_KENGENBUNRUI×MCM_MO_KINO）を機能ID指定で実行するのと同じロジック。
     *   「納入機器関連」（KENGENBUNRUI_ID=1）はMCM0010U/MCM0021U等のKINO_IDに
     *   紐づくため、機能IDにMCM0010Uを渡せば当該カテゴリの実効権限が得られる。
     *
     * @param loginId    ログインID
     * @param functionId 機能ID（画面ID）
     * @return "UPDATE"（作成権限あり） / "INSPECTION"（参照のみ） / "NONE"（権限なし）
     */
    public String findAuthorityDivision(String loginId, String functionId) {
        if (loginId == null) {
            return AppConstants.AUTHORITY_DIVISION_INSPECTION;
        }
        String sql = "SELECT MAX(C.RIYOKENGEN_KBN) FROM MCM.MCM_MO_TANTO A " +
                "JOIN MCM.MCM_MO_TANTOKENGEN B ON B.TANTO_ID=A.TANTO_ID " +
                "JOIN MCM.MCM_MO_KENGENKOSEI C ON C.KENGENBUNRUI_ID=B.KENGENBUNRUI_ID AND C.RIYOKENGEN_KBN=B.RIYOKENGEN_KBN " +
                "JOIN MCM.MCM_MO_KENGENBUNRUI D ON D.KENGENBUNRUI_ID=C.KENGENBUNRUI_ID " +
                "JOIN MCM.MCM_MO_KINO E ON E.KINO_ID=C.KINO_ID WHERE A.LOGIN_ID=? AND E.KINO_ID=?";
        try {
            String kbn = jdbcTemplate.queryForObject(sql, String.class, loginId, functionId);
            if (McmConstants.RIYOKENGEN_KBN_SAKUSEI_CD.equals(kbn)) {
                return AppConstants.AUTHORITY_DIVISION_UPDATE;
            } else if (McmConstants.RIYOKENGEN_KBN_SANSYO_CD.equals(kbn)) {
                return AppConstants.AUTHORITY_DIVISION_INSPECTION;
            }
            return AppConstants.AUTHORITY_DIVISION_NONE;
        } catch (org.springframework.dao.EmptyResultDataAccessException ex) {
            return AppConstants.AUTHORITY_DIVISION_NONE;
        }
    }

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
        String cd = toSearchParam(form.getNonyusakiCd());
        String nk = nvl(form.getNonyusakiNk());
        String jusyo = nvl(form.getJusyo());
        String sid = toSearchParam(form.getSupportId());
        String pnk = toSearchParam(form.getPlantNk());

        List<NonyusakiEntity> list = nonyusakiRepository.searchNonyusaki(cd, nk, jusyo, sid, pnk);
        // 計算列 PLANTTSUIKA を付与
        for (NonyusakiEntity e : list) {
            e.setPlanttsuika("プラント追加");
        }
        return list;
    }

    /**
     * プラントマスタ検索
     * 元VB: Me.Fill(Me.Mcm0010uDataSet.MCM_MA_PLANT, ...)
     * @param authorityDivision 権限区分 ("INSPECTION" or "UPDATE")
     */
    public List<PlantEntity> searchPlants(Mcm0010uForm form, String authorityDivision) {
        String cd = toSearchParam(form.getNonyusakiCd());
        String nk = nvl(form.getNonyusakiNk());
        String jusyo = nvl(form.getJusyo());
        String sid = toSearchParam(form.getSupportId());
        String pnk = toSearchParam(form.getPlantNk());

        List<PlantEntity> list = plantRepository.searchPlants(cd, nk, jusyo, sid, pnk);
        // VBのmcm_fn_kengenLinkと同じ条件。作成権限は全件、それ以外は機器構成のあるプラントだけ。
        // 1プラントごとの関数呼出しをやめ、必要なIDをまとめて取得する。
        Set<BigDecimal> plantIdsWithKosei = new HashSet<>();
        boolean canCreate = AppConstants.AUTHORITY_DIVISION_UPDATE.equals(authorityDivision);
        if (!canCreate && !list.isEmpty()) {
            List<BigDecimal> ids = list.stream().map(PlantEntity::getPlantId).filter(id -> id != null)
                    .map(BigDecimal::stripTrailingZeros).distinct().toList();
            final int chunkSize = 1000;
            for (int from = 0; from < ids.size(); from += chunkSize) {
                List<BigDecimal> chunk = ids.subList(from, Math.min(from + chunkSize, ids.size()));
                String sql = "SELECT DISTINCT PLANT_ID FROM MCM.MCM_MA_KIKIKOSEI WHERE PLANT_ID IN ("
                        + String.join(",", Collections.nCopies(chunk.size(), "?")) + ")";
                jdbcTemplate.queryForList(sql, BigDecimal.class, chunk.toArray()).forEach(
                        id -> plantIdsWithKosei.add(id.stripTrailingZeros()));
            }
        }
        for (PlantEntity e : list) {
            boolean hasKosei = e.getPlantId() != null
                    && plantIdsWithKosei.contains(e.getPlantId().stripTrailingZeros());
            e.setKikikoseilink(canCreate || hasKosei ? "機器構成" : "");
        }
        return list;
    }

    /**
     * 納入先IDでプラントを取得（リレーション連動用）
     * 元VB: DataRelation MCM_MA_PLANT_MCM_MA_NONYUSAKI
     */
    public List<PlantEntity> getPlantsByNonyusakiId(BigDecimal nonyusakiId) {
        return plantRepository.findByNonyusakiIdOrderByPlantIdAsc(nonyusakiId);
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
            dto.setSuryo(toLong(row.get("SURYO")));
            result.add(dto);
        }
        return result;
    }

    /**
     * プラント行削除
     * 元VB: McmDBUtility.DeletePlant(plantId)
     * 関連テーブルの削除含む
     */
    @Transactional
    public void deletePlant(BigDecimal plantId) {
        // ブランド構成を削除
        jdbcTemplate.update(
            "DELETE FROM MCM.MCM_MA_BRAND_KOSEI WHERE PLANT_ID IN " +
            "(SELECT KIKIKOSEI_ID FROM MCM.MCM_MA_KIKIKOSEI WHERE PLANT_ID = ?)", plantId);
        // 機器明細を削除
        jdbcTemplate.update(
            "DELETE FROM MCM.MCM_MA_KIKIMEISAI WHERE KIKIKOSEI_ID IN " +
            "(SELECT KIKIKOSEI_ID FROM MCM.MCM_MA_KIKIKOSEI WHERE PLANT_ID = ?)", plantId);
        // 機器構成を削除
        jdbcTemplate.update("DELETE FROM MCM.MCM_MA_KIKIKOSEI WHERE PLANT_ID = ?", plantId);
        // プラントを削除
        jdbcTemplate.update("DELETE FROM MCM.MCM_MA_PLANT WHERE PLANT_ID = ?", plantId);
    }

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
    private void appendExistsConditions(StringBuilder sql, List<Object> params,
                                         Mcm0010uForm form, String targetAlias) {
        // targetAlias は MAD (ブランド構成) または MAE (機器構成)
        if (!isBlank(form.getNonyusakiCd())) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_PLANT MAB WHERE MAB.PLANT_ID = ")
               .append(targetAlias).append(".PLANT_ID ")
               .append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ")
               .append("WHERE MAA.NONYUSAKI_ID = MAB.NONYUSAKI_ID ")
               .append("AND MAA.NONYUSAKI_CD LIKE ?)) ");
            params.add(toSearchParam(form.getNonyusakiCd()));
        }
        if (!isBlank(form.getNonyusakiNk())) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_PLANT MAB WHERE MAB.PLANT_ID = ")
               .append(targetAlias).append(".PLANT_ID ")
               .append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ")
               .append("WHERE MAA.NONYUSAKI_ID = MAB.NONYUSAKI_ID ")
               .append("AND UPPER(ISNULL(MAA.NONYUSAKI_NK,'') + ' ' + ISNULL(MAA.KYUNONYUSAKI_NK,'') + ' ' + ")
               .append("ISNULL(MAA.NONYUSAKIKOJO_NK,'') + ' ' + ISNULL(MAA.NONYUSAKIKANA_KN,'') + ' ' + ")
               .append("ISNULL(MAA.NONYUSAKIEIMEI_EN,'')) LIKE '%' + UPPER(?) + '%')) ");
            params.add(form.getNonyusakiNk());
        }
        if (!isBlank(form.getJusyo())) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_PLANT MAB WHERE MAB.PLANT_ID = ")
               .append(targetAlias).append(".PLANT_ID ")
               .append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_NONYUSAKI MAA ")
               .append("WHERE MAA.NONYUSAKI_ID = MAB.NONYUSAKI_ID ")
               .append("AND UPPER(ISNULL(MAA.YUBIN_NO,'') + ' ' + ISNULL(MAA.JUSYO1_NK,'') + ' ' + ")
               .append("ISNULL(MAA.JUSYO2_NK,'') + ' ' + ISNULL(MAA.KUNI_NK,'') + ' ' + ")
               .append("ISNULL(MAA.TEL_NO,'') + ' ' + ISNULL(MAA.FAX_NO,'') + ' ' + ")
               .append("ISNULL(MAA.BIKO,'')) LIKE '%' + UPPER(?) + '%')) ");
            params.add(form.getJusyo());
        }
        if (!isBlank(form.getSupportId())) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_PLANT MAB WHERE MAB.PLANT_ID = ")
               .append(targetAlias).append(".PLANT_ID ")
               .append("AND MAB.SUPPORT_ID LIKE ?) ");
            params.add(toSearchParam(form.getSupportId()));
        }
        if (!isBlank(form.getPlantNk())) {
            sql.append("AND EXISTS (SELECT 1 FROM MCM.MCM_MA_PLANT MAB WHERE MAB.PLANT_ID = ")
               .append(targetAlias).append(".PLANT_ID ")
               .append("AND UPPER(MAB.PLANT_NK) LIKE UPPER(?)) ");
            params.add(toSearchParam(form.getPlantNk()));
        }
    }

    // ===== ユーティリティ =====
    private boolean isBlank(String s) {
        return s == null || s.trim().isEmpty();
    }

    private String nvl(String s) {
        return s == null ? "" : s.trim();
    }

    /** LIKE検索用にワイルドカード付きに変換（空なら空文字を返す） */
    private String toSearchParam(String s) {
        if (isBlank(s)) return "";
        String trimmed = s.trim();
        if (!trimmed.contains("%")) {
            return "%" + trimmed + "%";
        }
        return trimmed;
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
        return Long.valueOf(o.toString());
    }
}
