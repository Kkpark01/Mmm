/**
 * 【変換元】Mcm0014uScreen.vb — SEARCHButton_Click / CellContentClick /
 *           RowDeleteButton_Click / SETUPButton_Click
 *
 * MCM0014U — 納入機器明細選定 サービス
 *
 * 取扱機器の検索（動的SQL）と明細行の参照チェックを実装。
 * @since 2026-06-08
 */
package com.daifuku.mcm.service;

import com.daifuku.mcm.common.Mcm0014uConstants;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.util.*;

@Service
public class Mcm0014uService {

    private static final Logger log = LoggerFactory.getLogger(Mcm0014uService.class);
    private final JdbcTemplate jdbc;

    public Mcm0014uService(JdbcTemplate jdbc) {
        this.jdbc = jdbc;
    }

    // ========================================================================
    // 取扱機器検索（動的SQL）
    // 【変換元】Mcm0014uScreen.vb — SEARCHButton_Click + Mcm0014u_MCM_MA_ATSUKAIKIKI.xml
    //   元SQL: SELECT MAH.CONTROLLER_FLG, MAH.ATSUKAIKIKI_ID, MAJ.SEIZOMAKER_NK, ...
    //          FROM MCM_MA_ATSUKAIKIKI MAH
    //          INNER JOIN MCM_MA_SEIZOMAKER MAJ ON MAH.SEIZOMAKER_ID = MAJ.SEIZOMAKER_ID
    //          INNER JOIN MCM_MA_KIKIBUNRUI MAM ON MAH.KIKIBUNRUI_ID = MAM.KIKIBUNRUI_ID
    //          WHERE (dynamic) AND MAH.YUKO_FLG = '0'
    //          ORDER BY MAH.HYOJIJUN
    //
    // Oracle→SQL Server: || → +, NVL → ISNULL
    // ========================================================================
    public List<Map<String, Object>> searchAtsukaikiki(
            Long seizomakerId, String hinmei, String katashiki,
            Long kikibunruiId, Long torihikisakiId) {
        return searchAtsukaikiki(seizomakerId,hinmei,katashiki,kikibunruiId,torihikisakiId,null);
    }

    /** 選択済みメーカーはID一致、直接入力したメーカー名はVBと同じ前方一致。 */
    public List<Map<String, Object>> searchAtsukaikiki(
            Long seizomakerId, String hinmei, String katashiki,
            Long kikibunruiId, Long torihikisakiId, String makerText) {

        StringBuilder sql = new StringBuilder();
        List<Object> params = new ArrayList<>();

        sql.append("SELECT MAH.CONTROLLER_FLG, MAH.ATSUKAIKIKI_ID, ");
        sql.append("CASE WHEN MAJ.YUKO_FLG=0 THEN '' ELSE N'× ' END + MAJ.SEIZOMAKER_NK AS SEIZOMAKER_NK, MAH.SEIZOMAKER_ID, ");
        sql.append("MAH.ATSUKAIKIKI_NK, MAH.KATASHIKI, ");
        sql.append("CASE WHEN MAM.KOTAIKANRI_FLG=1 THEN N'○ ' ELSE N'　　' END + MAM.KIKIBUNRUI_NK AS KIKIBUNRUI_NK, MAH.HYOJIJUN, ");
        sql.append("MAM.KOTAIKANRI_FLG, MAM.KIKIBUNRUI_CD, ");
        sql.append("MAM.OYAKIKIBUNRUI_CD, MAH.BIKO, ");
        sql.append("CASE WHEN EXISTS (SELECT 1 FROM MCM.MCM_MA_ATSUKAIKIKI_KOSEI X INNER JOIN MCM.MCM_MA_TORIHIKISAKI T ON T.TORIHIKISAKI_ID=X.TORIHIKISAKI_ID WHERE X.ATSUKAIKIKI_ID=MAH.ATSUKAIKIKI_ID AND T.YUKO_FLG=0) THEN 1 ELSE 0 END AS SELECTABLE ");
        sql.append("FROM MCM.MCM_MA_ATSUKAIKIKI MAH ");
        sql.append("INNER JOIN MCM.MCM_MA_SEIZOMAKER MAJ ON MAH.SEIZOMAKER_ID = MAJ.SEIZOMAKER_ID ");
        sql.append("INNER JOIN MCM.MCM_MA_KIKIBUNRUI MAM ON MAH.KIKIBUNRUI_ID = MAM.KIKIBUNRUI_ID ");

        List<String> conds = new ArrayList<>();
        // 有効のみ
        conds.add("MAH.YUKOU_FLG = 0");

        if (seizomakerId != null) {
            conds.add("MAH.SEIZOMAKER_ID = ?");
            params.add(BigDecimal.valueOf(seizomakerId));
        } else if (!isEmpty(makerText)) {
            conds.add("UPPER(MAJ.SEIZOMAKER_NK) LIKE UPPER(?)");
            params.add(makerText.trim()+"%");
        }
        if (!isEmpty(hinmei)) {
            // 【変換元】UPPER(MAH.ATSUKAIKIKI_NK) LIKE UPPER('%' || :hinmei || '%')
            conds.add("UPPER(MAH.ATSUKAIKIKI_NK) LIKE UPPER(?)");
            params.add("%" + hinmei.trim() + "%");
        }
        if (!isEmpty(katashiki)) {
            conds.add("UPPER(MAH.KATASHIKI) LIKE UPPER(?)");
            params.add("%" + katashiki.trim() + "%");
        }
        if (kikibunruiId != null) {
            conds.add("MAH.KIKIBUNRUI_ID = ?");
            params.add(BigDecimal.valueOf(kikibunruiId));
        }
        if (torihikisakiId != null) {
            conds.add("EXISTS (SELECT 1 FROM MCM.MCM_MA_ATSUKAIKIKI_KOSEI MAI "
                     + "INNER JOIN MCM.MCM_MA_TORIHIKISAKI MAK ON MAK.TORIHIKISAKI_ID=MAI.TORIHIKISAKI_ID "
                     + "WHERE MAI.ATSUKAIKIKI_ID=MAH.ATSUKAIKIKI_ID AND MAK.YUKO_FLG=0 AND MAI.TORIHIKISAKI_ID=?)");
            params.add(BigDecimal.valueOf(torihikisakiId));
        }

        if (!conds.isEmpty()) {
            sql.append("WHERE ");
            sql.append(String.join(" AND ", conds));
            sql.append(" ");
        }
        sql.append("ORDER BY MAJ.HYOJIJUN ASC, MAH.HYOJIJUN ASC");

        log.debug("【MCM0014U】取扱機器検索SQL: {}", sql);
        return jdbc.queryForList(sql.toString(), params.toArray());
    }

    // ========================================================================
    // 取扱機器1件取得（選択リンク用）
    // ========================================================================
    public Map<String, Object> getAtsukaikikiById(Long atsukaikikiId) {
        String sql =
            "SELECT MAH.ATSUKAIKIKI_ID, CASE WHEN MAJ.YUKO_FLG=0 THEN '' ELSE N'× ' END + MAJ.SEIZOMAKER_NK AS SEIZOMAKER_NK, MAH.SEIZOMAKER_ID, " +
            "MAH.ATSUKAIKIKI_NK, MAH.KATASHIKI, " +
            "CASE WHEN MAM.KOTAIKANRI_FLG=1 THEN N'○ ' ELSE N'　　' END + MAM.KIKIBUNRUI_NK AS KIKIBUNRUI_NK, MAM.KIKIBUNRUI_CD, MAM.OYAKIKIBUNRUI_CD, " +
            "MAM.KOTAIKANRI_FLG, MAH.CONTROLLER_FLG, MAH.BIKO, " +
            "CASE WHEN EXISTS (SELECT 1 FROM MCM.MCM_MA_ATSUKAIKIKI_KOSEI X INNER JOIN MCM.MCM_MA_TORIHIKISAKI T ON T.TORIHIKISAKI_ID=X.TORIHIKISAKI_ID WHERE X.ATSUKAIKIKI_ID=MAH.ATSUKAIKIKI_ID AND T.YUKO_FLG=0) THEN 1 ELSE 0 END AS SELECTABLE " +
            "FROM MCM.MCM_MA_ATSUKAIKIKI MAH " +
            "INNER JOIN MCM.MCM_MA_SEIZOMAKER MAJ ON MAH.SEIZOMAKER_ID = MAJ.SEIZOMAKER_ID " +
            "INNER JOIN MCM.MCM_MA_KIKIBUNRUI MAM ON MAH.KIKIBUNRUI_ID = MAM.KIKIBUNRUI_ID " +
            "WHERE MAH.ATSUKAIKIKI_ID = ? AND MAH.YUKOU_FLG=0";
        try {
            return jdbc.queryForMap(sql, BigDecimal.valueOf(atsukaikikiId));
        } catch (org.springframework.dao.EmptyResultDataAccessException e) {
            return null;
        }
    }

    // ========================================================================
    // 明細行削除前参照チェック
    // 【変換元】Mcm0014uScreen.vb — BeforeDeleteCheck()
    //   元SQL: SELECT COUNT(*) FROM MCM_TK_KIKIMEISAI WHERE KIKIMEISAI_ID = :id
    //          SELECT COUNT(*) FROM MCM_TM_KIKIMEISAI WHERE KIKIMEISAI_ID = :id
    //          SELECT COUNT(*) FROM MCM_UK_KIKIMEISAI WHERE KIKIMEISAI_ID = :id
    //          SELECT COUNT(*) FROM MCM_UM_KIKIMEISAI WHERE KIKIMEISAI_ID = :id
    // ========================================================================
    public List<String> checkDeleteReference(Long kikimeisaiId) {
        if(kikimeisaiId == null || kikimeisaiId <= 0) return List.of();
        for(String table:List.of("MCM_TK_KIKIMEISAI","MCM_TM_KIKIMEISAI","MCM_UK_KIKIMEISAI","MCM_UM_KIKIMEISAI")) {
            // 照会失敗は上位へ伝え、削除を中断する。存在しないものとして扱わない。
            if(!jdbc.queryForList("SELECT TOP (1) KIKIMEISAI_ID FROM MCM."+table+" WHERE KIKIMEISAI_ID = ?",BigDecimal.valueOf(kikimeisaiId)).isEmpty())
                return List.of("見積または契約情報として、既に使用されている為、削除する事が出来ません。");
        }
        return List.of();
    }

    // ========================================================================
    // コンボボックスデータ
    // 【変換元】Mcm0014uScreen.vb — Search() → ComboBox設定
    // ========================================================================
    public List<Map<String, Object>> getSeizomakerCombo() {
        try {
            return jdbc.queryForList(
                "SELECT SEIZOMAKER_ID, SEIZOMAKER_NK FROM MCM.MCM_MA_SEIZOMAKER " +
                "WHERE YUKO_FLG = 0 ORDER BY HYOJIJUN");
        } catch (Exception e) { return List.of(); }
    }

    public List<Map<String, Object>> getKikibunruiCombo() {
        try {
            return jdbc.queryForList(
                "SELECT KIKIBUNRUI_ID, KIKIBUNRUI_NK FROM MCM.MCM_MA_KIKIBUNRUI " +
                "ORDER BY HYOJIJUN");
        } catch (Exception e) { return List.of(); }
    }

    public List<Map<String, Object>> getTorihikisakiCombo() {
        try {
            return jdbc.queryForList(
                "SELECT TORIHIKISAKI_ID, TORIHIKISAKI_NK FROM MCM.MCM_MA_TORIHIKISAKI " +
                "ORDER BY HYOJIJUN");
        } catch (Exception e) { return List.of(); }
    }

    /** VB の取扱機器 → 取引先の表示用データ。検索済み機器だけを一括取得する。 */
    public List<Map<String, Object>> getAtsukaikikiPartners(
            List<Map<String, Object>> equipment, Long torihikisakiId) {
        List<BigDecimal> ids = equipment.stream().map(row -> row.get("ATSUKAIKIKI_ID"))
            .filter(Objects::nonNull).map(id -> new BigDecimal(id.toString())).distinct().toList();
        List<Map<String, Object>> result = new ArrayList<>();
        // SQL Server のパラメータ数上限を超えないように分割する。
        for (int start = 0; start < ids.size(); start += 1000) {
            List<BigDecimal> batch = ids.subList(start, Math.min(start + 1000, ids.size()));
            List<Object> params = new ArrayList<>(batch);
            String sql = "SELECT MAI.ATSUKAIKIKI_ID, MAK.TORIHIKISAKI_CD, MAK.TORIHIKISAKI_NK, MAK.YUKO_FLG "
                + "FROM MCM.MCM_MA_ATSUKAIKIKI_KOSEI MAI "
                + "INNER JOIN MCM.MCM_MA_ATSUKAIKIKI MAH ON MAH.ATSUKAIKIKI_ID = MAI.ATSUKAIKIKI_ID "
                + "INNER JOIN MCM.MCM_MA_TORIHIKISAKI MAK ON MAI.TORIHIKISAKI_ID = MAK.TORIHIKISAKI_ID "
                + "WHERE MAH.YUKOU_FLG = 0 AND MAI.ATSUKAIKIKI_ID IN (" + String.join(",", Collections.nCopies(batch.size(), "?")) + ") ";
            if (torihikisakiId != null) {
                sql += "AND MAI.TORIHIKISAKI_ID = ? ";
                params.add(BigDecimal.valueOf(torihikisakiId));
            }
            sql += "ORDER BY MAK.YUKO_FLG, MAI.HYOJIJUN ASC";
            result.addAll(jdbc.queryForList(sql, params.toArray()));
        }
        return result;
    }

    private static boolean isEmpty(String s) {
        return s == null || s.trim().isEmpty();
    }
}
