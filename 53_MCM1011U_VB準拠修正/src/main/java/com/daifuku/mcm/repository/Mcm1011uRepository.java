package com.daifuku.mcm.repository;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Repository;

import com.daifuku.mcm.form.Mcm1011uForm;

/**
 * 【変換元】Mcm1011uDataSet.xsd / Mcm1011u*.xml
 * MCM1011U 審査・承認（取引先契約）リポジトリ
 *
 * 正式VBの期間＋添付履歴単位で検索。承認済は秒単位の更新日時、未承認はPENDINGで区分。
 */
@Repository
public class Mcm1011uRepository {

    @Autowired
    private JdbcTemplate jdbc;

    private static final DateTimeFormatter DT_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");

    // ===================================================================
    // 申請一覧検索
    // 【変換元】Mcm1011uDataSet — TKGSelectAll相当
    //   承認状態 <> '0'（作成中除外）を基本条件とし、
    //   各検索条件を動的に追加。
    //   MCM_TK_KIKAN は全期間を添付履歴と結合。
    // ===================================================================

    public List<Mcm1011uForm.KeiyakuRowForm> search(
            String nonyusakiCd, String nonyusakiNk,
            String supportId, String plantNk,
            List<String> shoninjotaiList) {
        if (shoninjotaiList == null || shoninjotaiList.isEmpty()) return List.of();

        StringBuilder sql = new StringBuilder(
            "SELECT TKA.TK_KEIYAKU_ID, COALESCE(NULLIF(TKA.KEIYAKU_NO,''),'------') AS KEIYAKU_NO, TKG.SHONINJOTAI, TKG.HISTORY_KEY, TKA.SHONINJOTAI AS CONTRACT_STATE, TKA.JOTAI, " +
            "       TKA.IRAITANTOSYA, TKA.LASTUPDATE_DT, TKA.SHINSA_DT, TKA.SHINSA_BY, TKA.SHONIN_DT, TKA.SHONIN_BY, " +
            "       TKB.TK_KIKAN_ID, TKB.LASTUPDATE_DT AS PERIOD_VERSION, TKB.KAISI_DT, TKB.KEIYAKUJIKANTAI, " +
            "       TKB.NONYUSAKI_CD, TKB.NONYUSAKI_NK, MAA.KYUNONYUSAKI_NK, MAA.NONYUSAKIKOJO_NK, TKB.SUPPORT_ID, TKB.PLANT_NK, " +
            "       TKB.TORIHIKISAKI_ID, TKB.TORIHIKISAKI_NK, TKB.SIKIRIGOKEI_KIN " +
            "FROM MCM.MCM_TK_KEIYAKU TKA " +
            "JOIN MCM.MCM_TK_KIKAN TKB ON TKB.TK_KEIYAKU_ID=TKA.TK_KEIYAKU_ID " +
            "JOIN (SELECT TK_KIKAN_ID, " + historySql() + " AS HISTORY_KEY, MAX(SHONINJOTAI) AS SHONINJOTAI " +
            "FROM MCM.MCM_TK_TENPU GROUP BY TK_KIKAN_ID, " + historySql() + ") TKG " +
            "ON TKG.TK_KIKAN_ID=TKB.TK_KIKAN_ID " +
            "JOIN MCM.MCM_MA_NONYUSAKI MAA ON MAA.NONYUSAKI_ID=TKB.NONYUSAKI_ID " +
            "WHERE TKG.SHONINJOTAI <> '0' "
        );

        List<Object> params = new ArrayList<>();

        if (shoninjotaiList != null && !shoninjotaiList.isEmpty()) {
            sql.append("AND TKG.SHONINJOTAI IN (");
            for (int i = 0; i < shoninjotaiList.size(); i++) {
                if (i > 0) sql.append(",");
                sql.append("?");
                params.add(shoninjotaiList.get(i));
            }
            sql.append(") ");
        }

        if (nonyusakiCd != null && !nonyusakiCd.isBlank()) {
            sql.append("AND TKB.NONYUSAKI_CD LIKE ? ");
            params.add(nonyusakiCd.trim() + "%");
        }
        if (nonyusakiNk != null && !nonyusakiNk.isBlank()) {
            sql.append("AND CHARINDEX(?, MAA.NONYUSAKI_NK) > 0 ");
            params.add(nonyusakiNk.trim());
        }
        if (supportId != null && !supportId.isBlank()) {
            sql.append("AND CHARINDEX(?, TKB.SUPPORT_ID) > 0 ");
            params.add(supportId.trim());
        }
        if (plantNk != null && !plantNk.isBlank()) {
            sql.append("AND CHARINDEX(?, TKB.PLANT_NK) > 0 ");
            params.add(plantNk.trim());
        }

        sql.append("ORDER BY TKA.LASTUPDATE_DT ASC, TKA.TK_KEIYAKU_ID, TKB.TK_KIKAN_ID, TKG.HISTORY_KEY");

        return jdbc.query(sql.toString(), params.toArray(), (rs, rowNum) -> {
            Mcm1011uForm.KeiyakuRowForm row = new Mcm1011uForm.KeiyakuRowForm();
            row.setTkKeiyakuId(rs.getBigDecimal("TK_KEIYAKU_ID"));
            row.setTkKikanId(rs.getBigDecimal("TK_KIKAN_ID"));
            row.setKeiyakuNo(rs.getString("KEIYAKU_NO"));
            row.setShoninjotai(rs.getString("SHONINJOTAI"));
            row.setHistoryKey(rs.getString("HISTORY_KEY"));
            row.setContractState(rs.getString("CONTRACT_STATE"));
            row.setPeriodVersion(rs.getObject("PERIOD_VERSION", LocalDateTime.class));
            row.setJotai(rs.getString("JOTAI"));
            row.setLastupdateDt(rs.getObject("LASTUPDATE_DT", LocalDateTime.class));
            row.setIraitantosya(rs.getString("IRAITANTOSYA"));
            LocalDateTime shinsaDt = rs.getObject("SHINSA_DT", LocalDateTime.class);
            row.setShinsaDt(shinsaDt != null ? DT_FMT.format(shinsaDt) : "");
            row.setShinsaBy(rs.getString("SHINSA_BY"));
            LocalDateTime shoninDt = rs.getObject("SHONIN_DT", LocalDateTime.class);
            row.setShoninDt(shoninDt != null ? DT_FMT.format(shoninDt) : "");
            row.setShoninBy(rs.getString("SHONIN_BY"));
            LocalDateTime kaisiDt = rs.getObject("KAISI_DT", LocalDateTime.class);
            row.setKaisiDt(kaisiDt != null ? DT_FMT.format(kaisiDt) : "");
            row.setKeiyakujikantai(rs.getString("KEIYAKUJIKANTAI"));
            row.setNonyusakiCd(rs.getString("NONYUSAKI_CD"));
            row.setNonyusakiNk(rs.getString("NONYUSAKI_NK"));
            row.setOldNonyusakiNk(rs.getString("KYUNONYUSAKI_NK"));
            row.setNonyusakiKojoNk(rs.getString("NONYUSAKIKOJO_NK"));
            row.setSupportId(rs.getString("SUPPORT_ID"));
            row.setPlantNk(rs.getString("PLANT_NK"));
            row.setTorihikisakiId(rs.getBigDecimal("TORIHIKISAKI_ID"));
            row.setTorihikisakiNk(rs.getString("TORIHIKISAKI_NK"));
            row.setKingaku(rs.getBigDecimal("SIKIRIGOKEI_KIN"));
            return row;
        });
    }

    // The pending key deliberately cannot collide with an approved timestamp (VB used the search time).
    private static String historySql() {
        return "CASE WHEN SHONINJOTAI='3' THEN COALESCE(CONVERT(varchar(19), LASTUPDATE_DT, 126),'UNKNOWN') ELSE 'PENDING' END";
    }

    public List<Mcm1011uForm.TenpuRowForm> findTenpuByKikanId(BigDecimal id) {
        return attachments(id, false);
    }

    private List<Mcm1011uForm.TenpuRowForm> attachments(BigDecimal id, boolean lock) {
        return jdbc.query("SELECT TK_TENPU_ID, TENPUFILE_NK, DIRECTORY, SHONINJOTAI, LASTUPDATE_DT, "
                + historySql() + " AS HISTORY_KEY FROM MCM.MCM_TK_TENPU "
                + (lock ? "WITH (UPDLOCK, HOLDLOCK) " : "") + "WHERE TK_KIKAN_ID=? ORDER BY TK_TENPU_ID",
            (rs, n) -> {
                var row = new Mcm1011uForm.TenpuRowForm();
                row.setTkTenpuId(rs.getBigDecimal("TK_TENPU_ID")); row.setTenpufileNk(rs.getString("TENPUFILE_NK"));
                row.setDirectory(rs.getString("DIRECTORY")); row.setShoninjotai(rs.getString("SHONINJOTAI"));
                row.setLastupdateDt(rs.getObject("LASTUPDATE_DT", LocalDateTime.class));
                row.setHistoryKey(rs.getString("HISTORY_KEY")); return row;
            }, id);
    }

    // ===================================================================
    // 承認処理: 審査中(1) → 承認中(2)
    // 【変換元】Mcm1011uScreen.vb ShoninButton_Click → ShinsachuShonin()
    // ===================================================================


    /** Lock and re-read the current contract/period before changing approval state. */
    public Mcm1011uForm.KeiyakuRowForm lockCurrent(BigDecimal contractId, BigDecimal periodId) {
        var rows = jdbc.query("SELECT c.TK_KEIYAKU_ID, c.SHONINJOTAI, c.JOTAI, c.LASTUPDATE_DT, "
                + "k.TK_KIKAN_ID, k.SIKIRIGOKEI_KIN, k.LASTUPDATE_DT AS PERIOD_VERSION "
                + "FROM MCM.MCM_TK_KEIYAKU c WITH (UPDLOCK, HOLDLOCK) "
                + "JOIN MCM.MCM_TK_KIKAN k WITH (UPDLOCK, HOLDLOCK) ON k.TK_KEIYAKU_ID=c.TK_KEIYAKU_ID "
                + "WHERE c.TK_KEIYAKU_ID=? AND k.TK_KIKAN_ID=?",
                (rs, n) -> {
                    var row = new Mcm1011uForm.KeiyakuRowForm();
                    row.setTkKeiyakuId(rs.getBigDecimal("TK_KEIYAKU_ID"));
                    row.setTkKikanId(rs.getBigDecimal("TK_KIKAN_ID"));
                    row.setContractState(rs.getString("SHONINJOTAI")); row.setJotai(rs.getString("JOTAI"));
                    row.setKingaku(rs.getBigDecimal("SIKIRIGOKEI_KIN"));
                    row.setLastupdateDt(rs.getObject("LASTUPDATE_DT", LocalDateTime.class));
                    row.setPeriodVersion(rs.getObject("PERIOD_VERSION", LocalDateTime.class)); return row;
                }, contractId, periodId);
        if (rows.isEmpty()) return null;
        var row = rows.get(0);
        row.setHistoryKey("PENDING");
        row.setPendingAttachments(attachments(periodId, true).stream().filter(t -> !"3".equals(t.getShoninjotai())).toList());
        row.setShoninjotai(row.getPendingAttachments().stream().map(Mcm1011uForm.TenpuRowForm::getShoninjotai)
                .filter(java.util.Objects::nonNull).max(String::compareTo).orElse("0"));
        return row;
    }

    public void advanceContract(BigDecimal id, String before, String user, String actor) {
        approveContract(id, before, "1".equals(before) ? "2" : "3", "1".equals(before) ? "SHINSA" : "SHONIN", user, actor);
    }

    public void advancePeriod(BigDecimal contractId, BigDecimal periodId, String before, String user, String actor) {
        if ("2".equals(before)) activatePeriod(contractId, periodId, user);
        approveAttachments(periodId, user, actor);
    }

    private void approveContract(BigDecimal id, String before, String after, String audit, String user, String actor) {
        int changed = jdbc.update("UPDATE MCM.MCM_TK_KEIYAKU SET SHONINJOTAI=?, " + audit
                + "_DT=GETDATE(), " + audit + "_BY=?, LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? "
                + "WHERE TK_KEIYAKU_ID=? AND SHONINJOTAI=? AND (JOTAI IS NULL OR JOTAI NOT IN ('3','4'))", after, actor, user, id, before);
        if (changed != 1) throw new com.daifuku.mcm.exception.McmBusinessException("申請が更新されています。再検索してください。");
    }

    /** VB: each pending attachment advances one stage; approved/history files stay unchanged. */
    private void approveAttachments(BigDecimal id, String user, String actor) {
        jdbc.update("UPDATE MCM.MCM_TK_TENPU SET "
                + "SHINSA_DT=CASE WHEN SHONINJOTAI='1' THEN GETDATE() ELSE SHINSA_DT END, "
                + "SHINSA_BY=CASE WHEN SHONINJOTAI='1' THEN ? ELSE SHINSA_BY END, "
                + "SHONIN_DT=CASE WHEN SHONINJOTAI='2' THEN GETDATE() ELSE SHONIN_DT END, "
                + "SHONIN_BY=CASE WHEN SHONINJOTAI='2' THEN ? ELSE SHONIN_BY END, "
                + "SHONINJOTAI=CASE WHEN SHONINJOTAI='1' THEN '2' ELSE '3' END, "
                + "LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? WHERE TK_KIKAN_ID=? AND SHONINJOTAI IN ('1','2')",
                actor, actor, user, id);
    }

    private void activatePeriod(BigDecimal contractId, BigDecimal periodId, String user) {
        int changed = jdbc.update("UPDATE MCM.MCM_TK_KIKAN SET YUKO_FLG=0, LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? "
                + "WHERE TK_KEIYAKU_ID=? AND TK_KIKAN_ID=?", user, contractId, periodId);
        if (changed != 1) throw new com.daifuku.mcm.exception.McmBusinessException("対象期間が更新されています。再検索してください。");
    }

    public void rejectContract(BigDecimal contractId, String user) {
        int changed = jdbc.update("UPDATE MCM.MCM_TK_KEIYAKU SET SHONINJOTAI='4', LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? "
                + "WHERE TK_KEIYAKU_ID=? AND SHONINJOTAI IN ('1','2')", user, contractId);
        if (changed != 1) throw new com.daifuku.mcm.exception.McmBusinessException("申請が更新されています。再検索してください。");
    }

    public void rejectPeriod(BigDecimal contractId, BigDecimal periodId, String user) {
        activatePeriod(contractId, periodId, user);
        jdbc.update("UPDATE MCM.MCM_TK_TENPU SET SHONINJOTAI='4', LASTUPDATE_DT=GETDATE(), LASTUPDATE_BY=? "
                + "WHERE TK_KIKAN_ID=? AND SHONINJOTAI<>'3'", user, periodId);
    }
}