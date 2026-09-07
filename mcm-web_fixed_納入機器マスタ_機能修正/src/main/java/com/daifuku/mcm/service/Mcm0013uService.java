/**
 * 【変換元】Mcm0013uScreen.vb — Search() / InsertButton_Click() /
 *           RowDeleteButton1_Click() / RowDeleteButton2_Click() /
 *           MCMMAKOSEIROWINSERTBUTTON_Click() / RowAddSetNum()
 *
 * MCM0013U — 機器構成マスタ登録 サービス
 *
 * 3テーブル（KIKIKOSEI / KIKIMEISAI / KIKIKOTAIKANRI）の一括更新を
 * @Transactional で管理する。
 *
 * JOINが必要な検索は JdbcTemplate、単純CRUDは JPA Repository を使用。
 *
 * ★ セッション管理方式（方針A）対応:
 *   - 仮ID（負数）の構成行をDB永続化時に正式ID採番
 *   - pendingDeliveryのkikikoseiIdを正式IDに自動差替え
 *
 * @since 2026-06-08
 * @modified 2026-06-16 セッション管理方式対応
 */
package com.daifuku.mcm.service;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.daifuku.mcm.common.Mcm0013uConstants;
import com.daifuku.mcm.dto.Mcm0014uDeliveryDto;
import com.daifuku.mcm.entity.KikikoseiEntity;
import com.daifuku.mcm.entity.KikikotaikanriEntity;
import com.daifuku.mcm.entity.KikimeisaiEntity;
import com.daifuku.mcm.form.Mcm0013uForm;
import com.daifuku.mcm.form.Mcm0013uForm.KoseiRowForm;
import com.daifuku.mcm.form.Mcm0013uForm.KotaikanriRowForm;
import com.daifuku.mcm.repository.KikikoseiRepository;
import com.daifuku.mcm.repository.KikikotaikanriRepository;
import com.daifuku.mcm.repository.KikimeisaiRepository;

@Service
public class Mcm0013uService {

    private static final Logger log = LoggerFactory.getLogger(Mcm0013uService.class);
    private static final DateTimeFormatter DT_FMT = DateTimeFormatter.ofPattern("uuuu/MM/dd").withResolverStyle(java.time.format.ResolverStyle.STRICT);

    private final JdbcTemplate jdbc;
    private final KikikoseiRepository koseiRepo;
    private final KikimeisaiRepository meisaiRepo;
    private final KikikotaikanriRepository kotaiRepo;

    public Mcm0013uService(JdbcTemplate jdbc,
    					   KikikoseiRepository koseiRepo,
                           KikimeisaiRepository meisaiRepo,
                           KikikotaikanriRepository kotaiRepo) {
        this.jdbc = jdbc;
        this.koseiRepo = koseiRepo;
        this.meisaiRepo = meisaiRepo;
        this.kotaiRepo = kotaiRepo;
    }

    // ========================================================================
    // ヘッダ情報取得
    // 【変換元】Mcm0013uScreen.vb — Search()
    //   元コード: NonyusakiCdLabel.Text = plant.NonyusakiCd
    //             NonyusakiNkLabel.Text = nonyusaki.NonyusakiNk
    //             SupportIdLabel.Text   = plant.SupportId
    //             PlantNkLabel.Text     = plant.PlantNk
    // ========================================================================
    public Map<String, String> getHeaderInfo(Long plantId) {
        String sql =
            "SELECT p.SUPPORT_ID, p.PLANT_NK, n.NONYUSAKI_CD, n.NONYUSAKI_NK " +
            "FROM MCM.MCM_MA_PLANT p " +
            "INNER JOIN MCM.MCM_MA_NONYUSAKI n ON p.NONYUSAKI_ID = n.NONYUSAKI_ID " +
            "WHERE p.PLANT_ID = ?";
        try {
            return jdbc.queryForObject(sql, new Object[]{plantId}, (rs, i) -> {
                Map<String, String> m = new HashMap<>();
                m.put("supportId",   rs.getString("SUPPORT_ID"));
                m.put("plantNk",     rs.getString("PLANT_NK"));
                m.put("nonyusakiCd", rs.getString("NONYUSAKI_CD"));
                m.put("nonyusakiNk", rs.getString("NONYUSAKI_NK"));
                return m;
            });
        } catch (Exception e) {
            log.warn("ヘッダ情報取得失敗: plantId={}", plantId, e);
            return Map.of();
        }
    }

    // ========================================================================
    // DataGridView1: 機器構成パターン検索
    // 【変換元】Mcm0013uScreen.vb — Search() → Fill MCM_MA_KIKIKOSEI
    //   元SQL(DataSet): SELECT * FROM MCM_MA_KIKIKOSEI WHERE PLANT_ID = :plantId
    //                   ORDER BY HYOJIJUN
    // ========================================================================
    public List<Map<String, Object>> findKoseiByPlantId(BigDecimal plantId) {
        String sql =
            "SELECT KIKIKOSEI_ID, PLANT_ID, KIKIKOSEI_NK, SET_NM, TANI, " +
            "       TEHAISEIBAN, HYOJIJUN, BIKO, CONTROLLER_FLG, " +
            "       CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY " +
            "FROM MCM.MCM_MA_KIKIKOSEI " +
            "WHERE PLANT_ID = ? " +
            "ORDER BY HYOJIJUN ASC";
        return jdbc.queryForList(sql, plantId);
    }

    // ========================================================================
    // DataGridView2: 構成明細検索（ReadOnly表示 + JOIN）
    // 【変換元】Mcm0013uScreen.vb — Search() → Fill MCM_MA_KIKIMEISAI
    //   元SQL(DataSet): SELECT m.*, mk.SEIZOMAKER_NK, kb.KIKIBUNRUI_NK
    //                   FROM MCM_MA_KIKIMEISAI m
    //                   LEFT JOIN MCM_MA_SEIZOMAKER mk ON ...
    //                   LEFT JOIN MCM_MA_KIKIBUNRUI kb ON ...
    //                   WHERE m.KIKIKOSEI_ID = :kikikoseiId
    //                   ORDER BY m.HYOJIJUN
    // ========================================================================
    public List<Map<String, Object>> findMeisaiByKikikoseiId(BigDecimal kikikoseiId) {
        String sql =
        		"SELECT " +
				"    m.KIKIMEISAI_ID, " +
				"    m.KIKIKOSEI_ID, " +
				"    m.KIKIHINMEI_NK, " +
				"    m.KIKIKATASHIKI, " +
				"    m.SURYO_NM, " +
				"    m.ATSUKAIKIKI_ID, " +
				"    m.HYOJIJUN, " +
				"    m.BIKO, " +
				"    m.CREATED_DT, " +
				"    m.CREATED_BY, " +
				"    m.LASTUPDATE_DT, " +
				"    m.LASTUPDATE_BY, " +
				"    m.NOUNYU_KBN, " +
				"    a.CONTROLLER_FLG, " +
				"    kb.KOTAIKANRI_FLG, " +
				"    CASE WHEN mk.YUKO_FLG=0 THEN '' ELSE N'× ' END + mk.SEIZOMAKER_NK AS SEIZOMAKER_NK, " +
				"    mk.SEIZOMAKER_ID, " +
				"    kb.KIKIBUNRUI_CD, " +
				"    kb.OYAKIKIBUNRUI_CD, " +
				"    CASE kb.KOTAIKANRI_FLG " +
				"        WHEN 1 THEN N'○ ' + kb.KIKIBUNRUI_NK " +
				"        ELSE N'　　' + kb.KIKIBUNRUI_NK " +
				"    END AS KIKIBUNRUI_NK " +
				"FROM MCM.MCM_MA_KIKIMEISAI m " +
				"INNER JOIN MCM.MCM_MA_ATSUKAIKIKI a ON m.ATSUKAIKIKI_ID = a.ATSUKAIKIKI_ID " +
				"INNER JOIN MCM.MCM_MA_KIKIBUNRUI kb ON a.KIKIBUNRUI_ID = kb.KIKIBUNRUI_ID " +
				"INNER JOIN MCM.MCM_MA_SEIZOMAKER mk ON a.SEIZOMAKER_ID = mk.SEIZOMAKER_ID " +
				"WHERE m.KIKIKOSEI_ID = ? " +
				"ORDER BY m.HYOJIJUN ASC";
        return jdbc.queryForList(sql, kikikoseiId);
    }

    // ========================================================================
    // DataGridView3: 個体管理検索（JOIN）
    // 【変換元】Mcm0013uScreen.vb — Search() → Fill MCM_MA_KIKIKOTAIKANRI
    //   元SQL(DataSet): SELECT k.*, ak.ATSUKAIKIKI_NK, ak.KATASHIKI,
    //                          t.TORIHIKISAKI_NK
    //                   FROM MCM_MA_KIKIKOTAIKANRI k
    //                   LEFT JOIN MCM_MA_ATSUKAIKIKI_KOSEI ak ON ...
    //                   LEFT JOIN MCM_MA_BRAND_KOSEI bk ON ...
    //                   LEFT JOIN MCM_MA_BRAND b ON ...
    //                   LEFT JOIN MCM_MA_TORIHIKISAKI t ON ...
    //                   WHERE k.KIKIKOSEI_ID = :kikikoseiId
    //                   ORDER BY k.HYOJIJUN
    // ========================================================================
    public List<Map<String, Object>> findKotaikanriByKikikoseiId(BigDecimal kikikoseiId) {
        String sql =
        		"SELECT k.KOTAIKANRI_ID, k.KIKIKOSEI_ID, k.BRANDKOSEI_ID, k.CREATED_DT, k.CREATED_BY, k.LASTUPDATE_DT, k.LASTUPDATE_BY, " +
				"       k.KOTAI_NK, k.SERIAL_NO, " +
				"       FORMAT(k.ITIJINONYU_DT, 'yyyy/MM/dd') AS ITIJINONYU_DT, " +
				"       k.SETCHIBASYO, " +
				"       FORMAT(k.TEKKYO_DT, 'yyyy/MM/dd') AS TEKKYO_DT, " +
				"       FORMAT(k.KEIYAKUKIGEN_DT, 'yyyy/MM/dd') AS KEIYAKUKIGEN_DT, " +
				"       FORMAT(k.ENCHOKEIYAKUKIGEN_DT, 'yyyy/MM/dd') AS ENCHOKEIYAKUKIGEN_DT, " +
				"       FORMAT(k.KEIYAKUMANRYOYOTEI_DT, 'yyyy/MM/dd') AS KEIYAKUMANRYOYOTEI_DT, " +
				"       FORMAT(k.UPSKOKAN_DT, 'yyyy/MM/dd') AS UPSKOKAN_DT, " +
				"       k.HYOJIJUN, k.BIKO, " +
				"       k.ATSUKAIKIKIKOSEI_ID, " +
				"       akk.ATSUKAIKIKI_ID, " +
				"       ak.ATSUKAIKIKI_NK, ak.KATASHIKI, " +
				"       (SELECT MIN(mf.KIKIMEISAI_ID) " +
				"        FROM MCM.MCM_MA_KIKIMEISAI mf " +
				"        WHERE mf.KIKIKOSEI_ID = k.KIKIKOSEI_ID " +
				"          AND mf.ATSUKAIKIKI_ID = ak.ATSUKAIKIKI_ID " +
				"       ) AS KIKIMEISAI_ID, " +
				"       t.TORIHIKISAKI_NK " +
				"FROM MCM.MCM_MA_KIKIKOTAIKANRI k " +
				"LEFT JOIN MCM.MCM_MA_ATSUKAIKIKI_KOSEI akk " +
				"    ON k.ATSUKAIKIKIKOSEI_ID = akk.ATSUKAIKIKIKOSEI_ID " +
				"LEFT JOIN MCM.MCM_MA_ATSUKAIKIKI ak " +
				"    ON akk.ATSUKAIKIKI_ID = ak.ATSUKAIKIKI_ID " +
				"LEFT JOIN MCM.MCM_MA_TORIHIKISAKI t " +
				"    ON akk.TORIHIKISAKI_ID = t.TORIHIKISAKI_ID " +
				"WHERE k.KIKIKOSEI_ID = ? " +
				"ORDER BY k.HYOJIJUN ASC";
        return jdbc.queryForList(sql, kikikoseiId);
    }

    // ========================================================================
    // コンボボックスデータ
    // 【変換元】Mcm0013uScreen.vb — Search() → ComboBox設定
    // ========================================================================
    public List<Map<String, Object>> getBrandKoseiCombo() {
        // TODO: 実際のブランド構成コンボデータ取得
        //       SELECT BRANDKOSEI_ID, BRANDKOSEI_NK FROM MCM_MA_BRAND_KOSEI ORDER BY HYOJIJUN
        try {
            return jdbc.queryForList(
                "SELECT BRANDKOSEI_ID, BRANDSYOSAI_NK FROM MCM.MCM_MA_BRAND_KOSEI ORDER BY HYOJIJUN");
        } catch (Exception e) {
            log.warn("ブランド構成コンボ取得失敗", e);
            return List.of();
        }
    }
    
    /**
     * ★ 取扱機器構成コンボデータ取得
     * 機器個体の取引先選択用。取引先名 + 取扱機器名 を選択肢として表示。
     */
    public List<Map<String, Object>> getAtsukaikikiKoseiCombo() {
        String sql =
            "SELECT akk.ATSUKAIKIKIKOSEI_ID, akk.ATSUKAIKIKI_ID, " +
            "       a.ATSUKAIKIKI_NK, t.TORIHIKISAKI_NK " +
            "FROM MCM.MCM_MA_ATSUKAIKIKI_KOSEI akk " +
            "INNER JOIN MCM.MCM_MA_ATSUKAIKIKI a ON akk.ATSUKAIKIKI_ID = a.ATSUKAIKIKI_ID " +
            "INNER JOIN MCM.MCM_MA_TORIHIKISAKI t ON akk.TORIHIKISAKI_ID = t.TORIHIKISAKI_ID " +
            "ORDER BY t.TORIHIKISAKI_NK, a.ATSUKAIKIKI_NK";
        try {
            return jdbc.queryForList(sql);
        } catch (Exception e) {
            log.warn("取扱機器構成コンボ取得失敗", e);
            return List.of();
        }
    }

    public List<Map<String, Object>> getTaniCombo() {
        // VB McmTaniDataTable: 登録済み単位を使用件数降順・名称昇順で表示。
        return jdbc.queryForList("SELECT TANI AS value, TANI AS label FROM MCM.MCM_MA_KIKIKOSEI WHERE TANI IS NOT NULL GROUP BY TANI ORDER BY COUNT(TANI) DESC, TANI ASC");
    }
    // 画面の全構成・全個体を一度読み、以降の操作はこの作業データを編集する。
    public Mcm0013uForm loadForm(Long plantId) {
        Mcm0013uForm f = new Mcm0013uForm(); f.setPlantId(plantId);
        var h=getHeaderInfo(plantId);
        f.setNonyusakiCd(h.get("nonyusakiCd")); f.setNonyusakiNk(h.get("nonyusakiNk"));
        f.setSupportId(h.get("supportId")); f.setPlantNk(h.get("plantNk"));
        for (var m:findKoseiByPlantId(longToBd(plantId))) {
            KoseiRowForm r=mapBean(m,new KoseiRowForm()); r.setRowStatus("unchanged");
            f.getKoseiRows().add(r); f.getKoseiDisplay().put(r.getKikikoseiId(),new HashMap<>(m));
            List<Mcm0014uDeliveryDto.MeisaiRow> details=new ArrayList<>();
            for(var d:findMeisaiByKikikoseiId(longToBd(r.getKikikoseiId()))) {
                var detail=mapBean(d,new Mcm0014uDeliveryDto.MeisaiRow());
                detail.setKotaikanriFlg(bd(d.get("KOTAIKANRI_FLG"))); detail.setRowStatus("unchanged");
                details.add(detail);
            }
            f.getDetails().put(r.getKikikoseiId(),details);
            for(var k:findKotaikanriByKikikoseiId(longToBd(r.getKikikoseiId()))) {
                var individual=mapBean(k,new KotaikanriRowForm()); individual.setRowStatus("unchanged");
                f.getKotaikanriRows().add(individual); f.getKotaiDisplay().put(individual.getKotaikanriId(),new HashMap<>(k));
            }
        }
        if(!f.getKoseiRows().isEmpty()) f.setSelectedKikikoseiId(f.getKoseiRows().get(0).getKikikoseiId());
        return f;
    }

    private static <T> T mapBean(Map<String,Object> data,T bean) {
        var w=new org.springframework.beans.BeanWrapperImpl(bean);
        for(var p:w.getPropertyDescriptors()) {
            String key=p.getName().replaceAll("([a-z0-9])([A-Z])","$1_$2").toUpperCase(java.util.Locale.ROOT);
            if(w.isWritableProperty(p.getName()) && data.containsKey(key)) {
                Object v=data.get(key);
                if(v!=null&&p.getPropertyType()==String.class) v=v instanceof BigDecimal n?n.stripTrailingZeros().toPlainString():v.toString();
                if(v instanceof java.sql.Timestamp && p.getPropertyType()==LocalDateTime.class) v=((java.sql.Timestamp)v).toLocalDateTime();
                w.setPropertyValue(p.getName(),v);
            }
        }
        return bean;
    }

    public List<Map<String,Object>> getBrandKoseiCombo(Long plantId) {
        return jdbc.queryForList("SELECT BRANDKOSEI_ID, BRANDSYOSAI_NK FROM MCM.MCM_MA_BRAND_KOSEI WHERE PLANT_ID = ? ORDER BY HYOJIJUN",plantId);
    }
    public static boolean active(String status) { return !"deleted".equals(status); }
    private static boolean changed(String status) { return "new".equals(status)||"modified".equals(status)||"deleted".equals(status); }
    public boolean hasChanges(Mcm0013uForm f) {
        return f.getKoseiRows().stream().anyMatch(r->changed(r.getRowStatus()))
            || f.getKotaikanriRows().stream().anyMatch(r->changed(r.getRowStatus()))
            || f.getDetails().values().stream().flatMap(List::stream).anyMatch(r->changed(r.getRowStatus()));
    }
    public List<Mcm0014uDeliveryDto.MeisaiRow> activeDetails(Mcm0013uForm f,Long id) {
        return f.getDetails().getOrDefault(id,List.of()).stream().filter(r->active(r.getRowStatus())).toList();
    }
    public List<KotaikanriRowForm> activeIndividuals(Mcm0013uForm f,Long id) {
        return f.getKotaikanriRows().stream().filter(r->active(r.getRowStatus())&&java.util.Objects.equals(id,r.getKikikoseiId())).toList();
    }
    public List<String> checkRequired(KoseiRowForm r) {
        if(r==null) return List.of(Mcm0013uConstants.MSG_NO_SELECTION);
        if(isEmpty(r.getKikikoseiNk())||r.getSetNm()==null||isEmpty(r.getTani())) return List.of(Mcm0013uConstants.MSG_REQUIRED_FIELDS_EMPTY);
        if(r.getSetNm()<1) return List.of(Mcm0013uConstants.MSG_SETNUM_INVALID);
        if(r.getSetNm()>9999) return List.of("セット数は4桁以下で入力してください。");
        return List.of();
    }
    public List<String> checkDeleteKosei(Mcm0013uForm f,Long id) {
        var r=f.findKoseiRow(id);
        if(r==null||!active(r.getRowStatus())) return List.of(Mcm0013uConstants.MSG_NO_SELECTION);
        if(!activeIndividuals(f,id).isEmpty()) return List.of(Mcm0013uConstants.MSG_DELETE_KOTAI_REF);
        if(!activeDetails(f,id).isEmpty()) return List.of(Mcm0013uConstants.MSG_DELETE_MEISAI_REF);
        return List.of();
    }
    public List<String> checkDeleteKotai(Mcm0013uForm f,Long id) {
        var r=f.findKotaiRow(id);
        if(r==null||!active(r.getRowStatus())) return List.of(Mcm0013uConstants.MSG_NO_SELECTION);
        var parent=f.findKoseiRow(r.getKikikoseiId());
        for(var d:activeDetails(f,r.getKikikoseiId())) {
            if(same(d.getKikimeisaiId(),r.getKikimeisaiId())) {
                long remaining=activeIndividuals(f,r.getKikikoseiId()).stream()
                    .filter(k->same(d.getKikimeisaiId(),k.getKikimeisaiId())&&!java.util.Objects.equals(id,k.getKotaikanriId())).count();
                if(parent.getSetNm()==null||remaining<parent.getSetNm()) return List.of(Mcm0013uConstants.MSG_KOTAI_DELETE_SETNUM);
            }
        }
        return List.of();
    }

    // VB RowAddSetNum: 不足分だけ仮行を追加。減少時は自動削除せず、登録で数量を検査する。
    public void syncWorkingIndividuals(Mcm0013uForm f,Long id) {
        var parent=f.findKoseiRow(id);
        if(parent==null||parent.getSetNm()==null||parent.getSetNm()<1||parent.getSetNm()>9999) return;
        int order=f.getKotaikanriRows().stream().filter(k->java.util.Objects.equals(id,k.getKikikoseiId()))
            .mapToInt(k->k.getHyojijun()==null?0:k.getHyojijun()).max().orElse(0);
        for(var d:activeDetails(f,id)) {
            if(d.getControllerFlg()!=null&&d.getControllerFlg().intValue()==1&& !"1".equals(parent.getControllerFlg())) {
                parent.setControllerFlg("1"); mark(parent);
            }
            if(d.getKotaikanriFlg()==null||d.getKotaikanriFlg().intValue()!=1) continue;
            long count=activeIndividuals(f,id).stream().filter(k->same(d.getKikimeisaiId(),k.getKikimeisaiId())).count();
            for(long n=count;n<parent.getSetNm();n++) {
                var k=new KotaikanriRowForm(); k.setKotaikanriId(f.nextTempId()); k.setKikikoseiId(id);
                k.setKikimeisaiId(d.getKikimeisaiId().longValue());
                k.setAtsukaikikiId(d.getAtsukaikikiId()==null?null:d.getAtsukaikikiId().longValue());
                k.setKotaiNk(d.getKikihinmeiNk()); k.setHyojijun(++order); k.setRowStatus("new");
                f.getKotaikanriRows().add(k);
                var display=new HashMap<String,Object>(); display.put("ATSUKAIKIKI_NK",d.getKikihinmeiNk()); display.put("KATASHIKI",d.getKikikatashiki());
                f.getKotaiDisplay().put(k.getKotaikanriId(),display);
            }
        }
    }
    public static void mark(KoseiRowForm r) { if(!"new".equals(r.getRowStatus())&&active(r.getRowStatus())) r.setRowStatus("modified"); }

    // 14Uの復帰内容はDBへ適用しない。複数構成・複数往復でも元の仮IDを保持する。
    public void acceptDelivery(Mcm0013uForm f,Mcm0014uDeliveryDto delivery) {
        Long id=delivery.getKikikoseiId(); var parent=f.findKoseiRow(id);
        if(parent==null||!java.util.Objects.equals(f.getPlantId(),delivery.getPlantId())) return;
        var current=f.getDetails().computeIfAbsent(id,k->new ArrayList<>());
        for(var incoming:delivery.getMeisaiRows()) {
            var old=current.stream().filter(d->d.getKikimeisaiId()!=null&&incoming.getKikimeisaiId()!=null&&d.getKikimeisaiId().compareTo(incoming.getKikimeisaiId())==0).findFirst().orElse(null);
            if(old!=null) {
                String status=old.getRowStatus(); current.remove(old);
                if("new".equals(status)&&active(incoming.getRowStatus())) incoming.setRowStatus("new");
                else if("modified".equals(status)&&"unchanged".equals(incoming.getRowStatus())) incoming.setRowStatus("modified");
            } else if(active(incoming.getRowStatus())) {
                incoming.setKikimeisaiId(BigDecimal.valueOf(f.nextTempId())); incoming.setRowStatus("new");
            } else continue;
            incoming.setKikikoseiId(longToBd(id)); current.add(incoming);
        }
        if(!java.util.Objects.equals(parent.getSetNm(),delivery.getSetNum())) {parent.setSetNm(delivery.getSetNum());mark(parent);}
        f.setSelectedKikikoseiId(id); syncWorkingIndividuals(f,id);
    }

    /** 全構成を検証し終えるまで、INSERT/UPDATE/DELETEを一切実行しない。 */
    public List<String> validateWorking(Mcm0013uForm f) {
        if(!hasChanges(f)) return List.of(Mcm0013uConstants.MSG_NO_CHANGE);
        for(var r:f.getActiveKoseiRows()) {
            var e=checkRequired(r); if(!e.isEmpty()) return e;
            e=validateKoseiFields(r); if(!e.isEmpty()) return e;
        }
        for(var k:f.getKotaikanriRows()) if(active(k.getRowStatus())) {
            var e=validateKotaiFields(k); if(!e.isEmpty()) return e;
        }
        var partners=getAtsukaikikiKoseiCombo();
        for(var r:f.getActiveKoseiRows()) {
            var details=activeDetails(f,r.getKikikoseiId()); var individuals=activeIndividuals(f,r.getKikikoseiId());
            if(details.isEmpty()||individuals.isEmpty()) return List.of(Mcm0013uConstants.MSG_SETNUM_KOTAI_MISMATCH);
            // VBの明細IDは「構成＋取扱機器」から求める。重複時に別明細の個体を流用しない。
            Set<BigDecimal> equipment=new HashSet<>();
            for(var d:details) {
                if(!equipment.add(d.getAtsukaikikiId())) return List.of("データが重複しています。");
                if(d.getKotaikanriFlg()!=null&&d.getKotaikanriFlg().intValue()==1) {
                    long count=individuals.stream().filter(k->same(d.getKikimeisaiId(),k.getKikimeisaiId())).count();
                    if(count!=r.getSetNm()) return List.of(Mcm0013uConstants.MSG_SETNUM_KOTAI_MISMATCH);
                }
            }
            for(var k:individuals) {
                var detail=details.stream().filter(d->same(d.getKikimeisaiId(),k.getKikimeisaiId())&&d.getKotaikanriFlg()!=null&&d.getKotaikanriFlg().intValue()==1).findFirst().orElse(null);
                if(detail==null||partners.stream().noneMatch(p->same(bd(p.get("ATSUKAIKIKIKOSEI_ID")),k.getAtsukaikikikoseiId())&&same(bd(p.get("ATSUKAIKIKI_ID")),k.getAtsukaikikiId())))
                    return List.of(Mcm0013uConstants.MSG_KOTAI_MEISAI_MISMATCH);
            }
            boolean ctrl=details.stream().anyMatch(d->d.getControllerFlg()!=null&&d.getControllerFlg().intValue()==1);
            if(ctrl&&!"1".equals(r.getControllerFlg())) return List.of(Mcm0013uConstants.MSG_CONTROLLER_FLG_MISMATCH_91);
            if(!ctrl&&"1".equals(r.getControllerFlg())) return List.of(Mcm0013uConstants.MSG_CONTROLLER_FLG_MISMATCH_92);
        }
        for(var r:f.getKoseiRows()) if(!active(r.getRowStatus())) {
            Long id=r.getKikikoseiId();
            if(!activeIndividuals(f,id).isEmpty()) return List.of(Mcm0013uConstants.MSG_DELETE_KOTAI_REF);
            if(!activeDetails(f,id).isEmpty()) return List.of(Mcm0013uConstants.MSG_DELETE_MEISAI_REF);
            // 画面を開いた後に別処理で追加された子も削除しない。
            if(id>0) {
                for(var k:findKotaikanriByKikikoseiId(longToBd(id))) {
                    var known=f.findKotaiRow(((Number)k.get("KOTAIKANRI_ID")).longValue());
                    if(known==null||active(known.getRowStatus())) return List.of(Mcm0013uConstants.MSG_DELETE_KOTAI_REF);
                }
                for(var d:findMeisaiByKikikoseiId(longToBd(id)))
                    if(f.getDetails().getOrDefault(id,List.of()).stream().noneMatch(x->!active(x.getRowStatus())&&x.getKikimeisaiId().compareTo(bd(d.get("KIKIMEISAI_ID")))==0)) return List.of(Mcm0013uConstants.MSG_DELETE_MEISAI_REF);
            }
        }
        for(var details:f.getDetails().values()) for(var d:details)
            if(!active(d.getRowStatus()) && d.getKikimeisaiId()!=null && d.getKikimeisaiId().signum()>0) {
                var errors=new Mcm0014uService(jdbc).checkDeleteReference(d.getKikimeisaiId().longValueExact());
                if(!errors.isEmpty()) return errors;
            }
        for(var k:f.getKotaikanriRows()) if(!active(k.getRowStatus())&&k.getKotaikanriId()>0) {
            var e=checkKotaiReferences(longToBd(k.getKotaikanriId())); if(!e.isEmpty()) return e;
        }
        return List.of();
    }
    private List<String> checkKotaiReferences(BigDecimal id) {
        if(kotaiRepo.countTmKotaimeisaiByKotaikanriId(id)>0||kotaiRepo.countUkKotaimeisaiByKotaikanriId(id)>0) return List.of(Mcm0013uConstants.MSG_KOTAI_REF_TM);
        return List.of();
    }

    // Shift-JISバイト数と厳密な日付書式を、登録前・確認前で共通に使用する。
    private static String lengthError(String label,String value,int max) {
        return value!=null&&value.getBytes(java.nio.charset.Charset.forName("Shift_JIS")).length>max ? label+"は"+max+"桁以下で入力してください。":null;
    }
    private static List<String> first(String... messages) {for(String m:messages) if(m!=null)return List.of(m);return List.of();}
    private List<String> validateKoseiFields(KoseiRowForm r) {
        return first(numberError("No",r.getHyojijun(),6),lengthError("機器構成名",r.getKikikoseiNk(),100),lengthError("単位",r.getTani(),10),lengthError("手配製番",r.getTehaiseiban(),50),lengthError("備考",r.getBiko(),4000));
    }
    private static String numberError(String label,Integer n,int digits) {
        if(n==null)return label+"は必ず入力してください。";
        return Long.toString(Math.abs((long)n)).length()>digits?label+"は"+digits+"桁以下で入力してください。":null;
    }
    private List<String> validateKotaiFields(KotaikanriRowForm k) {
        return first(numberError("No",k.getHyojijun(),6),k.getAtsukaikikikoseiId()==null?"取引先は必ず入力してください。":null,k.getBrandkoseiId()==null?"ブランドは必ず入力してください。":null,
            lengthError("シリアル番号",k.getSerialNo(),100),lengthError("個体名",k.getKotaiNk(),60),lengthError("設置場所",k.getSetchibasyo(),60),lengthError("備考",k.getBiko(),4000),
            dateError("一次納入日",k.getItijinonyuDt()),dateError("撤去日",k.getTekkyoDt()),dateError("契約期限",k.getKeiyakukigenDt()),dateError("契約延長期限",k.getEnchokeiyakukigenDt()),dateError("契約満了予定日",k.getKeiyakumanryoyoteiDt()),dateError("UPS交換日",k.getUpskokanDt()));
    }
    private static String dateError(String label,String value) {
        if(isEmpty(value))return null;
        try {parseDate(value);return null;}catch(java.time.DateTimeException e){return label+"はYYYY/MM/DDの書式で入力してください。";}
    }
    private static LocalDateTime parseDate(String s) {
        if(isEmpty(s))return null;
        if(!s.matches("\\d{4}/\\d{2}/\\d{2}"))throw new java.time.DateTimeException("date format");
        var date=java.time.LocalDate.parse(s,DT_FMT);
        if(date.getYear()<1)throw new java.time.DateTimeException("date year");
        return date.atStartOfDay();
    }
    private static boolean isEmpty(String s) {return s==null||s.trim().isEmpty();}
    private static BigDecimal longToBd(Long v) {return v==null?null:BigDecimal.valueOf(v);}
    private static BigDecimal bd(Object v) {return v==null?null:new BigDecimal(v.toString());}
    private static boolean same(BigDecimal a,Long b) {return a!=null&&b!=null&&a.compareTo(BigDecimal.valueOf(b))==0;}
    private BigDecimal nextId(String table,String column) {
        // 引数はこのクラス内の固定のテーブル名・列名のみ。
        return jdbc.queryForObject("SELECT COALESCE(MAX("+column+"),0)+1 FROM MCM."+table+" WITH (UPDLOCK,HOLDLOCK)",BigDecimal.class);
    }

    @Transactional(isolation=org.springframework.transaction.annotation.Isolation.SERIALIZABLE)
    public List<String> saveWorking(Mcm0013uForm form,String user) {
        List<String> errors=validateWorking(form); if(!errors.isEmpty())return errors;
        LocalDateTime now=LocalDateTime.now();
        // 子の明示的な削除のみ。親や明細からの一括カスケード削除は行わない。
        for(var k:form.getKotaikanriRows()) if(!active(k.getRowStatus())&&k.getKotaikanriId()>0) kotaiRepo.deleteById(longToBd(k.getKotaikanriId()));
        kotaiRepo.flush();
        for(var entries:form.getDetails().values()) for(var d:entries) if(!active(d.getRowStatus())&&d.getKikimeisaiId().signum()>0) meisaiRepo.deleteById(d.getKikimeisaiId());
        meisaiRepo.flush();
        for(var r:form.getKoseiRows()) if(!active(r.getRowStatus())&&r.getKikikoseiId()>0) {
            BigDecimal id=longToBd(r.getKikikoseiId());
            if(koseiRepo.countKotaikanriByKikikoseiId(id)>0)throw new IllegalStateException(Mcm0013uConstants.MSG_DELETE_KOTAI_REF);
            if(koseiRepo.countMeisaiByKikikoseiId(id)>0)throw new IllegalStateException(Mcm0013uConstants.MSG_DELETE_MEISAI_REF);
            jdbc.update("DELETE FROM MCM.MCM_MA_KIKIKOSEI WHERE PLANT_ID = ? AND KIKIKOSEI_ID = ?",form.getPlantId(),id);
        }
        // 名前ではなく仮IDで対応させる。同名構成を複数追加しても混線しない。
        Map<Long,Long> ids=new HashMap<>();
        for(var r:form.getActiveKoseiRows()) {
            Long old=r.getKikikoseiId();
            if(changed(r.getRowStatus()))saveKoseiRow(r,form.getPlantId(),user,now,"new".equals(r.getRowStatus()));
            ids.put(old,r.getKikikoseiId());
        }
        for(var entry:form.getDetails().entrySet()) {
            Long parent=ids.get(entry.getKey()); if(parent==null)continue;
            for(var d:entry.getValue()) if(active(d.getRowStatus())&&changed(d.getRowStatus())) {
                boolean isNew="new".equals(d.getRowStatus());
                var entity=isNew?new KikimeisaiEntity():meisaiRepo.findById(d.getKikimeisaiId()).orElseThrow(()->new IllegalStateException("他のユーザがデータを変更した可能性があります。処理をやり直してください。"));
                if(isNew) {entity.setKikimeisaiId(nextId("MCM_MA_KIKIMEISAI","KIKIMEISAI_ID"));entity.setCreatedDt(now);entity.setCreatedBy(user);}
                entity.setKikikoseiId(longToBd(parent));entity.setHyojijun(d.getHyojijun());entity.setNounyuKbn(d.getNounyuKbn());entity.setKikihinmeiNk(d.getKikihinmeiNk());entity.setKikikatashiki(d.getKikikatashiki());entity.setSuryoNm(d.getSuryoNm());entity.setBiko(d.getBiko());entity.setAtsukaikikiId(d.getAtsukaikikiId());entity.setLastupdateDt(now);entity.setLastupdateBy(user);
                meisaiRepo.saveAndFlush(entity);
            }
        }
        for(var k:form.getKotaikanriRows()) if(active(k.getRowStatus())&&changed(k.getRowStatus())) {
            k.setKikikoseiId(ids.get(k.getKikikoseiId()));
            saveKotaiRow(k,user,now,"new".equals(k.getRowStatus()));
        }
        return List.of();
    }

    private void saveKoseiRow(KoseiRowForm row, Long plantId,
                              String loginUser, LocalDateTime now, boolean isNew) {
        KikikoseiEntity e;

        if (isNew) {
            e = new KikikoseiEntity();
            // 新規ID採番: MAX(KIKIKOSEI_ID) + 1
            // 【変換元】Mcm0013uScreen.vb — kikikoseiIdMaxCnt (負数方式 → Java側はMAX+1)
            // ★ Part 2 修正: D-005 同時実行対策 — UPDLOCK付きMAXで採番
            BigDecimal newId = koseiRepo.findMaxKikikoseiIdWithLock().add(BigDecimal.ONE);
            e.setKikikoseiId(newId);
            e.setPlantId(longToBd(plantId));
            e.setCreatedDt(now);
            e.setCreatedBy(loginUser);
        } else {
            // 既存行取得
            // 既存行取得（複合キー: PLANT_ID + KIKIKOSEI_ID）
            e = koseiRepo.findByPlantIdAndKikikoseiId(longToBd(plantId), longToBd(row.getKikikoseiId()));
            if (e == null) throw new IllegalStateException("他のユーザがデータを変更した可能性があります。処理をやり直してください。");
        }

        e.setKikikoseiNk(row.getKikikoseiNk());
        e.setSetNm(row.getSetNm() != null ? BigDecimal.valueOf(row.getSetNm()) : BigDecimal.ONE);
        e.setTani(row.getTani());
        e.setTehaiseiban(row.getTehaiseiban());
        e.setControllerFlg(row.getControllerFlg() != null
                ? new BigDecimal(row.getControllerFlg())
                : BigDecimal.ZERO);
        e.setHyojijun(row.getHyojijun() != null ? BigDecimal.valueOf(row.getHyojijun()) : BigDecimal.ZERO);
        e.setBiko(row.getBiko());
        e.setLastupdateDt(now);
        e.setLastupdateBy(loginUser);

        koseiRepo.saveAndFlush(e);
        row.setKikikoseiId(e.getKikikoseiId().longValue());
    }

    // ========================================================================
    // 個体管理行保存（INSERT / UPDATE）
    // 【変換元】Mcm0013uScreen.vb — DaoContainer.Update MCM_MA_KIKIKOTAIKANRI
    // ========================================================================
    private void saveKotaiRow(KotaikanriRowForm row, String loginUser,
                              LocalDateTime now, boolean isNew) {
        KikikotaikanriEntity e;
        if (isNew) {
            e = new KikikotaikanriEntity();
            BigDecimal newId = nextId("MCM_MA_KIKIKOTAIKANRI", "KOTAIKANRI_ID");
            e.setKotaikanriId(newId);
            e.setCreatedDt(now);
            e.setCreatedBy(loginUser);
        } else {
            e = kotaiRepo.findById(BigDecimal.valueOf(row.getKotaikanriId()))
                    .orElseThrow(() -> new IllegalStateException("他のユーザがデータを変更した可能性があります。処理をやり直してください。"));
        }
        e.setKikikoseiId(longToBd(row.getKikikoseiId()));
        e.setBrandkoseiId(longToBd(row.getBrandkoseiId()));
        e.setAtsukaikikikoseiId(longToBd(row.getAtsukaikikikoseiId()));
        e.setKotaiNk(row.getKotaiNk());
        e.setSerialNo(row.getSerialNo());
        e.setItijinonyuDt(parseDate(row.getItijinonyuDt()));
        e.setSetchibasyo(row.getSetchibasyo());
        e.setTekkyoDt(parseDate(row.getTekkyoDt()));
        e.setKeiyakukigenDt(parseDate(row.getKeiyakukigenDt()));
        e.setEnchokeiyakukigenDt(parseDate(row.getEnchokeiyakukigenDt()));
        e.setKeiyakumanryoyoteiDt(parseDate(row.getKeiyakumanryoyoteiDt()));
        e.setUpskokanDt(parseDate(row.getUpskokanDt()));
        e.setBiko(row.getBiko());
        e.setHyojijun(row.getHyojijun() != null ? BigDecimal.valueOf(row.getHyojijun()) : null);
        e.setLastupdateDt(now);
        e.setLastupdateBy(loginUser);
        kotaiRepo.saveAndFlush(e);
        row.setKotaikanriId(e.getKotaikanriId().longValue());
    }
    

}
