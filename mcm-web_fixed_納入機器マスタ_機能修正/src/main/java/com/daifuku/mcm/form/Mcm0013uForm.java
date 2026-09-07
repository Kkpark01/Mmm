/**
 * 【変換元】Mcm0013uScreen.Designer.vb — コントロール定義
 *   DataGridView1(機器構成) / DataGridView2(構成明細) / DataGridView3(個体管理)
 *   ヘッダ: NonyusakiCdLabel / NonyusakiNkLabel / SupportIdLabel / PlantNkLabel
 *
 * MCM0013U — 機器構成マスタ登録 フォーム
 * @since 2026-06-08
 */
package com.daifuku.mcm.form;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class Mcm0013uForm implements Serializable {
    private static final long serialVersionUID = 1L;

    // ── ヘッダ情報（MCM0012Uから受取）──
    /** プラントID — VB.NET: DeliveryData.PlantId */
    private Long plantId;
    /** 納入先コード（表示用） */
    private String nonyusakiCd;
    /** 納入先名（表示用） */
    private String nonyusakiNk;
    /** サポートID（表示用） */
    private String supportId;
    /** プラント名（表示用） */
    private String plantNk;

    /** 選択中の機器構成ID（DataGridView1で選択した行） */
    private Long selectedKikikoseiId;

    // ★ 追加: セッション管理用フィールド
    /** 仮ID採番カウンター（負数: -1, -2, -3, ...） */
    private int tempIdCounter = 0;
    /** 戻り先URL（MCM0012Uから引き継ぎ） */
    private String returnTo0012Url;

    /** DataGridView1: 機器構成パターン行リスト */
    private List<KoseiRowForm> koseiRows = new ArrayList<>();

    /** DataGridView3: 機器個体管理行リスト */
    private List<KotaikanriRowForm> kotaikanriRows = new ArrayList<>();

    // 全構成の編集中明細を保持する。DB保存は登録時だけ行う。
    private java.util.Map<Long, java.util.List<com.daifuku.mcm.dto.Mcm0014uDeliveryDto.MeisaiRow>> details = new java.util.LinkedHashMap<>();
    private java.util.Map<Long, java.util.Map<String,Object>> koseiDisplay = new java.util.HashMap<>();
    private java.util.Map<Long, java.util.Map<String,Object>> kotaiDisplay = new java.util.HashMap<>();
    public java.util.Map<Long, java.util.List<com.daifuku.mcm.dto.Mcm0014uDeliveryDto.MeisaiRow>> getDetails() { return details; }
    public java.util.Map<Long, java.util.Map<String,Object>> getKoseiDisplay() { return koseiDisplay; }
    public java.util.Map<Long, java.util.Map<String,Object>> getKotaiDisplay() { return kotaiDisplay; }
    public KotaikanriRowForm findKotaiRow(Long id) {
        return kotaikanriRows.stream().filter(r -> java.util.Objects.equals(id,r.getKotaikanriId())).findFirst().orElse(null);
    }
    public Mcm0013uForm copy() { return org.springframework.util.SerializationUtils.clone(this); }

    // ── Getter / Setter ──
    public Long getPlantId() { return plantId; }
    public void setPlantId(Long v) { this.plantId = v; }
    public String getNonyusakiCd() { return nonyusakiCd; }
    public void setNonyusakiCd(String v) { this.nonyusakiCd = v; }
    public String getNonyusakiNk() { return nonyusakiNk; }
    public void setNonyusakiNk(String v) { this.nonyusakiNk = v; }
    public String getSupportId() { return supportId; }
    public void setSupportId(String v) { this.supportId = v; }
    public String getPlantNk() { return plantNk; }
    public void setPlantNk(String v) { this.plantNk = v; }
    public Long getSelectedKikikoseiId() { return selectedKikikoseiId; }
    public void setSelectedKikikoseiId(Long v) { this.selectedKikikoseiId = v; }
    public List<KoseiRowForm> getKoseiRows() { return koseiRows; }
    public void setKoseiRows(List<KoseiRowForm> v) { this.koseiRows = v; }
    public List<KotaikanriRowForm> getKotaikanriRows() { return kotaikanriRows; }
    public void setKotaikanriRows(List<KotaikanriRowForm> v) { this.kotaikanriRows = v; }

    // ================================================================
    // 内部クラス: 機器構成パターン行（DataGridView1）
    // 【変換元】DataGridView1列: KIKIKOSEI_ID, KIKIKOSEI_NK, SET_NM,
    //   TANI, TEHAISEIBAN, CONTROLLER_FLG, HYOJIJUN, BIKO, SENTAKU
    // ================================================================
    public static class KoseiRowForm implements Serializable {
        private static final long serialVersionUID = 1L;
        private Long kikikoseiId;
        private Long plantId;
        private String kikikoseiNk;
        private Integer setNm;
        private String tani;
        private String tehaiseiban;
        private String controllerFlg;
        private Integer hyojijun;
        private String biko;
        private String rowStatus;

        public Long getKikikoseiId() { return kikikoseiId; }
        public void setKikikoseiId(Long v) { this.kikikoseiId = v; }
        public Long getPlantId() { return plantId; }
        public void setPlantId(Long v) { this.plantId = v; }
        public String getKikikoseiNk() { return kikikoseiNk; }
        public void setKikikoseiNk(String v) { this.kikikoseiNk = v; }
        public Integer getSetNm() { return setNm; }
        public void setSetNm(Integer v) { this.setNm = v; }
        public String getTani() { return tani; }
        public void setTani(String v) { this.tani = v; }
        public String getTehaiseiban() { return tehaiseiban; }
        public void setTehaiseiban(String v) { this.tehaiseiban = v; }
        public String getControllerFlg() { return controllerFlg; }
        public void setControllerFlg(String v) { this.controllerFlg = v; }
        public Integer getHyojijun() { return hyojijun; }
        public void setHyojijun(Integer v) { this.hyojijun = v; }
        public String getBiko() { return biko; }
        public void setBiko(String v) { this.biko = v; }
        public String getRowStatus() { return rowStatus; }
        public void setRowStatus(String v) { this.rowStatus = v; }
    }

    // ================================================================
    // 内部クラス: 機器個体管理行（DataGridView3）
    // 【変換元】DataGridView3列: KOTAIKANRI_ID, KIKIKOSEI_ID,
    //   BRANDKOSEI_ID, ATSUKAIKIKIKOSEI_ID, KOTAI_NK, SERIAL_NO,
    //   ITIJINONYU_DT, SETCHIBASYO, TEKKYO_DT, KEIYAKUKIGEN_DT,
    //   ENCHOKEIYAKUKIGEN_DT, KEIYAKUMANRYOYOTEI_DT, UPSKOKAN_DT, BIKO
    // ================================================================
    public static class KotaikanriRowForm implements Serializable {
        private static final long serialVersionUID = 1L;
        // 明細IDはVB同様の画面内関連キー（DBの追加列ではない）。
        private Long kikimeisaiId;
        private Long atsukaikikiId;
        public Long getKikimeisaiId() { return kikimeisaiId; }
        public void setKikimeisaiId(Long v) { kikimeisaiId=v; }
        public Long getAtsukaikikiId() { return atsukaikikiId; }
        public void setAtsukaikikiId(Long v) { atsukaikikiId=v; }
        private Long kotaikanriId;
        private Long kikikoseiId;
        private Long brandkoseiId;
        private Long atsukaikikikoseiId;
        private String kotaiNk;
        private String serialNo;
        private String itijinonyuDt;
        private String setchibasyo;
        private String tekkyoDt;
        private String keiyakukigenDt;
        private String enchokeiyakukigenDt;
        private String keiyakumanryoyoteiDt;
        private String upskokanDt;
        private String biko;
        private Integer hyojijun;
        private String rowStatus;

        public Long getKotaikanriId() { return kotaikanriId; }
        public void setKotaikanriId(Long v) { this.kotaikanriId = v; }
        public Long getKikikoseiId() { return kikikoseiId; }
        public void setKikikoseiId(Long v) { this.kikikoseiId = v; }
        public Long getBrandkoseiId() { return brandkoseiId; }
        public void setBrandkoseiId(Long v) { this.brandkoseiId = v; }
        public Long getAtsukaikikikoseiId() { return atsukaikikikoseiId; }
        public void setAtsukaikikikoseiId(Long v) { this.atsukaikikikoseiId = v; }
        public String getKotaiNk() { return kotaiNk; }
        public void setKotaiNk(String v) { this.kotaiNk = v; }
        public String getSerialNo() { return serialNo; }
        public void setSerialNo(String v) { this.serialNo = v; }
        public String getItijinonyuDt() { return itijinonyuDt; }
        public void setItijinonyuDt(String v) { this.itijinonyuDt = v; }
        public String getSetchibasyo() { return setchibasyo; }
        public void setSetchibasyo(String v) { this.setchibasyo = v; }
        public String getTekkyoDt() { return tekkyoDt; }
        public void setTekkyoDt(String v) { this.tekkyoDt = v; }
        public String getKeiyakukigenDt() { return keiyakukigenDt; }
        public void setKeiyakukigenDt(String v) { this.keiyakukigenDt = v; }
        public String getEnchokeiyakukigenDt() { return enchokeiyakukigenDt; }
        public void setEnchokeiyakukigenDt(String v) { this.enchokeiyakukigenDt = v; }
        public String getKeiyakumanryoyoteiDt() { return keiyakumanryoyoteiDt; }
        public void setKeiyakumanryoyoteiDt(String v) { this.keiyakumanryoyoteiDt = v; }
        public String getUpskokanDt() { return upskokanDt; }
        public void setUpskokanDt(String v) { this.upskokanDt = v; }
        public String getBiko() { return biko; }
        public void setBiko(String v) { this.biko = v; }
        public Integer getHyojijun() { return hyojijun; }
        public void setHyojijun(Integer v) { this.hyojijun = v; }
        public String getRowStatus() { return rowStatus; }
        public void setRowStatus(String v) { this.rowStatus = v; }
    }

    // ================================================================
    // ★ 追加: セッション管理用ヘルパーメソッド
    // ================================================================

    /** 新しい仮IDを生成する（-1, -2, -3, ...）。DB永続化時に正式IDに差替え。 */
    public Long nextTempId() {
        tempIdCounter--;
        return (long) tempIdCounter;
    }

    /** 指定IDが仮ID（負数）かどうか判定する */
    public static boolean isTempId(Long id) {
        return id != null && id < 0;
    }

    /** 削除されていない構成行のみ返す */
    public List<KoseiRowForm> getActiveKoseiRows() {
        List<KoseiRowForm> active = new ArrayList<>();
        for (KoseiRowForm row : koseiRows) {
            if (!"deleted".equals(row.getRowStatus())) {
                active.add(row);
            }
        }
        return active;
    }

    /** 指定kikikoseiIdの構成行を取得する */
    public KoseiRowForm findKoseiRow(Long kikikoseiId) {
        if (kikikoseiId == null) return null;
        for (KoseiRowForm row : koseiRows) {
            if (kikikoseiId.equals(row.getKikikoseiId())) {
                return row;
            }
        }
        return null;
    }

    // ── 追加フィールドの Getter / Setter ──
    public int getTempIdCounter() { return tempIdCounter; }
    public void setTempIdCounter(int v) { this.tempIdCounter = v; }
    public String getReturnTo0012Url() { return returnTo0012Url; }
    public void setReturnTo0012Url(String v) { this.returnTo0012Url = v; }
}
