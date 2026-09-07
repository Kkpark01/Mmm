/**
 * 【変換元】Mcm0014uScreen.Designer.vb — コントロール定義
 *   SEIZOMAKERComboBox / KIKIHINMEITextBox / KIKIKATASHIKITextBox
 *   KIKIBUNRUIComboBox / TORIHIKISAKIComboBox
 *   ATSUKAIKIKIDataGridView(取扱機器) / KIKIMEISAIDataGridView(構成明細)
 *
 * MCM0014U — 納入機器明細選定 フォーム
 * @since 2026-06-08
 */
package com.daifuku.mcm.form;

import java.util.ArrayList;
import java.util.List;

public class Mcm0014uForm implements java.io.Serializable {
    private String seizomakerText;
    public String getSeizomakerText(){return seizomakerText;}
    public void setSeizomakerText(String v){seizomakerText=v;}
    private String setNmText;
    private String editToken;
    public String getSetNmText(){return setNmText;}
    public void setSetNmText(String v){setNmText=v;}
    public String getEditToken(){return editToken;}
    public void setEditToken(String v){editToken=v;}

    // ── MCM0013Uからの受渡しデータ ──
    /** 機器構成ID — VB.NET: DeliveryData.KikikoseiId */
    private Long kikikoseiId;
    /** 機器構成名（表示用） */
    private String kikikoseiNk;
    /** セット数（表示用） */
    private Integer setNm;
    
    private Long plantId;

    // ── 検索条件 ──
    /** 製造メーカーID — VB.NET: SEIZOMAKERComboBox */
    private Long seizomakerId;
    /** 品名 — VB.NET: KIKIHINMEITextBox */
    private String kikihinmeiNk;
    /** 型式 — VB.NET: KIKIKATASHIKITextBox */
    private String kikikatashiki;
    /** 機器分類ID — VB.NET: KIKIBUNRUIComboBox */
    private Long kikibunruiId;
    /** 取引先ID — VB.NET: TORIHIKISAKIComboBox */
    private Long torihikisakiId;

    /** 選択した取扱機器ID */
    private Long selectedAtsukaikikiId;

    /** DataGridView2: 構成明細行リスト */
    private List<MeisaiRowForm> meisaiRows = new ArrayList<>();
    
    // Getter / Setter
    public Long getKikikoseiId() { return kikikoseiId; }
    public void setKikikoseiId(Long v) { this.kikikoseiId = v; }
    public String getKikikoseiNk() { return kikikoseiNk; }
    public void setKikikoseiNk(String v) { this.kikikoseiNk = v; }
    public Integer getSetNm() { return setNm; }
    public void setSetNm(Integer v) { this.setNm = v; }
    public Long getSeizomakerId() { return seizomakerId; }
    public void setSeizomakerId(Long v) { this.seizomakerId = v; }
    public String getKikihinmeiNk() { return kikihinmeiNk; }
    public void setKikihinmeiNk(String v) { this.kikihinmeiNk = v; }
    public String getKikikatashiki() { return kikikatashiki; }
    public void setKikikatashiki(String v) { this.kikikatashiki = v; }
    public Long getKikibunruiId() { return kikibunruiId; }
    public void setKikibunruiId(Long v) { this.kikibunruiId = v; }
    public Long getTorihikisakiId() { return torihikisakiId; }
    public void setTorihikisakiId(Long v) { this.torihikisakiId = v; }
    public Long getSelectedAtsukaikikiId() { return selectedAtsukaikikiId; }
    public void setSelectedAtsukaikikiId(Long v) { this.selectedAtsukaikikiId = v; }
    public List<MeisaiRowForm> getMeisaiRows() { return meisaiRows; }
    public void setMeisaiRows(List<MeisaiRowForm> v) { this.meisaiRows = v; }

    public Long getPlantId() {
		return plantId;
	}
	public void setPlantId(Long plantId) {
		this.plantId = plantId;
	}

	// ================================================================
    // 内部クラス: 構成明細行（KIKIMEISAIDataGridView）
    // 【変換元】KIKIMEISAIDataGridView列
    // ================================================================
    public static class MeisaiRowForm implements java.io.Serializable {
        private String hyojijunText;
        public String getHyojijunText(){return hyojijunText;}
        public void setHyojijunText(String v){hyojijunText=v;}
        private String suryoNmText;
        public String getSuryoNmText(){return suryoNmText;}
        public void setSuryoNmText(String v){suryoNmText=v;}
        private Long kikimeisaiId;
        private Long kikikoseiId;
        private Long atsukaikikiId;
        private Long seizomakerId;
        private String seizomakerNk;
        private String kikihinmeiNk;
        private String kikikatashiki;
        private Integer suryoNm;
        private String biko;
        private String nounyuKbn;
        private Integer hyojijun;
        private String kotaikanriFlg;
        private String controllerFlg;
        private String kikibunruiCd;
        private String oyakikibunruiCd;
        private String kikibunruiNk;
        private String rowStatus;
    	private java.time.LocalDateTime createdDt;
    	private String createdBy;
    	private java.time.LocalDateTime lastupdateDt;
    	private String lastupdateBy;


        public Long getKikimeisaiId() { return kikimeisaiId; }
        public void setKikimeisaiId(Long v) { this.kikimeisaiId = v; }
        public Long getKikikoseiId() { return kikikoseiId; }
        public void setKikikoseiId(Long v) { this.kikikoseiId = v; }
        public Long getAtsukaikikiId() { return atsukaikikiId; }
        public void setAtsukaikikiId(Long v) { this.atsukaikikiId = v; }
        public Long getSeizomakerId() { return seizomakerId; }
        public void setSeizomakerId(Long v) { this.seizomakerId = v; }
        public String getSeizomakerNk() { return seizomakerNk; }
        public void setSeizomakerNk(String v) { this.seizomakerNk = v; }
        public String getKikihinmeiNk() { return kikihinmeiNk; }
        public void setKikihinmeiNk(String v) { this.kikihinmeiNk = v; }
        public String getKikikatashiki() { return kikikatashiki; }
        public void setKikikatashiki(String v) { this.kikikatashiki = v; }
        public Integer getSuryoNm() { return suryoNm; }
        public void setSuryoNm(Integer v) { this.suryoNm = v; }
        public String getBiko() { return biko; }
        public void setBiko(String v) { this.biko = v; }
        public String getNounyuKbn() { return nounyuKbn; }
        public void setNounyuKbn(String v) { this.nounyuKbn = v; }
        public Integer getHyojijun() { return hyojijun; }
        public void setHyojijun(Integer v) { this.hyojijun = v; }
        public String getKotaikanriFlg() { return kotaikanriFlg; }
        public void setKotaikanriFlg(String v) { this.kotaikanriFlg = v; }
        public String getControllerFlg() { return controllerFlg; }
        public void setControllerFlg(String v) { this.controllerFlg = v; }
        public String getKikibunruiCd() { return kikibunruiCd; }
        public void setKikibunruiCd(String v) { this.kikibunruiCd = v; }
        public String getOyakikibunruiCd() { return oyakikibunruiCd; }
        public void setOyakikibunruiCd(String v) { this.oyakikibunruiCd = v; }
        public String getKikibunruiNk() { return kikibunruiNk; }
        public void setKikibunruiNk(String v) { this.kikibunruiNk = v; }
        public String getRowStatus() { return rowStatus; }
        public void setRowStatus(String v) { this.rowStatus = v; }
		public java.time.LocalDateTime getCreatedDt() { return createdDt; }
		public void setCreatedDt(java.time.LocalDateTime v) { this.createdDt = v; }
		public String getCreatedBy() { return createdBy; }
		public void setCreatedBy(String v) { this.createdBy = v; }
		public java.time.LocalDateTime getLastupdateDt() { return lastupdateDt; }
		public void setLastupdateDt(java.time.LocalDateTime v) { this.lastupdateDt = v; }
		public String getLastupdateBy() { return lastupdateBy; }
		public void setLastupdateBy(String v) { this.lastupdateBy = v; }

    }
}
