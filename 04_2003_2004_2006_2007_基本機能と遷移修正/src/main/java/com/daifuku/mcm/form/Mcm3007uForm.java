package com.daifuku.mcm.form;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * 【変換元】Mcm3007uScreen.vb
 * MCM3007U 契約・解約フォーム
 *
 * グリッド構成:
 *   Grid1: MCM_MA_NONYUSAKI（納入先マスタ）
 *   Grid2: MCM_2004_V（UVA 店舗契約一覧）
 *   Grid3: MCM_1003_V（TKA 取引先契約一覧）
 *
 * 破棄操作:
 *   UM破棄: MCM_UK_KEIYAKU.JOTAI=3 + MCM_UM_KIHON_MITSUMORI.JOTAI=1
 *   TK破棄: MCM_TK_KEIYAKU.JOTAI=3 + MCM_TM_KEIYAKUJIKAN.JOTAI=1
 *          （CHECK: MCM_TK_SIHAIRAMEISAI MIN(TSUKI) >= today）
 */
public class Mcm3007uForm implements Serializable {

    private static final long serialVersionUID = 1L;

    // ===================================================================
    // 検索条件
    // ===================================================================

    private String nonyusakiCd  = "";
    private String nonyusakiNk  = "";
    private String supportId    = "";
    private String plantNk      = "";

    /** 店舗(UM)状態フィルタ */
    private boolean jotaiUmMitsumori = true;
    private boolean jotaiUmKeiyaku   = true;
    private boolean jotaiUmKaiyaku   = false;
    private boolean jotaiUmHaki      = false;

    /** 取引先(TK)状態フィルタ */
    private boolean jotaiTkMitsumori = true;
    private boolean jotaiTkKeiyaku   = true;
    private boolean jotaiTkKaiyaku   = false;
    private boolean jotaiTkHaki      = false;

    // ===================================================================
    // 検索結果: Grid1 納入先マスタ
    // ===================================================================
    private List<Grid1RowForm> grid1Rows = new ArrayList<>();

    // ===================================================================
    // 検索結果: Grid2 UVA店舗契約（MCM_2004_V）
    // ===================================================================
    private List<Grid2RowForm> grid2Rows = new ArrayList<>();

    // ===================================================================
    // 検索結果: Grid3 TKA取引先契約（MCM_1003_V）
    // ===================================================================
    private List<Grid3RowForm> grid3Rows = new ArrayList<>();

    // ===================================================================
    // 内部クラス
    // ===================================================================

    /** Grid1: 納入先マスタ */
    public static class Grid1RowForm implements Serializable {
        private java.util.Map<String,String> display = new java.util.LinkedHashMap<>();
        public java.util.Map<String,String> getDisplay(){return display;}
        public void setDisplay(java.util.Map<String,String> value){display=value;}
        private static final long serialVersionUID = 1L;
        private BigDecimal nonyusakiId;
        private String nonyusakiCd;
        private String nonyusakiNk;
        private String plantNk;
        private String supportId;
        private String torihikisakiNk;

        public BigDecimal getNonyusakiId() { return nonyusakiId; }
        public void setNonyusakiId(BigDecimal v) { this.nonyusakiId = v; }
        public String getNonyusakiCd() { return nonyusakiCd; }
        public void setNonyusakiCd(String v) { this.nonyusakiCd = v; }
        public String getNonyusakiNk() { return nonyusakiNk; }
        public void setNonyusakiNk(String v) { this.nonyusakiNk = v; }
        public String getPlantNk() { return plantNk; }
        public void setPlantNk(String v) { this.plantNk = v; }
        public String getSupportId() { return supportId; }
        public void setSupportId(String v) { this.supportId = v; }
        public String getTorihikisakiNk() { return torihikisakiNk; }
        public void setTorihikisakiNk(String v) { this.torihikisakiNk = v; }
    }

    /** Grid2: UVA店舗契約（MCM_2004_V） */
    public static class Grid2RowForm implements Serializable {
        private java.util.Map<String,String> display = new java.util.LinkedHashMap<>();
        public java.util.Map<String,String> getDisplay(){return display;}
        public void setDisplay(java.util.Map<String,String> value){display=value;}
        private static final long serialVersionUID = 1L;
        private BigDecimal ukKeiyakuId;
        private BigDecimal umKihonMitsumoriId;
        private String jotai;
        private String shoruiNo;
        private String nonyusakiNk;
        private String supportId;
        private String plantNk;
        private String keiyakuNo;
        private String keiyakuDt;
        private String kaiyakuDt;
        private String keiyakumanryoDt;
        private String hosyuGkin;
        private String torihikisakiNk;
        private String shoninJotai;

        public BigDecimal getUkKeiyakuId() { return ukKeiyakuId; }
        public void setUkKeiyakuId(BigDecimal v) { this.ukKeiyakuId = v; }
        public BigDecimal getUmKihonMitsumoriId() { return umKihonMitsumoriId; }
        public void setUmKihonMitsumoriId(BigDecimal v) { this.umKihonMitsumoriId = v; }
        public String getJotai() { return jotai; }
        public void setJotai(String v) { this.jotai = v; }
        public String getShoruiNo() { return shoruiNo; }
        public void setShoruiNo(String v) { this.shoruiNo = v; }
        public String getNonyusakiNk() { return nonyusakiNk; }
        public void setNonyusakiNk(String v) { this.nonyusakiNk = v; }
        public String getSupportId() { return supportId; }
        public void setSupportId(String v) { this.supportId = v; }
        public String getPlantNk() { return plantNk; }
        public void setPlantNk(String v) { this.plantNk = v; }
        public String getKeiyakuNo() { return keiyakuNo; }
        public void setKeiyakuNo(String v) { this.keiyakuNo = v; }
        public String getKeiyakuDt() { return keiyakuDt; }
        public void setKeiyakuDt(String v) { this.keiyakuDt = v; }
        public String getKaiyakuDt() { return kaiyakuDt; }
        public void setKaiyakuDt(String v) { this.kaiyakuDt = v; }
        public String getKeiyakumanryoDt() { return keiyakumanryoDt; }
        public void setKeiyakumanryoDt(String v) { this.keiyakumanryoDt = v; }
        public String getHosyuGkin() { return hosyuGkin; }
        public void setHosyuGkin(String v) { this.hosyuGkin = v; }
        public String getTorihikisakiNk() { return torihikisakiNk; }
        public void setTorihikisakiNk(String v) { this.torihikisakiNk = v; }
        public String getShoninJotai() { return shoninJotai; }
        public void setShoninJotai(String v) { this.shoninJotai = v; }
    }

    /** Grid3: TKA取引先契約（MCM_1003_V） */
    public static class Grid3RowForm implements Serializable {
        private java.util.Map<String,String> display = new java.util.LinkedHashMap<>();
        public java.util.Map<String,String> getDisplay(){return display;}
        public void setDisplay(java.util.Map<String,String> value){display=value;}
        private static final long serialVersionUID = 1L;
        private BigDecimal tkKeiyakuId;
        private BigDecimal tmKeiyakujikanId;
        private String jotai;
        private String iraino;
        private String torihikisakiCd;
        private String torihikisakiNk;
        private String nonyusakiNk;
        private String supportId;
        private String plantNk;
        private String keiyakuNo;
        private String keiyakuDt;
        private String kaiyakuDt;
        private String keiyakumanryoDt;
        private String hosyuhoho;
        private String keiyakujikantai;
        private String shoninJotai;

        public BigDecimal getTkKeiyakuId() { return tkKeiyakuId; }
        public void setTkKeiyakuId(BigDecimal v) { this.tkKeiyakuId = v; }
        public BigDecimal getTmKeiyakujikanId() { return tmKeiyakujikanId; }
        public void setTmKeiyakujikanId(BigDecimal v) { this.tmKeiyakujikanId = v; }
        public String getJotai() { return jotai; }
        public void setJotai(String v) { this.jotai = v; }
        public String getIraino() { return iraino; }
        public void setIraino(String v) { this.iraino = v; }
        public String getTorihikisakiCd() { return torihikisakiCd; }
        public void setTorihikisakiCd(String v) { this.torihikisakiCd = v; }
        public String getTorihikisakiNk() { return torihikisakiNk; }
        public void setTorihikisakiNk(String v) { this.torihikisakiNk = v; }
        public String getNonyusakiNk() { return nonyusakiNk; }
        public void setNonyusakiNk(String v) { this.nonyusakiNk = v; }
        public String getSupportId() { return supportId; }
        public void setSupportId(String v) { this.supportId = v; }
        public String getPlantNk() { return plantNk; }
        public void setPlantNk(String v) { this.plantNk = v; }
        public String getKeiyakuNo() { return keiyakuNo; }
        public void setKeiyakuNo(String v) { this.keiyakuNo = v; }
        public String getKeiyakuDt() { return keiyakuDt; }
        public void setKeiyakuDt(String v) { this.keiyakuDt = v; }
        public String getKaiyakuDt() { return kaiyakuDt; }
        public void setKaiyakuDt(String v) { this.kaiyakuDt = v; }
        public String getKeiyakumanryoDt() { return keiyakumanryoDt; }
        public void setKeiyakumanryoDt(String v) { this.keiyakumanryoDt = v; }
        public String getHosyuhoho() { return hosyuhoho; }
        public void setHosyuhoho(String v) { this.hosyuhoho = v; }
        public String getKeiyakujikantai() { return keiyakujikantai; }
        public void setKeiyakujikantai(String v) { this.keiyakujikantai = v; }
        public String getShoninJotai() { return shoninJotai; }
        public void setShoninJotai(String v) { this.shoninJotai = v; }
    }

    // ===================================================================
    // Getter / Setter（検索条件）
    // ===================================================================

    public String getNonyusakiCd() { return nonyusakiCd; }
    public void setNonyusakiCd(String v) { this.nonyusakiCd = v; }
    public String getNonyusakiNk() { return nonyusakiNk; }
    public void setNonyusakiNk(String v) { this.nonyusakiNk = v; }
    public String getSupportId() { return supportId; }
    public void setSupportId(String v) { this.supportId = v; }
    public String getPlantNk() { return plantNk; }
    public void setPlantNk(String v) { this.plantNk = v; }

    public boolean isJotaiUmMitsumori() { return jotaiUmMitsumori; }
    public void setJotaiUmMitsumori(boolean v) { this.jotaiUmMitsumori = v; }
    public boolean isJotaiUmKeiyaku() { return jotaiUmKeiyaku; }
    public void setJotaiUmKeiyaku(boolean v) { this.jotaiUmKeiyaku = v; }
    public boolean isJotaiUmKaiyaku() { return jotaiUmKaiyaku; }
    public void setJotaiUmKaiyaku(boolean v) { this.jotaiUmKaiyaku = v; }
    public boolean isJotaiUmHaki() { return jotaiUmHaki; }
    public void setJotaiUmHaki(boolean v) { this.jotaiUmHaki = v; }

    public boolean isJotaiTkMitsumori() { return jotaiTkMitsumori; }
    public void setJotaiTkMitsumori(boolean v) { this.jotaiTkMitsumori = v; }
    public boolean isJotaiTkKeiyaku() { return jotaiTkKeiyaku; }
    public void setJotaiTkKeiyaku(boolean v) { this.jotaiTkKeiyaku = v; }
    public boolean isJotaiTkKaiyaku() { return jotaiTkKaiyaku; }
    public void setJotaiTkKaiyaku(boolean v) { this.jotaiTkKaiyaku = v; }
    public boolean isJotaiTkHaki() { return jotaiTkHaki; }
    public void setJotaiTkHaki(boolean v) { this.jotaiTkHaki = v; }

    public List<Grid1RowForm> getGrid1Rows() { return grid1Rows; }
    public void setGrid1Rows(List<Grid1RowForm> v) { this.grid1Rows = v; }
    public List<Grid2RowForm> getGrid2Rows() { return grid2Rows; }
    public void setGrid2Rows(List<Grid2RowForm> v) { this.grid2Rows = v; }
    public List<Grid3RowForm> getGrid3Rows() { return grid3Rows; }
    public void setGrid3Rows(List<Grid3RowForm> v) { this.grid3Rows = v; }
}
