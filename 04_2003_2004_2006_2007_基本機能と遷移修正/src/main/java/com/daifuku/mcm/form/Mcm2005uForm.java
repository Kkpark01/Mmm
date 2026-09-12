package com.daifuku.mcm.form;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * 【変換元】Mcm2005uDelivery.vb / Mcm2005uTabControl.vb
 * MCM2005U 店舗見積ブランド詳細設定 フォーム
 *
 * Lombok不使用（プロジェクトルール）。
 */
public class Mcm2005uForm implements Serializable {

    private static final long serialVersionUID = 1L;

    private Mcm2003uForm header;
    public Mcm2003uForm getHeader(){return header;}
    public void setHeader(Mcm2003uForm value){header=value;}
    private String periodBiko;
    public String getPeriodBiko(){return periodBiko;}
    public void setPeriodBiko(String value){periodBiko=value;}
    private java.util.Map<String,Object> periodMeta;
    public java.util.Map<String,Object> getPeriodMeta(){return periodMeta;}
    public void setPeriodMeta(java.util.Map<String,Object> value){periodMeta=value;}
    private java.util.Map<String,Object> brandPeriodMeta;
    public java.util.Map<String,Object> getBrandPeriodMeta(){return brandPeriodMeta;}
    public void setBrandPeriodMeta(java.util.Map<String,Object> value){brandPeriodMeta=value;}
    private boolean headerEdited;
    public boolean isHeaderEdited(){return headerEdited;}
    public void setHeaderEdited(boolean value){headerEdited=value;}
    private String selectedKoseiId;
    public String getSelectedKoseiId(){return selectedKoseiId;}
    public void setSelectedKoseiId(String value){selectedKoseiId=value;}
    private String version;
    public String getVersion(){return version;}
    public void setVersion(String value){version=value;}
    private boolean dremosFlg;
    public boolean getDremosFlg(){return dremosFlg;}
    public void setDremosFlg(boolean value){dremosFlg=value;}
    private boolean remoteFlg;
    public boolean getRemoteFlg(){return remoteFlg;}
    public void setRemoteFlg(boolean value){remoteFlg=value;}

    // ===================================================================
    // エントリポイント識別子
    // ===================================================================

    /** MCM_UM_KIHON_BRAND.UM_KIHON_BRAND_ID */
    private BigDecimal umKihonBrandId;

    /** MCM_UM_MITSUMORI.UM_MITSUMORI_ID */
    private BigDecimal umMitsumoriId;

    /** MCM_UM_KIHON_MITSUMORI.UM_KIHON_MITSUMORI_ID (MCM2003U戻り先用) */
    private BigDecimal umKihonMitsumoriId;

    // ===================================================================
    // ブランド情報 (MCM_UM_KIHON_BRAND)
    // ===================================================================

    /** ブランド名称 */
    private String brandNk;

    /** ブランド詳細名称 */
    private String brandsyosaiNk;

    /** 契約時間帯 */
    private String keiyakujikantai;

    /** ハード保守費（基本：MCM_UM_KIHON_BRAND.HARDHOSYU_KIN） */
    private BigDecimal hardhosyuKinKihon;

    /** 補正ソフト保守費（MCM_UM_KIHON_BRAND.HOSEISOFTHOSHU_KIN）— HOSHU_KIN計算用 */
    private BigDecimal hoseisofthosyuKin;

    // ===================================================================
    // 期間情報 (MCM_UM_MITSUMORI)
    // ===================================================================

    /** 開始日 (yyyy/MM/dd) */
    private String kaisiDt;

    /** 終了日 (yyyy/MM/dd) */
    private String syuryoDt;

    // ===================================================================
    // 見積状態 (MCM_UM_KIHON_MITSUMORI)
    // ===================================================================

    /** 状態コード */
    private String jotai;

    /** 承認状態コード */
    private String syouninJotai;

    // ===================================================================
    // 見積条件 (MCM_UM_BRAND) — 編集対象
    // ===================================================================

    /** 見積注記 (MCM_UM_BRAND.MITSUMORICHUKI) */
    private String mitsumorichuki;

    /** ハード保守費 (MCM_UM_BRAND.HARDHOSYU_KIN) */
    private BigDecimal hardhosyuKin;

    /** 調整費 (MCM_UM_BRAND.CHOSEI_KIN) — ユーザー入力 */
    private BigDecimal choseiKin;

    /** 保守費 (MCM_UM_BRAND.HOSHU_KIN) — 計算値: hardhosyuKin + hoseisofthosyuKin - choseiKin */
    private BigDecimal hoshuKin;

    /** 備考 (MCM_UM_BRAND.BIKO) */
    private String biko;

    // ===================================================================
    // 機器単価グリッド (MCM_UM_TANKA JOIN MCM_UM_KIKIKOSEI JOIN MCM_UM_KIKIMEISAI)
    // ===================================================================

    /** 機器単価明細リスト */
    private List<TankaRowForm> tankaRows = new ArrayList<>();

    // -------------------------------------------------------------------
    // getters / setters
    // -------------------------------------------------------------------

    public BigDecimal getUmKihonBrandId() { return umKihonBrandId; }
    public void setUmKihonBrandId(BigDecimal umKihonBrandId) { this.umKihonBrandId = umKihonBrandId; }

    public BigDecimal getUmMitsumoriId() { return umMitsumoriId; }
    public void setUmMitsumoriId(BigDecimal umMitsumoriId) { this.umMitsumoriId = umMitsumoriId; }

    public BigDecimal getUmKihonMitsumoriId() { return umKihonMitsumoriId; }
    public void setUmKihonMitsumoriId(BigDecimal umKihonMitsumoriId) { this.umKihonMitsumoriId = umKihonMitsumoriId; }

    public String getBrandNk() { return brandNk; }
    public void setBrandNk(String brandNk) { this.brandNk = brandNk; }

    public String getBrandsyosaiNk() { return brandsyosaiNk; }
    public void setBrandsyosaiNk(String brandsyosaiNk) { this.brandsyosaiNk = brandsyosaiNk; }

    public String getKeiyakujikantai() { return keiyakujikantai; }
    public void setKeiyakujikantai(String keiyakujikantai) { this.keiyakujikantai = keiyakujikantai; }

    public BigDecimal getHardhosyuKinKihon() { return hardhosyuKinKihon; }
    public void setHardhosyuKinKihon(BigDecimal hardhosyuKinKihon) { this.hardhosyuKinKihon = hardhosyuKinKihon; }

    public BigDecimal getHoseisofthosyuKin() { return hoseisofthosyuKin; }
    public void setHoseisofthosyuKin(BigDecimal hoseisofthosyuKin) { this.hoseisofthosyuKin = hoseisofthosyuKin; }

    public String getKaisiDt() { return kaisiDt; }
    public void setKaisiDt(String kaisiDt) { this.kaisiDt = kaisiDt; }

    public String getSyuryoDt() { return syuryoDt; }
    public void setSyuryoDt(String syuryoDt) { this.syuryoDt = syuryoDt; }

    public String getJotai() { return jotai; }
    public void setJotai(String jotai) { this.jotai = jotai; }

    public String getSyouninJotai() { return syouninJotai; }
    public void setSyouninJotai(String syouninJotai) { this.syouninJotai = syouninJotai; }

    public String getMitsumorichuki() { return mitsumorichuki; }
    public void setMitsumorichuki(String mitsumorichuki) { this.mitsumorichuki = mitsumorichuki; }

    public BigDecimal getHardhosyuKin() { return hardhosyuKin; }
    public void setHardhosyuKin(BigDecimal hardhosyuKin) { this.hardhosyuKin = hardhosyuKin; }

    public BigDecimal getChoseiKin() { return choseiKin; }
    public void setChoseiKin(BigDecimal choseiKin) { this.choseiKin = choseiKin; }

    public BigDecimal getHoshuKin() { return hoshuKin; }
    public void setHoshuKin(BigDecimal hoshuKin) { this.hoshuKin = hoshuKin; }

    public String getBiko() { return biko; }
    public void setBiko(String biko) { this.biko = biko; }

    public List<TankaRowForm> getTankaRows() { return tankaRows; }
    public void setTankaRows(List<TankaRowForm> tankaRows) { this.tankaRows = tankaRows; }

    // ===================================================================
    // 機器単価行フォーム
    // 【変換元】MCM_UM_TANKACPDataGridView (Mcm2005uTabControl.vb)
    // MCM_UM_TANKA + MCM_UM_KIKIKOSEI (JOIN) + MCM_UM_KIKIMEISAI (LEFT JOIN)
    // ===================================================================

    public static class TankaRowForm implements Serializable {

        private static final long serialVersionUID = 1L;
    private String seizomakerNk;
    public String getSeizomakerNk(){return seizomakerNk;}
    public void setSeizomakerNk(String value){seizomakerNk=value;}
    private String kikikatashiki;
    public String getKikikatashiki(){return kikikatashiki;}
    public void setKikikatashiki(String value){kikikatashiki=value;}
    private String sosuNm;
    public String getSosuNm(){return sosuNm;}
    public void setSosuNm(String value){sosuNm=value;}
    private String createdDt;
    public String getCreatedDt(){return createdDt;}
    public void setCreatedDt(String value){createdDt=value;}
    private String createdBy;
    public String getCreatedBy(){return createdBy;}
    public void setCreatedBy(String value){createdBy=value;}
    private String lastupdateDt;
    public String getLastupdateDt(){return lastupdateDt;}
    public void setLastupdateDt(String value){lastupdateDt=value;}
    private String lastupdateBy;
    public String getLastupdateBy(){return lastupdateBy;}
    public void setLastupdateBy(String value){lastupdateBy=value;}
    private String torihosyujikanNk;
    public String getTorihosyujikanNk(){return torihosyujikanNk;}
    public void setTorihosyujikanNk(String value){torihosyujikanNk=value;}
    private String kikikoseiNk;
    public String getKikikoseiNk(){return kikikoseiNk;}
    public void setKikikoseiNk(String value){kikikoseiNk=value;}
    private BigDecimal sikiriKin;
    public BigDecimal getSikiriKin(){return sikiriKin;}
    public void setSikiriKin(BigDecimal value){sikiriKin=value;}
    private BigDecimal torihosyujikanId;
    public BigDecimal getTorihosyujikanId(){return torihosyujikanId;}
    public void setTorihosyujikanId(BigDecimal value){torihosyujikanId=value;}
    private BigDecimal daifukuhosyujikanId;
    public BigDecimal getDaifukuhosyujikanId(){return daifukuhosyujikanId;}
    public void setDaifukuhosyujikanId(BigDecimal value){daifukuhosyujikanId=value;}
    private BigDecimal tenkenkaisu;
    public BigDecimal getTenkenkaisu(){return tenkenkaisu;}
    public void setTenkenkaisu(BigDecimal value){tenkenkaisu=value;}
    private String tenkenyobi;
    public String getTenkenyobi(){return tenkenyobi;}
    public void setTenkenyobi(String value){tenkenyobi=value;}
    private String hosyuhoho;
    public String getHosyuhoho(){return hosyuhoho;}
    public void setHosyuhoho(String value){hosyuhoho=value;}
    private boolean packFlg;
    public boolean getPackFlg(){return packFlg;}
    public void setPackFlg(boolean value){packFlg=value;}
    private String keiyakunaiyo;
    public String getKeiyakunaiyo(){return keiyakunaiyo;}
    public void setKeiyakunaiyo(String value){keiyakunaiyo=value;}
    private String keiyakuNo;
    public String getKeiyakuNo(){return keiyakuNo;}
    public void setKeiyakuNo(String value){keiyakuNo=value;}
    private String biko;
    public String getBiko(){return biko;}
    public void setBiko(String value){biko=value;}


        /** MCM_UM_TANKA.UM_TANKA_ID */
        private BigDecimal umTankaId;

        /** MCM_UM_TANKA.UM_KIKIKOSEI_ID */
        private BigDecimal umKikoseiId;

        /** MCM_UM_KIKIKOSEI.SET_NM — 機器構成名称 */
        private String setNm;

        /** MCM_UM_KIKIMEISAI.MEISAI_NM — 機器明細名称 */
        private String meisaiNm;

        /** 選定数量 — 数量（表示のみ） */
        private String suryoNm;

        /** MCM_UM_TANKA.HYOJUN_KIN — 標準単価（参照専用） */
        private BigDecimal hyojunKin;

        /** 標準単価×選定数量 — 標準小計 = hyojunKin × suryoNm（計算値・表示） */
        private BigDecimal hyojunshokeiKin;

        /** 仕切単価×選定数量 — 仕切り小計（計算表示） */
        private BigDecimal sikirishokeiKin;

        /** 取扱機器マスタのCONTROLLER_FLG — コントローラ行フラグ */
        private String controllerFlg;

        // getters / setters

        public BigDecimal getUmTankaId() { return umTankaId; }
        public void setUmTankaId(BigDecimal umTankaId) { this.umTankaId = umTankaId; }

        public BigDecimal getUmKikoseiId() { return umKikoseiId; }
        public void setUmKikoseiId(BigDecimal umKikoseiId) { this.umKikoseiId = umKikoseiId; }

        public String getSetNm() { return setNm; }
        public void setSetNm(String setNm) { this.setNm = setNm; }

        public String getMeisaiNm() { return meisaiNm; }
        public void setMeisaiNm(String meisaiNm) { this.meisaiNm = meisaiNm; }

        public String getSuryoNm() { return suryoNm; }
        public void setSuryoNm(String suryoNm) { this.suryoNm = suryoNm; }

        public BigDecimal getHyojunKin() { return hyojunKin; }
        public void setHyojunKin(BigDecimal hyojunKin) { this.hyojunKin = hyojunKin; }

        public BigDecimal getHyojunshokeiKin() { return hyojunshokeiKin; }
        public void setHyojunshokeiKin(BigDecimal hyojunshokeiKin) { this.hyojunshokeiKin = hyojunshokeiKin; }

        public BigDecimal getSikirishokeiKin() { return sikirishokeiKin; }
        public void setSikirishokeiKin(BigDecimal sikirishokeiKin) { this.sikirishokeiKin = sikirishokeiKin; }

        public String getControllerFlg() { return controllerFlg; }
        public void setControllerFlg(String controllerFlg) { this.controllerFlg = controllerFlg; }
    }
}
