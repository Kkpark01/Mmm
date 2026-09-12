package com.daifuku.mcm.form;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * 【変換元】Mcm2007uScreen1.vb / Mcm2007uScreen2.vb
 * MCM2007U ユーザ契約内容変更フォーム（2ステップウィザード）
 *
 * Step1: 店舗見積（MCM_UM_MITSUMORI）+ 基本ブランド（MCM_UM_KIHON_BRAND）選択
 * Step2: 機器構成（MCM_MA_KIKIKOSEI）+ 機器明細（MCM_MA_KIKIMEISAI）+ 個体管理（MCM_MA_KIKIKOTAIKANRI）選択
 *
 * DB書き込みなし（SELECT専用画面）。
 * 「適用」ボタンで選択結果をセッションに格納しMCM2006Uへリダイレクト。
 */
public class Mcm2007uForm implements Serializable {

    private static final long serialVersionUID = 1L;

    // MCM2006Uから渡される情報
    private int seniMotoKbn;
    private BigDecimal ukKeiyakuId;
    private BigDecimal ukKikanId;    // 変更対象タブID（SENIMOTO_TAB_HENKO時）
    private BigDecimal plantId;
    private BigDecimal nonyusakiId;
    private String nonyusakiCd;
    private String nonyusakiNk;
    private String nonyusakijusyo1Nk;
    private String nonyusakijusyo2Nk;
    private String supportId;
    private String plantNk;

    // 表示中の見積。選択チェックと独立し、戻る・入力エラー後も保持する。
    private int selectedEstimateIndex = -1;
    public int getSelectedEstimateIndex(){return selectedEstimateIndex;}
    public void setSelectedEstimateIndex(int value){selectedEstimateIndex=value;}

    // Step1入力: 適用開始日・終了日
    private String kaisiDt;
    private String syuryoDt;

    // Step1選択行（MCM_UM_MITSUMORI = 店舗見積）
    private List<MitsumoriRowForm> mitsumoriRows = new ArrayList<>();

    private List<java.util.Map<String,String>> step2Brands = new ArrayList<>();
    public List<java.util.Map<String,String>> getStep2Brands(){return step2Brands;}
    public void setStep2Brands(List<java.util.Map<String,String>> value){step2Brands=value;}
    private int selectedBrandIndex = -1;
    private int selectedKoseiIndex = -1;
    public int getSelectedBrandIndex(){return selectedBrandIndex;}
    public void setSelectedBrandIndex(int value){selectedBrandIndex=value;}
    public int getSelectedKoseiIndex(){return selectedKoseiIndex;}
    public void setSelectedKoseiIndex(int value){selectedKoseiIndex=value;}

    // Step2選択行
    private List<KoseiRowForm> koseiRows   = new ArrayList<>();
    private List<MeisaiRowForm> meisaiRows = new ArrayList<>();
    private List<KotaiRowForm> kotaiRows   = new ArrayList<>();

    // ─── getter / setter ─────────────────────────────────────────────────

    public int getSeniMotoKbn() { return seniMotoKbn; }
    public void setSeniMotoKbn(int v) { this.seniMotoKbn = v; }
    public BigDecimal getUkKeiyakuId() { return ukKeiyakuId; }
    public void setUkKeiyakuId(BigDecimal v) { this.ukKeiyakuId = v; }
    public BigDecimal getUkKikanId() { return ukKikanId; }
    public void setUkKikanId(BigDecimal v) { this.ukKikanId = v; }
    public BigDecimal getPlantId() { return plantId; }
    public void setPlantId(BigDecimal v) { this.plantId = v; }
    public BigDecimal getNonyusakiId() { return nonyusakiId; }
    public void setNonyusakiId(BigDecimal v) { this.nonyusakiId = v; }
    public String getNonyusakiCd() { return nonyusakiCd; }
    public void setNonyusakiCd(String v) { this.nonyusakiCd = v; }
    public String getNonyusakiNk() { return nonyusakiNk; }
    public void setNonyusakiNk(String v) { this.nonyusakiNk = v; }
    public String getNonyusakijusyo1Nk() { return nonyusakijusyo1Nk; }
    public void setNonyusakijusyo1Nk(String v) { this.nonyusakijusyo1Nk = v; }
    public String getNonyusakijusyo2Nk() { return nonyusakijusyo2Nk; }
    public void setNonyusakijusyo2Nk(String v) { this.nonyusakijusyo2Nk = v; }
    public String getSupportId() { return supportId; }
    public void setSupportId(String v) { this.supportId = v; }
    public String getPlantNk() { return plantNk; }
    public void setPlantNk(String v) { this.plantNk = v; }
    public String getKaisiDt() { return kaisiDt; }
    public void setKaisiDt(String v) { this.kaisiDt = v; }
    public String getSyuryoDt() { return syuryoDt; }
    public void setSyuryoDt(String v) { this.syuryoDt = v; }

    public List<MitsumoriRowForm> getMitsumoriRows() { return mitsumoriRows; }
    public void setMitsumoriRows(List<MitsumoriRowForm> v) { this.mitsumoriRows = v; }
    public List<KoseiRowForm> getKoseiRows() { return koseiRows; }
    public void setKoseiRows(List<KoseiRowForm> v) { this.koseiRows = v; }
    public List<MeisaiRowForm> getMeisaiRows() { return meisaiRows; }
    public void setMeisaiRows(List<MeisaiRowForm> v) { this.meisaiRows = v; }
    public List<KotaiRowForm> getKotaiRows() { return kotaiRows; }
    public void setKotaiRows(List<KotaiRowForm> v) { this.kotaiRows = v; }

    // ═══════════════════════════════════════════════════════════════════════
    // 内部クラス: 店舗見積行 (MCM_UM_MITSUMORI + MCM_UM_KIHON_BRAND)
    // ═══════════════════════════════════════════════════════════════════════

    public static class MitsumoriRowForm implements Serializable {
        // DBから取得する参照表示。HTTP入力にはバインドしない。
        private java.util.Map<String,String> display = new java.util.LinkedHashMap<>();
        public java.util.Map<String,String> getDisplay(){return display;}
        public void setDisplay(java.util.Map<String,String> value){display=value;}
        private List<java.util.Map<String,String>> previewRows = new ArrayList<>();
        public List<java.util.Map<String,String>> getPreviewRows(){return previewRows;}
        public void setPreviewRows(List<java.util.Map<String,String>> value){previewRows=value;}
        public boolean isCheckFlg(){return brandRows.stream().anyMatch(KihonBrandRowForm::isCheckFlg);}

        private static final long serialVersionUID = 1L;

        // MCM_UM_MITSUMORI
        private BigDecimal umMitsumoriId;
        private String umMitsumoriNo;
        private String kaisiDt;
        private String syuryoDt;
        private String jotai;

        // MCM_UM_KIHON_BRAND (子: 1見積に複数ブランド)
        private List<KihonBrandRowForm> brandRows = new ArrayList<>();

        public BigDecimal getUmMitsumoriId() { return umMitsumoriId; }
        public void setUmMitsumoriId(BigDecimal v) { this.umMitsumoriId = v; }
        public String getUmMitsumoriNo() { return umMitsumoriNo; }
        public void setUmMitsumoriNo(String v) { this.umMitsumoriNo = v; }
        public String getKaisiDt() { return kaisiDt; }
        public void setKaisiDt(String v) { this.kaisiDt = v; }
        public String getSyuryoDt() { return syuryoDt; }
        public void setSyuryoDt(String v) { this.syuryoDt = v; }
        public String getJotai() { return jotai; }
        public void setJotai(String v) { this.jotai = v; }
        public List<KihonBrandRowForm> getBrandRows() { return brandRows; }
        public void setBrandRows(List<KihonBrandRowForm> v) { this.brandRows = v; }
    }

    public static class KihonBrandRowForm implements Serializable {
        // DBから取得する参照表示。HTTP入力にはバインドしない。
        private java.util.Map<String,String> display = new java.util.LinkedHashMap<>();
        public java.util.Map<String,String> getDisplay(){return display;}
        public void setDisplay(java.util.Map<String,String> value){display=value;}

        private BigDecimal hoshuKin;
        public BigDecimal getHoshuKin(){return hoshuKin;}
        public void setHoshuKin(BigDecimal value){hoshuKin=value;}
        private static final long serialVersionUID = 1L;

        private BigDecimal umKihonBrandId;
        private BigDecimal umMitsumoriId;
        private BigDecimal brandkoseiId;    // MCM_MA_BRAND_KOSEI.BRANDKOSEI_ID
        private String brandNk;
        private String brandsyosaiNk;
        private String keiyakujikantai;
        private String hosyuhoho;
        private boolean checkFlg;           // チェックボックス状態

        public BigDecimal getUmKihonBrandId() { return umKihonBrandId; }
        public void setUmKihonBrandId(BigDecimal v) { this.umKihonBrandId = v; }
        public BigDecimal getUmMitsumoriId() { return umMitsumoriId; }
        public void setUmMitsumoriId(BigDecimal v) { this.umMitsumoriId = v; }
        public BigDecimal getBrandkoseiId() { return brandkoseiId; }
        public void setBrandkoseiId(BigDecimal v) { this.brandkoseiId = v; }
        public String getBrandNk() { return brandNk; }
        public void setBrandNk(String v) { this.brandNk = v; }
        public String getBrandsyosaiNk() { return brandsyosaiNk; }
        public void setBrandsyosaiNk(String v) { this.brandsyosaiNk = v; }
        public String getKeiyakujikantai() { return keiyakujikantai; }
        public void setKeiyakujikantai(String v) { this.keiyakujikantai = v; }
        public String getHosyuhoho() { return hosyuhoho; }
        public void setHosyuhoho(String v) { this.hosyuhoho = v; }
        public boolean isCheckFlg() { return checkFlg; }
        public void setCheckFlg(boolean v) { this.checkFlg = v; }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 内部クラス: 機器構成行 (MCM_MA_KIKIKOSEI)
    // ═══════════════════════════════════════════════════════════════════════

    public static class KoseiRowForm implements Serializable {
        // 選定元の表示・所属情報。HTTP入力にはバインドしない。
        private java.util.Map<String,String> display = new java.util.LinkedHashMap<>();
        private boolean quoted;
        public java.util.Map<String,String> getDisplay(){return display;}
        public void setDisplay(java.util.Map<String,String> value){display=value;}
        public boolean isQuoted(){return quoted;}
        public void setQuoted(boolean value){quoted=value;}

        private static final long serialVersionUID = 1L;

        private BigDecimal kikikoseiId;     // MCM_MA_KIKIKOSEI.KIKIKOSEI_ID
        private BigDecimal brandkoseiId;
        private BigDecimal umKihonBrandId;  // 参照元MCM_UM_KIHON_BRAND
        private String kikikoseiNk;
        private String setNm;
        private String tani;
        private String tehaiseiban;
        private String controllerFlg;
        private String hosyuhoho;
        private int hyojijun;
        private boolean checkFlg;

        public BigDecimal getKikikoseiId() { return kikikoseiId; }
        public void setKikikoseiId(BigDecimal v) { this.kikikoseiId = v; }
        public BigDecimal getBrandkoseiId() { return brandkoseiId; }
        public void setBrandkoseiId(BigDecimal v) { this.brandkoseiId = v; }
        public BigDecimal getUmKihonBrandId() { return umKihonBrandId; }
        public void setUmKihonBrandId(BigDecimal v) { this.umKihonBrandId = v; }
        public String getKikikoseiNk() { return kikikoseiNk; }
        public void setKikikoseiNk(String v) { this.kikikoseiNk = v; }
        public String getSetNmText(){return setNm==null||setNm.isBlank()?"":new BigDecimal(setNm).stripTrailingZeros().toPlainString();}
        public String getSetNm() { return setNm; }
        public void setSetNm(String v) { this.setNm = v; }
        public String getTani() { return tani; }
        public void setTani(String v) { this.tani = v; }
        public String getTehaiseiban() { return tehaiseiban; }
        public void setTehaiseiban(String v) { this.tehaiseiban = v; }
        public String getControllerFlg() { return controllerFlg; }
        public void setControllerFlg(String v) { this.controllerFlg = v; }
        public String getHosyuhoho() { return hosyuhoho; }
        public void setHosyuhoho(String v) { this.hosyuhoho = v; }
        public int getHyojijun() { return hyojijun; }
        public void setHyojijun(int v) { this.hyojijun = v; }
        public boolean isCheckFlg() { return checkFlg; }
        public void setCheckFlg(boolean v) { this.checkFlg = v; }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 内部クラス: 機器明細行 (MCM_MA_KIKIMEISAI)
    // ═══════════════════════════════════════════════════════════════════════

    public static class MeisaiRowForm implements Serializable {
        // 選定元の表示・所属情報。HTTP入力にはバインドしない。
        private java.util.Map<String,String> display = new java.util.LinkedHashMap<>();
        private boolean quoted;
        public java.util.Map<String,String> getDisplay(){return display;}
        public void setDisplay(java.util.Map<String,String> value){display=value;}
        public boolean isQuoted(){return quoted;}
        public void setQuoted(boolean value){quoted=value;}

        private static final long serialVersionUID = 1L;

        private BigDecimal kikimeisaiId;
        private BigDecimal kikikoseiId;
        private BigDecimal brandkoseiId;
        private BigDecimal seizomakerId;
        private String seizomakerNk;
        private String kikihinmeiNk;
        private String kikikatashiki;
        private String suryoNm;
        private String keiyakunaiyo;
        private String keiyakuNo;
        private String servicekeitai;
        private BigDecimal torihosyujikanId;
        private BigDecimal daifukuhosyujikanId;
        private String tenkenkaisu;
        private String tenkenyobi;
        private String hosyuhoho;
        private int maeHyojijun;
        private int hyojijun;
        private boolean checkFlg;

        public BigDecimal getKikimeisaiId() { return kikimeisaiId; }
        public void setKikimeisaiId(BigDecimal v) { this.kikimeisaiId = v; }
        public BigDecimal getKikikoseiId() { return kikikoseiId; }
        public void setKikikoseiId(BigDecimal v) { this.kikikoseiId = v; }
        public BigDecimal getBrandkoseiId() { return brandkoseiId; }
        public void setBrandkoseiId(BigDecimal v) { this.brandkoseiId = v; }
        public BigDecimal getSeizomakerId() { return seizomakerId; }
        public void setSeizomakerId(BigDecimal v) { this.seizomakerId = v; }
        public String getSeizomankerNk() { return seizomakerNk; }
        public void setSeizomankerNk(String v) { this.seizomakerNk = v; }
        public String getKikihinmeiNk() { return kikihinmeiNk; }
        public void setKikihinmeiNk(String v) { this.kikihinmeiNk = v; }
        public String getKikikatashiki() { return kikikatashiki; }
        public void setKikikatashiki(String v) { this.kikikatashiki = v; }
        public String getSuryoNmText(){return suryoNm==null||suryoNm.isBlank()?"":new BigDecimal(suryoNm).stripTrailingZeros().toPlainString();}
        public String getSuryoNm() { return suryoNm; }
        public void setSuryoNm(String v) { this.suryoNm = v; }
        public String getKeiyakunaiyo() { return keiyakunaiyo; }
        public void setKeiyakunaiyo(String v) { this.keiyakunaiyo = v; }
        public String getKeiyakuNo() { return keiyakuNo; }
        public void setKeiyakuNo(String v) { this.keiyakuNo = v; }
        public String getServicekeitai() { return servicekeitai; }
        public void setServicekeitai(String v) { this.servicekeitai = v; }
        public BigDecimal getTorihosyujikanId() { return torihosyujikanId; }
        public void setTorihosyujikanId(BigDecimal v) { this.torihosyujikanId = v; }
        public BigDecimal getDaifukuhosyujikanId() { return daifukuhosyujikanId; }
        public void setDaifukuhosyujikanId(BigDecimal v) { this.daifukuhosyujikanId = v; }
        public String getTenkenkaisu() { return tenkenkaisu; }
        public void setTenkenkaisu(String v) { this.tenkenkaisu = v; }
        public String getTenkenyobi() { return tenkenyobi; }
        public void setTenkenyobi(String v) { this.tenkenyobi = v; }
        public String getHosyuhoho() { return hosyuhoho; }
        public void setHosyuhoho(String v) { this.hosyuhoho = v; }
        public int getMaeHyojijun() { return maeHyojijun; }
        public void setMaeHyojijun(int v) { this.maeHyojijun = v; }
        public int getHyojijun() { return hyojijun; }
        public void setHyojijun(int v) { this.hyojijun = v; }
        public boolean isCheckFlg() { return checkFlg; }
        public void setCheckFlg(boolean v) { this.checkFlg = v; }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 内部クラス: 個体管理行 (MCM_MA_KIKIKOTAIKANRI)
    // ═══════════════════════════════════════════════════════════════════════

    public static class KotaiRowForm implements Serializable {
        // 選定元の表示・所属情報。HTTP入力にはバインドしない。
        private java.util.Map<String,String> display = new java.util.LinkedHashMap<>();
        private boolean quoted;
        public java.util.Map<String,String> getDisplay(){return display;}
        public void setDisplay(java.util.Map<String,String> value){display=value;}
        public boolean isQuoted(){return quoted;}
        public void setQuoted(boolean value){quoted=value;}

        private static final long serialVersionUID = 1L;

        private BigDecimal kotaikanriId;
        private BigDecimal kikimeisaiId;
        private BigDecimal kikikoseiId;
        private BigDecimal brandkoseiId;
        private String kotaiNk;
        private String serialNo;
        private String itizinonnyuDt;
        private String setchibasyo;
        private boolean checkFlg;

        public BigDecimal getKotaikanriId() { return kotaikanriId; }
        public void setKotaikanriId(BigDecimal v) { this.kotaikanriId = v; }
        public BigDecimal getKikimeisaiId() { return kikimeisaiId; }
        public void setKikimeisaiId(BigDecimal v) { this.kikimeisaiId = v; }
        public BigDecimal getKikikoseiId() { return kikikoseiId; }
        public void setKikikoseiId(BigDecimal v) { this.kikikoseiId = v; }
        public BigDecimal getBrandkoseiId() { return brandkoseiId; }
        public void setBrandkoseiId(BigDecimal v) { this.brandkoseiId = v; }
        public String getKotaiNk() { return kotaiNk; }
        public void setKotaiNk(String v) { this.kotaiNk = v; }
        public String getSerialNo() { return serialNo; }
        public void setSerialNo(String v) { this.serialNo = v; }
        public String getItizinonnyuDt() { return itizinonnyuDt; }
        public void setItizinonnyuDt(String v) { this.itizinonnyuDt = v; }
        public String getSetchibasyo() { return setchibasyo; }
        public void setSetchibasyo(String v) { this.setchibasyo = v; }
        public boolean isCheckFlg() { return checkFlg; }
        public void setCheckFlg(boolean v) { this.checkFlg = v; }
    }
}
