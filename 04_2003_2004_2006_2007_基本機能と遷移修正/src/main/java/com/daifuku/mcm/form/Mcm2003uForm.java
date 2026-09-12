package com.daifuku.mcm.form;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * 【変換元】Mcm2003uDelivery.vb / Mcm2003uScreen.vb
 * MCM2003U 店舗見積内容基本設定フォーム
 */
public class Mcm2003uForm implements Serializable {

    private static final long serialVersionUID = 1L;

    // ===================================================================
    // セッション受け渡し情報
    // 【変換元】Mcm2003uDelivery.vb
    // ===================================================================

    private BigDecimal umKihonMitsumoriId;
    private int seniMotoKbn;

    // ===================================================================
    // ヘッダー情報（MCM_UM_KIHON_MITSUMORI）
    // ===================================================================

    private String umMitsumoriNo;
    private String nonyusakiId;
    private String nonyusakiCd;
    private String nonyusakiNk;
    private String nonyusakijusyo1Nk;
    private String nonyusakijusyo2Nk;
    private String plantId;
    private String supportId;
    private String plantNk;
    private String nonyubusyoNk;
    private String nonyutantosyaNk;
    private String nonyutelNo;
    private String nonyufaxNo;
    private String mitsumoriDt;
    private String mitsumorisakuseisyaNk;
    private String mitsumorigiken;
    private BigDecimal iraitenpoId;
    private String iraimeisho1Nk;
    private String iraimeisho2Nk;
    private String iraimeisho3Nk;
    private String iraimeisho4Nk;
    private String iraitenporyakuNk;
    private String iraitantoNk;
    private BigDecimal sofutenpoId;
    private String sofumeisho1Nk;
    private String sofumeisho2Nk;
    private String sofumeisho3Nk;
    private String sofumeisho4Nk;
    private String sofutenporyakuNk;
    private String sofutantoNk;
    private String keiyakujikantai;
    private String hosyuhoho;
    private String jotai;
    private String syouninJotai;
    private String biko;
    private String mitsumoriJouken;
    private String mitsumoriLevel;
    private String createdDt;
    private String createdBy;
    private String lastupdateDt;
    private String lastupdateBy;

    // ===================================================================
    // グリッドデータ
    // ===================================================================

    private List<MitsumoriRowForm> mitsumoriRows = new ArrayList<>();
    private List<TenpuRowForm> tenpuRows = new ArrayList<>();
    private List<BrandRowForm> brandRows = new ArrayList<>();
    private List<KoseiRowForm> koseiRows = new ArrayList<>();
    private List<MeisaiRowForm> meisaiRows = new ArrayList<>();

    // ===================================================================
    // 内部クラス: 契約金額変動時期（MCM_UM_MITSUMORI）
    // ===================================================================

    public static class MitsumoriRowForm implements Serializable {
        private static final long serialVersionUID = 1L;
        private BigDecimal umMitsumoriId;
        private BigDecimal umKihonMitsumoriId;
        private String kaisiDt;
        private String syuryoDt;
        private BigDecimal hardhosyuKin;
        private BigDecimal mitsumoriGkin;
        private boolean checkFlg;

        public BigDecimal getUmMitsumoriId() { return umMitsumoriId; }
        public void setUmMitsumoriId(BigDecimal v) { umMitsumoriId = v; }
        public BigDecimal getUmKihonMitsumoriId() { return umKihonMitsumoriId; }
        public void setUmKihonMitsumoriId(BigDecimal v) { umKihonMitsumoriId = v; }
        public String getKaisiDt() { return kaisiDt; }
        public void setKaisiDt(String v) { kaisiDt = v; }
        public String getSyuryoDt() { return syuryoDt; }
        public void setSyuryoDt(String v) { syuryoDt = v; }
        public BigDecimal getHardhosyuKin() { return hardhosyuKin; }
        public void setHardhosyuKin(BigDecimal v) { hardhosyuKin = v; }
        public BigDecimal getMitsumoriGkin() { return mitsumoriGkin; }
        public void setMitsumoriGkin(BigDecimal v) { mitsumoriGkin = v; }
        public boolean isCheckFlg() { return checkFlg; }
        public void setCheckFlg(boolean v) { checkFlg = v; }
    }

    // ===================================================================
    // 内部クラス: 添付ファイル（MCM_UM_TENPU）
    // ===================================================================

    public static class TenpuRowForm implements Serializable {
        private static final long serialVersionUID = 1L;
        private BigDecimal umTenpuId;
        private BigDecimal umKihonMitsumoriId;
        private String tenpufileNk;
        private String directory;
        private String biko;

        public BigDecimal getUmTenpuId() { return umTenpuId; }
        public void setUmTenpuId(BigDecimal v) { umTenpuId = v; }
        public BigDecimal getUmKihonMitsumoriId() { return umKihonMitsumoriId; }
        public void setUmKihonMitsumoriId(BigDecimal v) { umKihonMitsumoriId = v; }
        public String getTenpufileNk() { return tenpufileNk; }
        public void setTenpufileNk(String v) { tenpufileNk = v; }
        public String getDirectory() { return directory; }
        public void setDirectory(String v) { directory = v; }
        public String getBiko() { return biko; }
        public void setBiko(String v) { biko = v; }
    }

    // ===================================================================
    // 内部クラス: ブランドタブ（MCM_UM_KIHON_BRAND）
    // ===================================================================

    public static class BrandRowForm implements Serializable {
        private final java.util.Map<String,Object> editedOriginals = new java.util.LinkedHashMap<>();
        public java.util.Map<String,Object> getEditedOriginals() { return editedOriginals; }
        private static final long serialVersionUID = 1L;
        private BigDecimal umKihonBrandId;
        private BigDecimal umKihonMitsumoriId;
        private String brandNk;
        private String brandsyosaiNk;
        private String keiyakujikantai;
        private BigDecimal hardhosyuKin;
        private int hyojijun;
        // 【変換元】MCM_UM_KIHON_BRAND 追加列（システムサポート費/ソフトウェア保守費/補正/契約フラグ）
        private BigDecimal systemSekkeiKin;
        private BigDecimal kihonSekkeiKin;
        private BigDecimal programSakuseiKin;
        private BigDecimal systemSupportKin;
        private BigDecimal dtsSupportKin;
        private BigDecimal daifukuGijutsuKin;
        private BigDecimal softHosyuKin;
        private BigDecimal hoseiSystemSupportKin;
        private BigDecimal hoseiDtsSupportKin;
        private BigDecimal hoseiDaifukuGijutsuKin;
        private BigDecimal hoseiSoftHosyuKin;
        private String softHosyuhoho;
        private boolean softHosyuhohoEdited;
        private String originalSoftHosyuhoho;
        public String getOriginalSoftHosyuhoho() { return originalSoftHosyuhoho; }
        public void setOriginalSoftHosyuhoho(String value) { originalSoftHosyuhoho = value; }
        public boolean isSoftHosyuhohoEdited() { return softHosyuhohoEdited; }
        public void setSoftHosyuhohoEdited(boolean value) { softHosyuhohoEdited = value; }
        private boolean softFlg;
        private boolean dremosFlg;
        private boolean remoteFlg;
        private BigDecimal systemSupportSyokeiKin;
        private String createdDt;
        private String createdBy;
        private String lastupdateDt;
        private String lastupdateBy;

        public BigDecimal getUmKihonBrandId() { return umKihonBrandId; }
        public void setUmKihonBrandId(BigDecimal v) { umKihonBrandId = v; }
        public BigDecimal getUmKihonMitsumoriId() { return umKihonMitsumoriId; }
        public void setUmKihonMitsumoriId(BigDecimal v) { umKihonMitsumoriId = v; }
        public String getBrandNk() { return brandNk; }
        public void setBrandNk(String v) { brandNk = v; }
        public String getBrandsyosaiNk() { return brandsyosaiNk; }
        public void setBrandsyosaiNk(String v) { brandsyosaiNk = v; }
        public String getKeiyakujikantai() { return keiyakujikantai; }
        public void setKeiyakujikantai(String v) { keiyakujikantai = v; }
        public BigDecimal getHardhosyuKin() { return hardhosyuKin; }
        public void setHardhosyuKin(BigDecimal v) { hardhosyuKin = v; }
        public int getHyojijun() { return hyojijun; }
        public void setHyojijun(int v) { hyojijun = v; }
        public BigDecimal getSystemSekkeiKin() { return systemSekkeiKin; }
        public void setSystemSekkeiKin(BigDecimal v) { systemSekkeiKin = v; }
        public BigDecimal getKihonSekkeiKin() { return kihonSekkeiKin; }
        public void setKihonSekkeiKin(BigDecimal v) { kihonSekkeiKin = v; }
        public BigDecimal getProgramSakuseiKin() { return programSakuseiKin; }
        public void setProgramSakuseiKin(BigDecimal v) { programSakuseiKin = v; }
        public BigDecimal getSystemSupportKin() { return systemSupportKin; }
        public void setSystemSupportKin(BigDecimal v) { systemSupportKin = v; }
        public BigDecimal getDtsSupportKin() { return dtsSupportKin; }
        public void setDtsSupportKin(BigDecimal v) { dtsSupportKin = v; }
        public BigDecimal getDaifukuGijutsuKin() { return daifukuGijutsuKin; }
        public void setDaifukuGijutsuKin(BigDecimal v) { daifukuGijutsuKin = v; }
        public BigDecimal getSoftHosyuKin() { return softHosyuKin; }
        public void setSoftHosyuKin(BigDecimal v) { softHosyuKin = v; }
        public BigDecimal getHoseiSystemSupportKin() { return hoseiSystemSupportKin; }
        public void setHoseiSystemSupportKin(BigDecimal v) { hoseiSystemSupportKin = v; }
        public BigDecimal getHoseiDtsSupportKin() { return hoseiDtsSupportKin; }
        public void setHoseiDtsSupportKin(BigDecimal v) { hoseiDtsSupportKin = v; }
        public BigDecimal getHoseiDaifukuGijutsuKin() { return hoseiDaifukuGijutsuKin; }
        public void setHoseiDaifukuGijutsuKin(BigDecimal v) { hoseiDaifukuGijutsuKin = v; }
        public BigDecimal getHoseiSoftHosyuKin() { return hoseiSoftHosyuKin; }
        public void setHoseiSoftHosyuKin(BigDecimal v) { hoseiSoftHosyuKin = v; }
        public String getSoftHosyuhoho() { return softHosyuhoho; }
        public void setSoftHosyuhoho(String v) { softHosyuhoho = v; }
        public boolean isSoftFlg() { return softFlg; }
        public void setSoftFlg(boolean v) { softFlg = v; }
        public boolean isDremosFlg() { return dremosFlg; }
        public void setDremosFlg(boolean v) { dremosFlg = v; }
        public boolean isRemoteFlg() { return remoteFlg; }
        public void setRemoteFlg(boolean v) { remoteFlg = v; }
        public BigDecimal getSystemSupportSyokeiKin() { return systemSupportSyokeiKin; }
        public void setSystemSupportSyokeiKin(BigDecimal v) { systemSupportSyokeiKin = v; }
        public String getCreatedDt() { return createdDt; }
        public void setCreatedDt(String v) { createdDt = v; }
        public String getCreatedBy() { return createdBy; }
        public void setCreatedBy(String v) { createdBy = v; }
        public String getLastupdateDt() { return lastupdateDt; }
        public void setLastupdateDt(String v) { lastupdateDt = v; }
        public String getLastupdateBy() { return lastupdateBy; }
        public void setLastupdateBy(String v) { lastupdateBy = v; }
    }

    // ===================================================================
    // 内部クラス: 機器構成パターン（MCM_UM_KIKIKOSEI）
    // ===================================================================

    public static class KoseiRowForm implements Serializable {
        private final java.util.Map<String,Object> editedOriginals = new java.util.LinkedHashMap<>();
        public java.util.Map<String,Object> getEditedOriginals() { return editedOriginals; }
        private static final long serialVersionUID = 1L;
        private BigDecimal umKikikoseiId;
        private BigDecimal umKihonBrandId;
        private String kikikoseiNk;
        private BigDecimal setNm;
        private String tani;
        private String tehaiseiban;
        private String hosyuhoho;
        private String biko;
        private int hyojijun;

        public BigDecimal getUmKikikoseiId() { return umKikikoseiId; }
        public void setUmKikikoseiId(BigDecimal v) { umKikikoseiId = v; }
        public BigDecimal getUmKihonBrandId() { return umKihonBrandId; }
        public void setUmKihonBrandId(BigDecimal v) { umKihonBrandId = v; }
        public String getKikikoseiNk() { return kikikoseiNk; }
        public void setKikikoseiNk(String v) { kikikoseiNk = v; }
        public BigDecimal getSetNm() { return setNm; }
        public void setSetNm(BigDecimal v) { setNm = v; }
        public String getTani() { return tani; }
        public void setTani(String v) { tani = v; }
        public String getTehaiseiban() { return tehaiseiban; }
        public void setTehaiseiban(String v) { tehaiseiban = v; }
        public String getHosyuhoho() { return hosyuhoho; }
        public void setHosyuhoho(String v) { hosyuhoho = v; }
        public String getBiko() { return biko; }
        public void setBiko(String v) { biko = v; }
        public int getHyojijun() { return hyojijun; }
        public void setHyojijun(int v) { hyojijun = v; }
    }

    // ===================================================================
    // 内部クラス: 機器明細（MCM_UM_KIKIMEISAI）
    // ===================================================================

    public static class MeisaiRowForm implements Serializable {
        private static final long serialVersionUID = 1L;
        private BigDecimal umKikimeisaiId;
        private BigDecimal umKikikoseiId;
        private String seizomakerNk;
        private String kikihinmeiNk;
        private String kikikatashiki;
        private BigDecimal suryoNm;
        private BigDecimal sosuNm;
        private String biko;

        public BigDecimal getUmKikimeisaiId() { return umKikimeisaiId; }
        public void setUmKikimeisaiId(BigDecimal v) { umKikimeisaiId = v; }
        public BigDecimal getUmKikikoseiId() { return umKikikoseiId; }
        public void setUmKikikoseiId(BigDecimal v) { umKikikoseiId = v; }
        public String getSeizomakerNk() { return seizomakerNk; }
        public void setSeizomakerNk(String v) { seizomakerNk = v; }
        public String getKikihinmeiNk() { return kikihinmeiNk; }
        public void setKikihinmeiNk(String v) { kikihinmeiNk = v; }
        public String getKikikatashiki() { return kikikatashiki; }
        public void setKikikatashiki(String v) { kikikatashiki = v; }
        public BigDecimal getSuryoNm() { return suryoNm; }
        public void setSuryoNm(BigDecimal v) { suryoNm = v; }
        public BigDecimal getSosuNm() { return sosuNm; }
        public void setSosuNm(BigDecimal v) { sosuNm = v; }
        public String getBiko() { return biko; }
        public void setBiko(String v) { biko = v; }
    }

    // ===================================================================
    // getter / setter
    // ===================================================================

    public BigDecimal getUmKihonMitsumoriId() { return umKihonMitsumoriId; }
    public void setUmKihonMitsumoriId(BigDecimal v) { umKihonMitsumoriId = v; }
    public int getSeniMotoKbn() { return seniMotoKbn; }
    public void setSeniMotoKbn(int v) { seniMotoKbn = v; }

    public String getUmMitsumoriNo() { return umMitsumoriNo; }
    public void setUmMitsumoriNo(String v) { umMitsumoriNo = v; }
    public String getNonyusakiId() { return nonyusakiId; }
    public void setNonyusakiId(String v) { nonyusakiId = v; }
    public String getNonyusakiCd() { return nonyusakiCd; }
    public void setNonyusakiCd(String v) { nonyusakiCd = v; }
    public String getNonyusakiNk() { return nonyusakiNk; }
    public void setNonyusakiNk(String v) { nonyusakiNk = v; }
    public String getNonyusakijusyo1Nk() { return nonyusakijusyo1Nk; }
    public void setNonyusakijusyo1Nk(String v) { nonyusakijusyo1Nk = v; }
    public String getNonyusakijusyo2Nk() { return nonyusakijusyo2Nk; }
    public void setNonyusakijusyo2Nk(String v) { nonyusakijusyo2Nk = v; }
    public String getPlantId() { return plantId; }
    public void setPlantId(String v) { plantId = v; }
    public String getSupportId() { return supportId; }
    public void setSupportId(String v) { supportId = v; }
    public String getPlantNk() { return plantNk; }
    public void setPlantNk(String v) { plantNk = v; }
    public String getNonyubusyoNk() { return nonyubusyoNk; }
    public void setNonyubusyoNk(String v) { nonyubusyoNk = v; }
    public String getNonyutantosyaNk() { return nonyutantosyaNk; }
    public void setNonyutantosyaNk(String v) { nonyutantosyaNk = v; }
    public String getNonyutelNo() { return nonyutelNo; }
    public void setNonyutelNo(String v) { nonyutelNo = v; }
    public String getNonyufaxNo() { return nonyufaxNo; }
    public void setNonyufaxNo(String v) { nonyufaxNo = v; }
    public String getMitsumoriDt() { return mitsumoriDt; }
    public void setMitsumoriDt(String v) { mitsumoriDt = v; }
    public String getMitsumorisakuseisyaNk() { return mitsumorisakuseisyaNk; }
    public void setMitsumorisakuseisyaNk(String v) { mitsumorisakuseisyaNk = v; }
    public String getMitsumorigiken() { return mitsumorigiken; }
    public void setMitsumorigiken(String v) { mitsumorigiken = v; }
    public BigDecimal getIraitenpoId() { return iraitenpoId; }
    public void setIraitenpoId(BigDecimal v) { iraitenpoId = v; }
    public String getIraimeisho1Nk() { return iraimeisho1Nk; }
    public void setIraimeisho1Nk(String v) { iraimeisho1Nk = v; }
    public String getIraimeisho2Nk() { return iraimeisho2Nk; }
    public void setIraimeisho2Nk(String v) { iraimeisho2Nk = v; }
    public String getIraimeisho3Nk() { return iraimeisho3Nk; }
    public void setIraimeisho3Nk(String v) { iraimeisho3Nk = v; }
    public String getIraimeisho4Nk() { return iraimeisho4Nk; }
    public void setIraimeisho4Nk(String v) { iraimeisho4Nk = v; }
    public String getIraitenporyakuNk() { return iraitenporyakuNk; }
    public void setIraitenporyakuNk(String v) { iraitenporyakuNk = v; }
    public String getIraitantoNk() { return iraitantoNk; }
    public void setIraitantoNk(String v) { iraitantoNk = v; }
    public BigDecimal getSofutenpoId() { return sofutenpoId; }
    public void setSofutenpoId(BigDecimal v) { sofutenpoId = v; }
    public String getSofumeisho1Nk() { return sofumeisho1Nk; }
    public void setSofumeisho1Nk(String v) { sofumeisho1Nk = v; }
    public String getSofumeisho2Nk() { return sofumeisho2Nk; }
    public void setSofumeisho2Nk(String v) { sofumeisho2Nk = v; }
    public String getSofumeisho3Nk() { return sofumeisho3Nk; }
    public void setSofumeisho3Nk(String v) { sofumeisho3Nk = v; }
    public String getSofumeisho4Nk() { return sofumeisho4Nk; }
    public void setSofumeisho4Nk(String v) { sofumeisho4Nk = v; }
    public String getSofutenporyakuNk() { return sofutenporyakuNk; }
    public void setSofutenporyakuNk(String v) { sofutenporyakuNk = v; }
    public String getSofutantoNk() { return sofutantoNk; }
    public void setSofutantoNk(String v) { sofutantoNk = v; }
    public String getKeiyakujikantai() { return keiyakujikantai; }
    public void setKeiyakujikantai(String v) { keiyakujikantai = v; }
    public String getHosyuhoho() { return hosyuhoho; }
    public void setHosyuhoho(String v) { hosyuhoho = v; }
    public String getJotai() { return jotai; }
    public void setJotai(String v) { jotai = v; }
    public String getSyouninJotai() { return syouninJotai; }
    public void setSyouninJotai(String v) { syouninJotai = v; }
    public String getBiko() { return biko; }
    public void setBiko(String v) { biko = v; }
    public String getMitsumoriJouken() { return mitsumoriJouken; }
    public void setMitsumoriJouken(String v) { mitsumoriJouken = v; }
    public String getMitsumoriLevel() { return mitsumoriLevel; }
    public void setMitsumoriLevel(String v) { mitsumoriLevel = v; }
    public String getCreatedDt() { return createdDt; }
    public void setCreatedDt(String v) { createdDt = v; }
    public String getCreatedBy() { return createdBy; }
    public void setCreatedBy(String v) { createdBy = v; }
    public String getLastupdateDt() { return lastupdateDt; }
    public void setLastupdateDt(String v) { lastupdateDt = v; }
    public String getLastupdateBy() { return lastupdateBy; }
    public void setLastupdateBy(String v) { lastupdateBy = v; }

    /**
     * 依頼元店舗・事業所（表示用結合文字列）
     * 【変換元】Mcm2003uScreen.vb — IRAIMEISHO1_NK〜4_NKの結合表示
     */
    public String getIraiKaishaMeisho() {
        return joinNonBlank(iraimeisho1Nk, iraimeisho2Nk, iraimeisho3Nk, iraimeisho4Nk);
    }

    /**
     * 送付店舗・事業所（表示用結合文字列）
     * 【変換元】Mcm2003uScreen.vb — SOFUMEISHO1_NK〜4_NKの結合表示
     */
    public String getSofuKaishaMeisho() {
        return joinNonBlank(sofumeisho1Nk, sofumeisho2Nk, sofumeisho3Nk, sofumeisho4Nk);
    }

    private static String joinNonBlank(String... values) {
        StringBuilder sb = new StringBuilder();
        for (String v : values) {
            if (v != null && !v.trim().isEmpty()) {
                if (sb.length() > 0) {
                    sb.append(' ');
                }
                sb.append(v.trim());
            }
        }
        return sb.toString();
    }

    public List<MitsumoriRowForm> getMitsumoriRows() { return mitsumoriRows; }
    public void setMitsumoriRows(List<MitsumoriRowForm> v) { mitsumoriRows = v; }
    public List<TenpuRowForm> getTenpuRows() { return tenpuRows; }
    public void setTenpuRows(List<TenpuRowForm> v) { tenpuRows = v; }
    public List<BrandRowForm> getBrandRows() { return brandRows; }
    public void setBrandRows(List<BrandRowForm> v) { brandRows = v; }
    public List<KoseiRowForm> getKoseiRows() { return koseiRows; }
    public void setKoseiRows(List<KoseiRowForm> v) { koseiRows = v; }
    public List<MeisaiRowForm> getMeisaiRows() { return meisaiRows; }
    public void setMeisaiRows(List<MeisaiRowForm> v) { meisaiRows = v; }
}
