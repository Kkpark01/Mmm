package com.daifuku.mcm.form;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * 【変換元】Mcm2006uScreen.vb / Mcm2006uTabControl.vb
 * MCM2006U ユーザ契約内容フォーム
 *
 * 構成:
 *   ヘッダー  : MCM_UK_KEIYAKU（契約）+ MCM_UK_SEIBAN（製番）
 *   タブ一覧  : MCM_UK_KIKAN（期間タブ）
 *   タブ内容  : MCM_UK_BRAND + MCM_UK_KIKIKOSEI + MCM_UK_KIKIMEISAI + MCM_UK_KOTAIMEISAI + MCM_UK_TENKEN
 */
public class Mcm2006uForm implements Serializable {

    private static final long serialVersionUID = 1L;
    private String version;
    public String getVersion(){return version;}
    public void setVersion(String value){version=value;}

    // ─── 遷移元区分 ─────────────────────────────────────────────────────
    private int seniMotoKbn;

    // ─── ヘッダー: MCM_UK_KEIYAKU ───────────────────────────────────────
    private BigDecimal ukKeiyakuId;
    private BigDecimal nonyusakiId;
    private String nonyusakiCd;
    private String nonyusakiNk;
    private String nonyusakijusyo1Nk;
    private String nonyusakijusyo2Nk;
    private BigDecimal plantId;
    private String supportId;
    private String plantNk;
    private String nonyubusyoNk;
    private String nonyutantosyaNk;
    private String nonyutelNo;
    private String nonyufaxNo;
    private String keiyakuDt;
    private String shokaiKeiyakuDt;
    private String jikaikosinDt;
    private String keiyakumanryoDt;
    private String entyokeiyakumanryoDt;
    private String kaiyakuDt;
    private String jotai;
    private String autoFlg;

    // ─── 製番: MCM_UK_SEIBAN ────────────────────────────────────────────
    private List<SeibanRowForm> seibanRows = new ArrayList<>();

    // ─── 期間タブ一覧 ────────────────────────────────────────────────────
    private List<KikanTabForm> kikanTabs = new ArrayList<>();

    // 現在選択中のタブID（null = 先頭タブ）
    private BigDecimal selectedKikanId;

    // ─── getter / setter ─────────────────────────────────────────────────

    public int getSeniMotoKbn() { return seniMotoKbn; }
    public void setSeniMotoKbn(int seniMotoKbn) { this.seniMotoKbn = seniMotoKbn; }

    public BigDecimal getUkKeiyakuId() { return ukKeiyakuId; }
    public void setUkKeiyakuId(BigDecimal ukKeiyakuId) { this.ukKeiyakuId = ukKeiyakuId; }

    public BigDecimal getNonyusakiId() { return nonyusakiId; }
    public void setNonyusakiId(BigDecimal nonyusakiId) { this.nonyusakiId = nonyusakiId; }

    public String getNonyusakiCd() { return nonyusakiCd; }
    public void setNonyusakiCd(String nonyusakiCd) { this.nonyusakiCd = nonyusakiCd; }

    public String getNonyusakiNk() { return nonyusakiNk; }
    public void setNonyusakiNk(String nonyusakiNk) { this.nonyusakiNk = nonyusakiNk; }

    public String getNonyusakijusyo1Nk() { return nonyusakijusyo1Nk; }
    public void setNonyusakijusyo1Nk(String v) { this.nonyusakijusyo1Nk = v; }

    public String getNonyusakijusyo2Nk() { return nonyusakijusyo2Nk; }
    public void setNonyusakijusyo2Nk(String v) { this.nonyusakijusyo2Nk = v; }

    public BigDecimal getPlantId() { return plantId; }
    public void setPlantId(BigDecimal plantId) { this.plantId = plantId; }

    public String getSupportId() { return supportId; }
    public void setSupportId(String supportId) { this.supportId = supportId; }

    public String getPlantNk() { return plantNk; }
    public void setPlantNk(String plantNk) { this.plantNk = plantNk; }

    public String getNonyubusyoNk() { return nonyubusyoNk; }
    public void setNonyubusyoNk(String v) { this.nonyubusyoNk = v; }

    public String getNonyutantosyaNk() { return nonyutantosyaNk; }
    public void setNonyutantosyaNk(String v) { this.nonyutantosyaNk = v; }

    public String getNonyutelNo() { return nonyutelNo; }
    public void setNonyutelNo(String v) { this.nonyutelNo = v; }

    public String getNonyufaxNo() { return nonyufaxNo; }
    public void setNonyufaxNo(String v) { this.nonyufaxNo = v; }

    public String getKeiyakuDt() { return keiyakuDt; }
    public void setKeiyakuDt(String keiyakuDt) { this.keiyakuDt = keiyakuDt; }

    public String getShokaiKeiyakuDt() { return shokaiKeiyakuDt; }
    public void setShokaiKeiyakuDt(String v) { this.shokaiKeiyakuDt = v; }

    public String getJikaikosinDt() { return jikaikosinDt; }
    public void setJikaikosinDt(String jikaikosinDt) { this.jikaikosinDt = jikaikosinDt; }

    public String getKeiyakumanryoDt() { return keiyakumanryoDt; }
    public void setKeiyakumanryoDt(String v) { this.keiyakumanryoDt = v; }

    public String getEntyokeiyakumanryoDt() { return entyokeiyakumanryoDt; }
    public void setEntyokeiyakumanryoDt(String v) { this.entyokeiyakumanryoDt = v; }

    public String getKaiyakuDt() { return kaiyakuDt; }
    public void setKaiyakuDt(String kaiyakuDt) { this.kaiyakuDt = kaiyakuDt; }

    public String getJotai() { return jotai; }
    public void setJotai(String jotai) { this.jotai = jotai; }

    public String getAutoFlg() { return autoFlg; }
    public void setAutoFlg(String autoFlg) { this.autoFlg = autoFlg; }

    public List<SeibanRowForm> getSeibanRows() { return seibanRows; }
    public void setSeibanRows(List<SeibanRowForm> seibanRows) { this.seibanRows = seibanRows; }

    public List<KikanTabForm> getKikanTabs() { return kikanTabs; }
    public void setKikanTabs(List<KikanTabForm> kikanTabs) { this.kikanTabs = kikanTabs; }

    public BigDecimal getSelectedKikanId() { return selectedKikanId; }
    public void setSelectedKikanId(BigDecimal selectedKikanId) { this.selectedKikanId = selectedKikanId; }

    /** 選択中タブを返す。nullの場合は先頭タブを返す */
    public KikanTabForm getSelectedTab() {
        if (kikanTabs == null || kikanTabs.isEmpty()) return null;
        if (selectedKikanId == null) return kikanTabs.get(0);
        return kikanTabs.stream()
            .filter(t -> com.daifuku.mcm.common.CustomerScreenSupport.same(selectedKikanId, t.getUkKikanId()))
            .findFirst()
            .orElse(kikanTabs.get(0));
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 内部クラス: 製番行 (MCM_UK_SEIBAN)
    // ═══════════════════════════════════════════════════════════════════════

    public static class SeibanRowForm implements Serializable {
        private static final long serialVersionUID = 1L;

        private BigDecimal ukSeibanId;
        private BigDecimal ukKeiyakuId;
        private String kaisiDt;
        private String syuryoDt;
        private String kakuninKbn;

        public BigDecimal getUkSeibanId() { return ukSeibanId; }
        public void setUkSeibanId(BigDecimal v) { this.ukSeibanId = v; }
        public BigDecimal getUkKeiyakuId() { return ukKeiyakuId; }
        public void setUkKeiyakuId(BigDecimal v) { this.ukKeiyakuId = v; }
        public String getKaisiDt() { return kaisiDt; }
        public void setKaisiDt(String v) { this.kaisiDt = v; }
        public String getSyuryoDt() { return syuryoDt; }
        public void setSyuryoDt(String v) { this.syuryoDt = v; }
        public String getKakuninKbn() { return kakuninKbn; }
        public void setKakuninKbn(String v) { this.kakuninKbn = v; }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 内部クラス: 期間タブ (MCM_UK_KIKAN + 関連テーブル)
    // ═══════════════════════════════════════════════════════════════════════

    public static class KikanTabForm implements Serializable {
        private boolean replaceEquipment;
        public boolean isReplaceEquipment() { return replaceEquipment; }
        public void setReplaceEquipment(boolean value) { replaceEquipment = value; }
        private static final long serialVersionUID = 1L;

        private BigDecimal ukKikanId;        // null = 新規
        private BigDecimal ukKeiyakuId;
        private String kaisiDt;
        private String syuryoDt;
        private BigDecimal nonyusakiId;
        private String nonyusakiCd;
        private String nonyusakiNk;
        private String nonyusakijusyo1Nk;
        private String nonyusakijusyo2Nk;
        private BigDecimal plantId;
        private String supportId;
        private String plantNk;
        private String nonyubusyoNk;
        private String nonyutantosyaNk;
        private String nonyutelNo;
        private String nonyufaxNo;
        private String keiyakujikantai;
        private String hosyuhoho;
        private BigDecimal hosyuGkin;
        private BigDecimal iraitenpoId;
        private String iraimeisho1Nk;
        private String iraimeisho2Nk;
        private String iraimeisho3Nk;
        private String iraimeisho4Nk;
        private String iraitenporyakuNk;
        private String iraitantoNk;
        private String biko;
        private int tabFlg;                  // TAB_FLG_* 定数
        private String yukoFlg;

        // タブ内グリッド（全てMCM2007U経由で設定 or DBから読み込み）
        private List<BrandRowForm> brandRows     = new ArrayList<>();
        private List<KoseiRowForm> koseiRows     = new ArrayList<>();
        private List<MeisaiRowForm> meisaiRows   = new ArrayList<>();
        private List<KotaiRowForm> kotaiRows     = new ArrayList<>();
        private List<TenkenRowForm> tenkenRows   = new ArrayList<>();

        public BigDecimal getUkKikanId() { return ukKikanId; }
        public void setUkKikanId(BigDecimal v) { this.ukKikanId = v; }
        public BigDecimal getUkKeiyakuId() { return ukKeiyakuId; }
        public void setUkKeiyakuId(BigDecimal v) { this.ukKeiyakuId = v; }
        public String getKaisiDt() { return kaisiDt; }
        public void setKaisiDt(String v) { this.kaisiDt = v; }
        public String getSyuryoDt() { return syuryoDt; }
        public void setSyuryoDt(String v) { this.syuryoDt = v; }
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
        public BigDecimal getPlantId() { return plantId; }
        public void setPlantId(BigDecimal v) { this.plantId = v; }
        public String getSupportId() { return supportId; }
        public void setSupportId(String v) { this.supportId = v; }
        public String getPlantNk() { return plantNk; }
        public void setPlantNk(String v) { this.plantNk = v; }
        public String getNonyubusyoNk() { return nonyubusyoNk; }
        public void setNonyubusyoNk(String v) { this.nonyubusyoNk = v; }
        public String getNonyutantosyaNk() { return nonyutantosyaNk; }
        public void setNonyutantosyaNk(String v) { this.nonyutantosyaNk = v; }
        public String getNonyutelNo() { return nonyutelNo; }
        public void setNonyutelNo(String v) { this.nonyutelNo = v; }
        public String getNonyufaxNo() { return nonyufaxNo; }
        public void setNonyufaxNo(String v) { this.nonyufaxNo = v; }
        public String getKeiyakujikantai() { return keiyakujikantai; }
        public void setKeiyakujikantai(String v) { this.keiyakujikantai = v; }
        public String getHosyuhoho() { return hosyuhoho; }
        public void setHosyuhoho(String v) { this.hosyuhoho = v; }
        public BigDecimal getHosyuGkin() { return hosyuGkin; }
        public void setHosyuGkin(BigDecimal v) { this.hosyuGkin = v; }
        public BigDecimal getIraitenpoId() { return iraitenpoId; }
        public void setIraitenpoId(BigDecimal v) { this.iraitenpoId = v; }
        public String getIraimeisho1Nk() { return iraimeisho1Nk; }
        public void setIraimeisho1Nk(String v) { this.iraimeisho1Nk = v; }
        public String getIraimeisho2Nk() { return iraimeisho2Nk; }
        public void setIraimeisho2Nk(String v) { this.iraimeisho2Nk = v; }
        public String getIraimeisho3Nk() { return iraimeisho3Nk; }
        public void setIraimeisho3Nk(String v) { this.iraimeisho3Nk = v; }
        public String getIraimeisho4Nk() { return iraimeisho4Nk; }
        public void setIraimeisho4Nk(String v) { this.iraimeisho4Nk = v; }
        public String getIraitenporyakuNk() { return iraitenporyakuNk; }
        public void setIraitenporyakuNk(String v) { this.iraitenporyakuNk = v; }
        public String getIraitantoNk() { return iraitantoNk; }
        public void setIraitantoNk(String v) { this.iraitantoNk = v; }
        public String getBiko() { return biko; }
        public void setBiko(String biko) { this.biko = biko; }
        public int getTabFlg() { return tabFlg; }
        public void setTabFlg(int tabFlg) { this.tabFlg = tabFlg; }
        public String getYukoFlg() { return yukoFlg; }
        public void setYukoFlg(String yukoFlg) { this.yukoFlg = yukoFlg; }

        public List<BrandRowForm> getBrandRows() { return brandRows; }
        public void setBrandRows(List<BrandRowForm> v) { this.brandRows = v; }
        public List<KoseiRowForm> getKoseiRows() { return koseiRows; }
        public void setKoseiRows(List<KoseiRowForm> v) { this.koseiRows = v; }
        public List<MeisaiRowForm> getMeisaiRows() { return meisaiRows; }
        public void setMeisaiRows(List<MeisaiRowForm> v) { this.meisaiRows = v; }
        public List<KotaiRowForm> getKotaiRows() { return kotaiRows; }
        public void setKotaiRows(List<KotaiRowForm> v) { this.kotaiRows = v; }
        public List<TenkenRowForm> getTenkenRows() { return tenkenRows; }
        public void setTenkenRows(List<TenkenRowForm> v) { this.tenkenRows = v; }

        /** タブ表示ラベル（期間を "YYYY/MM/DD～YYYY/MM/DD" 形式で返す） */
        public String getTabLabel() {
            String s = kaisiDt != null ? kaisiDt : "";
            String e = syuryoDt != null ? syuryoDt : "";
            return s + "～" + e;
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 内部クラス: ブランド行 (MCM_UK_BRAND)
    // ═══════════════════════════════════════════════════════════════════════

    public static class BrandRowForm implements Serializable {
        private static final long serialVersionUID = 1L;

        private BigDecimal ukBrandId;       // null = 新規
        private BigDecimal ukKikanId;
        private BigDecimal brandkoseiId;    // MCM_MA_BRAND_KOSEI.BRANDKOSEI_ID
        private BigDecimal umMitsumoriId;
        private BigDecimal umKihonBrandId;
        private String brandNk;
        private String brandsyosaiNk;
        private String keiyakujikantai;
        private String hosyuhoho;

        public BigDecimal getUkBrandId() { return ukBrandId; }
        public void setUkBrandId(BigDecimal v) { this.ukBrandId = v; }
        public BigDecimal getUkKikanId() { return ukKikanId; }
        public void setUkKikanId(BigDecimal v) { this.ukKikanId = v; }
        public BigDecimal getBrandkoseiId() { return brandkoseiId; }
        public void setBrandkoseiId(BigDecimal v) { this.brandkoseiId = v; }
        public BigDecimal getUmMitsumoriId() { return umMitsumoriId; }
        public void setUmMitsumoriId(BigDecimal v) { this.umMitsumoriId = v; }
        public BigDecimal getUmKihonBrandId() { return umKihonBrandId; }
        public void setUmKihonBrandId(BigDecimal v) { this.umKihonBrandId = v; }
        public String getBrandNk() { return brandNk; }
        public void setBrandNk(String v) { this.brandNk = v; }
        public String getBrandsyosaiNk() { return brandsyosaiNk; }
        public void setBrandsyosaiNk(String v) { this.brandsyosaiNk = v; }
        public String getKeiyakujikantai() { return keiyakujikantai; }
        public void setKeiyakujikantai(String v) { this.keiyakujikantai = v; }
        public String getHosyuhoho() { return hosyuhoho; }
        public void setHosyuhoho(String v) { this.hosyuhoho = v; }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 内部クラス: 機器構成行 (MCM_UK_KIKIKOSEI)
    // ═══════════════════════════════════════════════════════════════════════

    public static class KoseiRowForm implements Serializable {
        private static final long serialVersionUID = 1L;

        private BigDecimal ukKikoseiId;     // null = 新規
        private BigDecimal ukBrandId;
        private BigDecimal kikikoseiId;     // MCM_MA_KIKIKOSEI.KIKIKOSEI_ID
        private BigDecimal brandkoseiId;
        private String kikikoseiNk;
        private String setNm;
        private String tani;
        private String tehaiseiban;
        private String controllerFlg;
        private String hosyuhoho;
        private int hyojijun;

        public BigDecimal getUkKikoseiId() { return ukKikoseiId; }
        public void setUkKikoseiId(BigDecimal v) { this.ukKikoseiId = v; }
        public BigDecimal getUkBrandId() { return ukBrandId; }
        public void setUkBrandId(BigDecimal v) { this.ukBrandId = v; }
        public BigDecimal getKikikoseiId() { return kikikoseiId; }
        public void setKikikoseiId(BigDecimal v) { this.kikikoseiId = v; }
        public BigDecimal getBrandkoseiId() { return brandkoseiId; }
        public void setBrandkoseiId(BigDecimal v) { this.brandkoseiId = v; }
        public String getKikikoseiNk() { return kikikoseiNk; }
        public void setKikikoseiNk(String v) { this.kikikoseiNk = v; }
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
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 内部クラス: 機器明細行 (MCM_UK_KIKIMEISAI)
    // ═══════════════════════════════════════════════════════════════════════

    public static class MeisaiRowForm implements Serializable {
        private static final long serialVersionUID = 1L;

        private BigDecimal ukKikimeisaiId;  // null = 新規
        private BigDecimal ukKikoseiId;
        private BigDecimal kikimeisaiId;    // MCM_MA_KIKIMEISAI.KIKIMEISAI_ID
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

        public BigDecimal getUkKikimeisaiId() { return ukKikimeisaiId; }
        public void setUkKikimeisaiId(BigDecimal v) { this.ukKikimeisaiId = v; }
        public BigDecimal getUkKikoseiId() { return ukKikoseiId; }
        public void setUkKikoseiId(BigDecimal v) { this.ukKikoseiId = v; }
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
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 内部クラス: 個体明細行 (MCM_UK_KOTAIMEISAI)
    // ═══════════════════════════════════════════════════════════════════════

    public static class KotaiRowForm implements Serializable {
        private static final long serialVersionUID = 1L;

        private BigDecimal ukKotaimeisaiId; // null = 新規
        private BigDecimal ukKikimeisaiId;
        private BigDecimal ukKikoseiId;
        private BigDecimal kotaikanriId;    // MCM_MA_KIKIKOTAIKANRI.KOTAIKANRI_ID
        private BigDecimal kikimeisaiId;
        private BigDecimal kikikoseiId;
        private BigDecimal brandkoseiId;
        private String kotaiNk;
        private String serialNo;
        private String itizinonnyuDt;
        private String setchibasyo;
        private String keiyakukigenDt;
        private String enchokeiyakukigenDt;

        public BigDecimal getUkKotaimeisaiId() { return ukKotaimeisaiId; }
        public void setUkKotaimeisaiId(BigDecimal v) { this.ukKotaimeisaiId = v; }
        public BigDecimal getUkKikimeisaiId() { return ukKikimeisaiId; }
        public void setUkKikimeisaiId(BigDecimal v) { this.ukKikimeisaiId = v; }
        public BigDecimal getUkKikoseiId() { return ukKikoseiId; }
        public void setUkKikoseiId(BigDecimal v) { this.ukKikoseiId = v; }
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
        public String getKeiyakukigenDt() { return keiyakukigenDt; }
        public void setKeiyakukigenDt(String v) { this.keiyakukigenDt = v; }
        public String getEnchokeiyakukigenDt() { return enchokeiyakukigenDt; }
        public void setEnchokeiyakukigenDt(String v) { this.enchokeiyakukigenDt = v; }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 内部クラス: 点検行 (MCM_UK_TENKEN)
    // ═══════════════════════════════════════════════════════════════════════

    public static class TenkenRowForm implements Serializable {
        private static final long serialVersionUID = 1L;

        private BigDecimal ukTenkenId;      // null = 新規
        private BigDecimal ukKikanId;
        private String tenkenDt;
        private String naiyo;
        private String biko;
        private List<String> months = new ArrayList<>();
        public String getNaiyo() { return naiyo; }
        public void setNaiyo(String v) { naiyo=v; }
        public String getBiko() { return biko; }
        public void setBiko(String v) { biko=v; }
        public List<String> getMonths() { return months; }
        private String tenkenJotai;

        public BigDecimal getUkTenkenId() { return ukTenkenId; }
        public void setUkTenkenId(BigDecimal v) { this.ukTenkenId = v; }
        public BigDecimal getUkKikanId() { return ukKikanId; }
        public void setUkKikanId(BigDecimal v) { this.ukKikanId = v; }
        public String getTenkenDt() { return tenkenDt; }
        public void setTenkenDt(String v) { this.tenkenDt = v; }
        public String getTenkenJotai() { return tenkenJotai; }
        public void setTenkenJotai(String v) { this.tenkenJotai = v; }
    }
}
