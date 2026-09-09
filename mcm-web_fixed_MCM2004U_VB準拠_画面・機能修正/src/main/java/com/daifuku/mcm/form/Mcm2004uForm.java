package com.daifuku.mcm.form;

import com.daifuku.mcm.dto.Mcm2004uRowDto;

import java.util.ArrayList;
import java.util.List;

/**
 * 【変換元】Mcm2004uScreen.Designer.vb - 検索条件コントロール群
 * MCM2004U（カスタマー見積・契約一覧）画面 Form
 *
 * <p>検索条件（テキスト9項目 + チェック9項目 + コンボ1項目）と
 * 検索結果グリッドを保持する。</p>
 *
 * @author MCM Migration Tool
 */
public class Mcm2004uForm {

    // ========================================================
    // 検索条件 — テキスト
    // 【変換元】Mcm2004uScreen.Designer.vb - TextBox群
    // ========================================================

    /** 見積NO（VB: MITSUMORI_NOTextBox） */
    private String umMitsumoriNo;

    /** 契約NO（VB: KEIYAKU_NOTextBox） */
    private String keiyakuNo;

    /** 見積作成日From（VB: MITSUMORI_DT_STTextBox） */
    private String mitsumoriDtStart;

    /** 見積作成日To（VB: MITSUMORI_DT_EDTextBox） */
    private String mitsumoriDtEnd;

    /** 納入先コード（VB: NONYUSAKI_CDTextBox） */
    private String nonyusakiCd;

    /** 納入先名（VB: NONYUSAKI_NKTextBox） */
    private String nonyusakiNk;

    /** サポートID（VB: SUPPORT_IDTextBox） */
    private String supportId;

    /** プラント名（VB: PLANT_NKTextBox） */
    private String plantNk;

    // ========================================================
    // 検索条件 — コンボボックス
    // ========================================================

    /** 送付先事業所（VB: SOFUTENPO_IDComboBox） */
    private String sofutenpoId;

    // ========================================================
    // 検索条件 — 状態チェックボックス
    // 【変換元】Mcm2004uScreen.Designer.vb - CheckBox群
    // ========================================================

    /** 見積（VB: JOTAI_MITSUMORICheckBox） */
    private boolean jotaiMitsumori = true;

    /** 契約（VB: JOTAI_KEIYAKUCheckBox） */
    private boolean jotaiKeiyaku = true;

    /** 解約（VB: JOTAI_KAIYAKUCheckBox） */
    private boolean jotaiKaiyaku;

    /** 破棄（VB: JOTAI_HAKICheckBox） */
    private boolean jotaiHaki;

    // ========================================================
    // 検索条件 — 承認状態チェックボックス
    // ========================================================

    /** 作成中（VB: SYOUNINJOTAI_SAKUSEICHUCheckBox） */
    private boolean syouninJotaiSakuseichu;

    /** 審査中（VB: SYOUNINJOTAI_SINSACHUCheckBox） */
    private boolean syouninJotaiSinsachu;

    /** 承認中（VB: SYOUNINJOTAI_SHONINCHUCheckBox） */
    private boolean syouninJotaiShoninchu;

    /** 承認済み（VB: SYOUNINJOTAI_SHONINZUMICheckBox） */
    private boolean syouninJotaiShoninzumi;

    /** 差し戻し中（VB: SYOUNINJOTAI_SASHIMODOSHICHUCheckBox） */
    private boolean syouninJotaiSashimodoshichu;

    // ========================================================
    // 検索結果
    // ========================================================

    /** 検索結果行リスト */
    private List<Mcm2004uRowDto> rows = new ArrayList<>();

    // ========================================================
    // Getters / Setters
    // ========================================================

    public String getUmMitsumoriNo() {
        return umMitsumoriNo;
    }

    public void setUmMitsumoriNo(String umMitsumoriNo) {
        this.umMitsumoriNo = umMitsumoriNo;
    }

    public String getKeiyakuNo() {
        return keiyakuNo;
    }

    public void setKeiyakuNo(String keiyakuNo) {
        this.keiyakuNo = keiyakuNo;
    }

    public String getMitsumoriDtStart() {
        return mitsumoriDtStart;
    }

    public void setMitsumoriDtStart(String mitsumoriDtStart) {
        this.mitsumoriDtStart = mitsumoriDtStart;
    }

    public String getMitsumoriDtEnd() {
        return mitsumoriDtEnd;
    }

    public void setMitsumoriDtEnd(String mitsumoriDtEnd) {
        this.mitsumoriDtEnd = mitsumoriDtEnd;
    }

    public String getNonyusakiCd() {
        return nonyusakiCd;
    }

    public void setNonyusakiCd(String nonyusakiCd) {
        this.nonyusakiCd = nonyusakiCd;
    }

    public String getNonyusakiNk() {
        return nonyusakiNk;
    }

    public void setNonyusakiNk(String nonyusakiNk) {
        this.nonyusakiNk = nonyusakiNk;
    }

    public String getSupportId() {
        return supportId;
    }

    public void setSupportId(String supportId) {
        this.supportId = supportId;
    }

    public String getPlantNk() {
        return plantNk;
    }

    public void setPlantNk(String plantNk) {
        this.plantNk = plantNk;
    }

    public String getSofutenpoId() {
        return sofutenpoId;
    }

    public void setSofutenpoId(String sofutenpoId) {
        this.sofutenpoId = sofutenpoId;
    }

    public boolean isJotaiMitsumori() {
        return jotaiMitsumori;
    }

    public void setJotaiMitsumori(boolean jotaiMitsumori) {
        this.jotaiMitsumori = jotaiMitsumori;
    }

    public boolean isJotaiKeiyaku() {
        return jotaiKeiyaku;
    }

    public void setJotaiKeiyaku(boolean jotaiKeiyaku) {
        this.jotaiKeiyaku = jotaiKeiyaku;
    }

    public boolean isJotaiKaiyaku() {
        return jotaiKaiyaku;
    }

    public void setJotaiKaiyaku(boolean jotaiKaiyaku) {
        this.jotaiKaiyaku = jotaiKaiyaku;
    }

    public boolean isJotaiHaki() {
        return jotaiHaki;
    }

    public void setJotaiHaki(boolean jotaiHaki) {
        this.jotaiHaki = jotaiHaki;
    }

    public boolean isSyouninJotaiSakuseichu() {
        return syouninJotaiSakuseichu;
    }

    public void setSyouninJotaiSakuseichu(boolean syouninJotaiSakuseichu) {
        this.syouninJotaiSakuseichu = syouninJotaiSakuseichu;
    }

    public boolean isSyouninJotaiSinsachu() {
        return syouninJotaiSinsachu;
    }

    public void setSyouninJotaiSinsachu(boolean syouninJotaiSinsachu) {
        this.syouninJotaiSinsachu = syouninJotaiSinsachu;
    }

    public boolean isSyouninJotaiShoninchu() {
        return syouninJotaiShoninchu;
    }

    public void setSyouninJotaiShoninchu(boolean syouninJotaiShoninchu) {
        this.syouninJotaiShoninchu = syouninJotaiShoninchu;
    }

    public boolean isSyouninJotaiShoninzumi() {
        return syouninJotaiShoninzumi;
    }

    public void setSyouninJotaiShoninzumi(boolean syouninJotaiShoninzumi) {
        this.syouninJotaiShoninzumi = syouninJotaiShoninzumi;
    }

    public boolean isSyouninJotaiSashimodoshichu() {
        return syouninJotaiSashimodoshichu;
    }

    public void setSyouninJotaiSashimodoshichu(boolean syouninJotaiSashimodoshichu) {
        this.syouninJotaiSashimodoshichu = syouninJotaiSashimodoshichu;
    }

    public List<Mcm2004uRowDto> getRows() {
        return rows;
    }

    public void setRows(List<Mcm2004uRowDto> rows) {
        this.rows = rows;
    }
}
