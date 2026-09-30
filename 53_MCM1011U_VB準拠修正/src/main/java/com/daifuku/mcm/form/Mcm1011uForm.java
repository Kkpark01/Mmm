package com.daifuku.mcm.form;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * 【変換元】Mcm1011uScreen.vb
 * MCM1011U 審査・承認（取引先契約）フォーム
 */
public class Mcm1011uForm implements Serializable {

    private static final long serialVersionUID = 1L;

    // ===================================================================
    // 検索条件
    // ===================================================================

    /** 納入先コード（前方一致） */
    private String nonyusakiCd = "";

    /** 納入先名（部分一致） */
    private String nonyusakiNk = "";

    /** サポートID（部分一致） */
    private String supportId = "";

    /** プラント名（部分一致） */
    private String plantNk = "";

    /** 審査中チェック */
    private boolean shinsachu = false;

    /** 承認中チェック */
    private boolean shoninchu = false;

    /** 承認済チェック */
    private boolean shoninzumi = false;

    /** 差戻しチェック */
    private boolean sashimodoshi = false;

    // ===================================================================
    // 検索結果
    // ===================================================================

    private List<KeiyakuRowForm> keiyakuRows = new ArrayList<>();

    // ===================================================================
    // 内部クラス: 申請一覧行
    // ===================================================================

    public static class KeiyakuRowForm implements Serializable {

        private static final long serialVersionUID = 1L;

        /** チェックフラグ（承認/差戻し対象） */
        private java.time.LocalDateTime lastupdateDt;
        public java.time.LocalDateTime getLastupdateDt() { return lastupdateDt; }
        public void setLastupdateDt(java.time.LocalDateTime value) { lastupdateDt = value; }
        private boolean checkFlg;
        private String historyKey;
        private String contractState;
        private java.time.LocalDateTime periodVersion;
        private List<TenpuRowForm> pendingAttachments = new ArrayList<>();
        public String getHistoryKey() { return historyKey; }
        public void setHistoryKey(String value) { historyKey = value; }
        public String getRowKey() { return tkKikanId.stripTrailingZeros().toPlainString() + ":" + historyKey; }
        public String getContractState() { return contractState; }
        public void setContractState(String value) { contractState = value; }
        public java.time.LocalDateTime getPeriodVersion() { return periodVersion; }
        public void setPeriodVersion(java.time.LocalDateTime value) { periodVersion = value; }
        public List<TenpuRowForm> getPendingAttachments() { return pendingAttachments; }
        public void setPendingAttachments(List<TenpuRowForm> value) { pendingAttachments = value; }
        public String getJotaiName() {
            if (jotai == null) return "";
            return switch (jotai) { case "0" -> "依頼"; case "1" -> "見積"; case "2" -> "契約";
                case "3" -> "破棄"; case "4" -> "解約"; case "9" -> "作成中"; default -> jotai; };
        }
        public String getHoursDisplay() {
            if (keiyakujikantai == null || keiyakujikantai.isBlank()) return "";
            try { return String.format(java.util.Locale.JAPAN, "%,.2f", new BigDecimal(keiyakujikantai)); }
            catch (NumberFormatException ex) { return keiyakujikantai; }
        }

        /** 取引先契約ID */
        private BigDecimal tkKeiyakuId;

        /** 取引先契約期間ID（一覧の選択期間） */
        private BigDecimal tkKikanId;

        /** 契約NO */
        private String keiyakuNo;

        /** 承認状態コード */
        private String shoninjotai;

        /** 状態コード */
        private String jotai;

        /** 取引先ID */
        private BigDecimal torihikisakiId;

        /** 取引先名 */
        private String torihikisakiNk;

        /** 納入先コード */
        private String nonyusakiCd;

        /** 納入先名 */
        private String nonyusakiNk;
        private String oldNonyusakiNk;
        private String nonyusakiKojoNk;

        /** サポートID */
        private String supportId;

        /** プラント名 */
        private String plantNk;

        /** 開始日（yyyy/MM/dd） */
        private String kaisiDt;

        /** 契約時間帯 */
        private String keiyakujikantai;

        /** 仕切合計金額 */
        private BigDecimal kingaku;

        /** 依頼担当者 */
        private String iraitantosya;

        /** VB版一覧に表示する審査・承認の記録 */
        private String shinsaDt;
        private String shinsaBy;
        private String shoninDt;
        private String shoninBy;

        /** 添付ファイル一覧 */
        private List<TenpuRowForm> tenpuList = new ArrayList<>();


        public boolean isCheckFlg() { return checkFlg; }
        public void setCheckFlg(boolean checkFlg) { this.checkFlg = checkFlg; }

        public BigDecimal getTkKeiyakuId() { return tkKeiyakuId; }
        public void setTkKeiyakuId(BigDecimal tkKeiyakuId) { this.tkKeiyakuId = tkKeiyakuId; }

        public BigDecimal getTkKikanId() { return tkKikanId; }
        public void setTkKikanId(BigDecimal tkKikanId) { this.tkKikanId = tkKikanId; }

        public String getKeiyakuNo() { return keiyakuNo; }
        public void setKeiyakuNo(String keiyakuNo) { this.keiyakuNo = keiyakuNo; }

        public String getShoninjotai() { return shoninjotai; }
        public void setShoninjotai(String shoninjotai) { this.shoninjotai = shoninjotai; }

        public String getJotai() { return jotai; }
        public void setJotai(String jotai) { this.jotai = jotai; }

        public BigDecimal getTorihikisakiId() { return torihikisakiId; }
        public void setTorihikisakiId(BigDecimal torihikisakiId) { this.torihikisakiId = torihikisakiId; }

        public String getTorihikisakiNk() { return torihikisakiNk; }
        public void setTorihikisakiNk(String torihikisakiNk) { this.torihikisakiNk = torihikisakiNk; }

        public String getNonyusakiCd() { return nonyusakiCd; }
        public void setNonyusakiCd(String nonyusakiCd) { this.nonyusakiCd = nonyusakiCd; }

        public String getNonyusakiNk() { return nonyusakiNk; }
        public void setNonyusakiNk(String nonyusakiNk) { this.nonyusakiNk = nonyusakiNk; }
        public String getOldNonyusakiNk() { return oldNonyusakiNk; }
        public void setOldNonyusakiNk(String value) { oldNonyusakiNk = value; }
        public String getNonyusakiKojoNk() { return nonyusakiKojoNk; }
        public void setNonyusakiKojoNk(String value) { nonyusakiKojoNk = value; }

        public String getSupportId() { return supportId; }
        public void setSupportId(String supportId) { this.supportId = supportId; }

        public String getPlantNk() { return plantNk; }
        public void setPlantNk(String plantNk) { this.plantNk = plantNk; }

        public String getKaisiDt() { return kaisiDt; }
        public void setKaisiDt(String kaisiDt) { this.kaisiDt = kaisiDt; }

        public String getKeiyakujikantai() { return keiyakujikantai; }
        public void setKeiyakujikantai(String keiyakujikantai) { this.keiyakujikantai = keiyakujikantai; }

        public BigDecimal getKingaku() { return kingaku; }
        public void setKingaku(BigDecimal kingaku) { this.kingaku = kingaku; }

        public String getIraitantosya() { return iraitantosya; }
        public void setIraitantosya(String iraitantosya) { this.iraitantosya = iraitantosya; }

        public String getShinsaDt() { return shinsaDt; }
        public void setShinsaDt(String value) { shinsaDt = value; }
        public String getShinsaBy() { return shinsaBy; }
        public void setShinsaBy(String value) { shinsaBy = value; }
        public String getShoninDt() { return shoninDt; }
        public void setShoninDt(String value) { shoninDt = value; }
        public String getShoninBy() { return shoninBy; }
        public void setShoninBy(String value) { shoninBy = value; }

        public List<TenpuRowForm> getTenpuList() { return tenpuList; }
        public void setTenpuList(List<TenpuRowForm> tenpuList) { this.tenpuList = tenpuList; }
    }

    // ===================================================================
    // 内部クラス: 添付ファイル行
    // ===================================================================

    public static class TenpuRowForm implements Serializable {
        private String shoninjotai;
        private String historyKey;
        private java.time.LocalDateTime lastupdateDt;
        public String getShoninjotai() { return shoninjotai; }
        public void setShoninjotai(String value) { shoninjotai = value; }
        public String getHistoryKey() { return historyKey; }
        public void setHistoryKey(String value) { historyKey = value; }
        public java.time.LocalDateTime getLastupdateDt() { return lastupdateDt; }
        public void setLastupdateDt(java.time.LocalDateTime value) { lastupdateDt = value; }
        public boolean sameVersion(TenpuRowForm other) {
            return other != null && tkTenpuId.compareTo(other.tkTenpuId) == 0
                && java.util.Objects.equals(shoninjotai, other.shoninjotai)
                && java.util.Objects.equals(lastupdateDt, other.lastupdateDt)
                && java.util.Objects.equals(directory, other.directory)
                && java.util.Objects.equals(tenpufileNk, other.tenpufileNk);
        }

        private static final long serialVersionUID = 1L;

        /** 添付ID */
        private BigDecimal tkTenpuId;

        /** ファイル名 */
        private String tenpufileNk;

        /** ディレクトリ */
        private String directory;

        public BigDecimal getTkTenpuId() { return tkTenpuId; }
        public void setTkTenpuId(BigDecimal tkTenpuId) { this.tkTenpuId = tkTenpuId; }

        public String getTenpufileNk() { return tenpufileNk; }
        public void setTenpufileNk(String tenpufileNk) { this.tenpufileNk = tenpufileNk; }

        public String getDirectory() { return directory; }
        public void setDirectory(String directory) { this.directory = directory; }
    }

    // ===================================================================
    // Getter / Setter（検索条件）
    // ===================================================================

    public String getNonyusakiCd() { return nonyusakiCd; }
    public void setNonyusakiCd(String nonyusakiCd) { this.nonyusakiCd = nonyusakiCd; }

    public String getNonyusakiNk() { return nonyusakiNk; }
    public void setNonyusakiNk(String nonyusakiNk) { this.nonyusakiNk = nonyusakiNk; }

    public String getSupportId() { return supportId; }
    public void setSupportId(String supportId) { this.supportId = supportId; }

    public String getPlantNk() { return plantNk; }
    public void setPlantNk(String plantNk) { this.plantNk = plantNk; }

    public boolean isShinsachu() { return shinsachu; }
    public void setShinsachu(boolean shinsachu) { this.shinsachu = shinsachu; }

    public boolean isShoninchu() { return shoninchu; }
    public void setShoninchu(boolean shoninchu) { this.shoninchu = shoninchu; }

    public boolean isShoninzumi() { return shoninzumi; }
    public void setShoninzumi(boolean shoninzumi) { this.shoninzumi = shoninzumi; }

    public boolean isSashimodoshi() { return sashimodoshi; }
    public void setSashimodoshi(boolean sashimodoshi) { this.sashimodoshi = sashimodoshi; }

    public List<KeiyakuRowForm> getKeiyakuRows() { return keiyakuRows; }
    public void setKeiyakuRows(List<KeiyakuRowForm> keiyakuRows) { this.keiyakuRows = keiyakuRows; }
}