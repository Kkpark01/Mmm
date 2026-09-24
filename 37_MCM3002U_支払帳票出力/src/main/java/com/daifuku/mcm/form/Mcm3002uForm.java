package com.daifuku.mcm.form;

import java.io.Serializable;
import java.math.BigDecimal;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

/**
 * 【変換元】Mcm3002uScreen.vb / Mcm3002uDataSet.xsd（MCM_TK_SIHARAI_V）
 * MCM3002U 支払い明細出力指示フォーム
 *
 * 帳票の金額・期間はサーバーで取得し、リクエストの明細値は出力に使わない。
 */
public class Mcm3002uForm implements Serializable {

    private static final long serialVersionUID = 1L;

    /** 支払い月（YYYYMM形式）。初期値=システム日付の当月 */
    private String shiharaiTsuki = YearMonth.now().format(DateTimeFormatter.ofPattern("yyyyMM"));

    /**
     * 支払区分。
     * 【変換元】NenbaraiRadioButton（デフォルト）/ TsukibaraiRadioButton
     * VB版定数: Mcm3002uConstant.NENBARAI="年払" / TSUKIBARAI="月払"
     */
    private String siharaiKbn = "年払";

    /** 検索結果 */
    private List<RowForm> rows = new ArrayList<>();

    // ===================================================================
    // 内部クラス: 検索結果行
    // ===================================================================

    public static class RowForm implements Serializable {

        private java.time.LocalDate kaishiDt;
        private java.time.LocalDate syuryoDt;
        private BigDecimal shiharaiKingaku;
        public java.time.LocalDate getKaishiDt() { return kaishiDt; }
        public void setKaishiDt(java.time.LocalDate value) { kaishiDt = value; }
        public java.time.LocalDate getSyuryoDt() { return syuryoDt; }
        public void setSyuryoDt(java.time.LocalDate value) { syuryoDt = value; }
        public BigDecimal getShiharaiKingaku() { return shiharaiKingaku; }
        public void setShiharaiKingaku(BigDecimal value) { shiharaiKingaku = value; }

        private static final long serialVersionUID = 1L;

        /** MCM_TK_KIKAN.TK_KIKAN_ID */
        private BigDecimal tkKikanId;

        /** 取引先ID */
        private BigDecimal torihikisakiId;

        /** 納入先コード */
        private String nonyusakiCd;

        /** サポートID */
        private String supportId;

        /** 業務状態 */
        private String jotai;

        /** 納入先名 */
        private String nonyusakiNk;

        /** 店舗名 */
        private String tenpoNk;

        /** 書類NO（KEIYAKU_NO） */
        private String keiyakuNo;

        /** 取引先コード */
        private String torihikisakiCd;

        /** 取引先名 */
        private String torihikisakiNk;

        /**
         * 取引窓口担当者名
         * 【変換元】MCM_MA_TORIMADOGUCHI サブクエリ（MADOGUCHI_KBN='2', YUKO_FLG=0）
         */
        private String torisyutantosyaNk;

        /** 支払区分（"年払" / "月払"） */
        private String siharai;

        /** 回数 */
        private BigDecimal kaisu;

        /** 備考 */
        private String biko;

        /** 支払月（DATE型、表示用は YYYYMM 文字列） */
        private String tsuki;

        /** ハードウェア製番 */
        private String hardSeiban;

        public BigDecimal getTkKikanId() { return tkKikanId; }
        public void setTkKikanId(BigDecimal tkKikanId) { this.tkKikanId = tkKikanId; }

        public BigDecimal getTorihikisakiId() { return torihikisakiId; }
        public void setTorihikisakiId(BigDecimal torihikisakiId) { this.torihikisakiId = torihikisakiId; }

        public String getNonyusakiCd() { return nonyusakiCd; }
        public void setNonyusakiCd(String nonyusakiCd) { this.nonyusakiCd = nonyusakiCd; }

        public String getSupportId() { return supportId; }
        public void setSupportId(String supportId) { this.supportId = supportId; }

        public String getJotai() { return jotai; }
        public void setJotai(String jotai) { this.jotai = jotai; }

        public String getNonyusakiNk() { return nonyusakiNk; }
        public void setNonyusakiNk(String nonyusakiNk) { this.nonyusakiNk = nonyusakiNk; }

        public String getTenpoNk() { return tenpoNk; }
        public void setTenpoNk(String tenpoNk) { this.tenpoNk = tenpoNk; }

        public String getKeiyakuNo() { return keiyakuNo; }
        public void setKeiyakuNo(String keiyakuNo) { this.keiyakuNo = keiyakuNo; }

        public String getTorihikisakiCd() { return torihikisakiCd; }
        public void setTorihikisakiCd(String torihikisakiCd) { this.torihikisakiCd = torihikisakiCd; }

        public String getTorihikisakiNk() { return torihikisakiNk; }
        public void setTorihikisakiNk(String torihikisakiNk) { this.torihikisakiNk = torihikisakiNk; }

        public String getTorisyutantosyaNk() { return torisyutantosyaNk; }
        public void setTorisyutantosyaNk(String torisyutantosyaNk) { this.torisyutantosyaNk = torisyutantosyaNk; }

        public String getSiharai() { return siharai; }
        public void setSiharai(String siharai) { this.siharai = siharai; }

        public BigDecimal getKaisu() { return kaisu; }
        public void setKaisu(BigDecimal kaisu) { this.kaisu = kaisu; }

        public String getBiko() { return biko; }
        public void setBiko(String biko) { this.biko = biko; }

        public String getTsuki() { return tsuki; }
        public void setTsuki(String tsuki) { this.tsuki = tsuki; }

        public String getHardSeiban() { return hardSeiban; }
        public void setHardSeiban(String hardSeiban) { this.hardSeiban = hardSeiban; }
    }

    // ===================================================================
    // Getter / Setter
    // ===================================================================

    public String getShiharaiTsuki() { return shiharaiTsuki; }
    public void setShiharaiTsuki(String shiharaiTsuki) { this.shiharaiTsuki = shiharaiTsuki; }

    public String getSiharaiKbn() { return siharaiKbn; }
    public void setSiharaiKbn(String siharaiKbn) { this.siharaiKbn = siharaiKbn; }

    public List<RowForm> getRows() { return rows; }
    public void setRows(List<RowForm> rows) { this.rows = rows; }
}
