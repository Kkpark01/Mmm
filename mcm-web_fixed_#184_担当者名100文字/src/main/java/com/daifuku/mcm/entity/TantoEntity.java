package com.daifuku.mcm.entity;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import java.math.BigDecimal;
import java.time.LocalDateTime;

/**
 * 【変換元】Mcm3006uDataSet.xsd / Mcm0026uDataSet.Designer.vb
 *   テーブル: MCM_MO_TANTO（担当者マスタ）
 *
 * MCM0026U（パスワード初期化）およびMCM3006U（担当者マスタ）で使用。
 * PK型は BigDecimal（プロジェクト規約に準拠）。
 */
@Entity
@Table(name = "MCM_MO_TANTO")
public class TantoEntity {

    /** 担当者ID（PK） */
    @Id
    @Column(name = "TANTO_ID", nullable = false)
    private BigDecimal tantoId;

    /** ログインID（最大20文字） */
    @Column(name = "LOGIN_ID", length = 20)
    private String loginId;

    /** パスワード（最大20文字） */
    @Column(name = "PASSWORD", length = 20)
    private String password;

    /** 担当者名（最大40文字） */
    @Column(name = "TANTO_NK", length = 100)
    private String tantoNk;

    /** 事業所名 */
    @Column(name = "JIGYOSYO_NK")
    private String jigyosyoNk;

    /** メールアドレス */
    @Column(name = "MAILADDRESS")
    private String mailaddress;

    /** 上位者ログインID */
    @Column(name = "OYALOGIN_ID")
    private String oyaloginId;

    /** 審査フラグ */
    @Column(name = "SHINSA_FLG")
    private Integer shinsaFlg;

    /** 見積審査上限金額 */
    @Column(name = "MITSUMORI_KIN")
    private BigDecimal mitsumoriKin;

    /** 契約審査上限金額 */
    @Column(name = "KEIYAKU_KIN")
    private BigDecimal keiyakuKin;

    /** 承認フラグ */
    @Column(name = "SHONIN_FLG")
    private Integer shoninFlg;

    /** 見積承認上限金額 */
    @Column(name = "MITSUMORISHONIN_KIN")
    private BigDecimal mitsumorishoninKin;

    /** 契約承認上限金額 */
    @Column(name = "KEIYAKUSHONIN_KIN")
    private BigDecimal keiyakushoninKin;

    /** システム利用フラグ */
    @Column(name = "SYSTEMRIYO_FLG")
    private String systemriyoFlg;

    /** メンテナンスフラグ */
    @Column(name = "MAINTENANCE_FLG")
    private String maintenanceFlg;

    /** 備考 */
    @Column(name = "BIKO")
    private String biko;

    /** 登録日時 */
    @Column(name = "CREATED_DT")
    private LocalDateTime createdDt;

    /** 登録者 */
    @Column(name = "CREATED_BY", length = 50)
    private String createdBy;

    /** 更新日時 */
    @Column(name = "LASTUPDATE_DT")
    private LocalDateTime lastupdateDt;

    /** 更新者 */
    @Column(name = "LASTUPDATE_BY", length = 50)
    private String lastupdateBy;

    // ========== Getters / Setters ==========

    public BigDecimal getTantoId() { return tantoId; }
    public void setTantoId(BigDecimal tantoId) { this.tantoId = tantoId; }

    public String getLoginId() { return loginId; }
    public void setLoginId(String loginId) { this.loginId = loginId; }

    public String getPassword() { return password; }
    public void setPassword(String password) { this.password = password; }

    public String getTantoNk() { return tantoNk; }
    public void setTantoNk(String tantoNk) { this.tantoNk = tantoNk; }

    public String getJigyosyoNk() { return jigyosyoNk; }
    public void setJigyosyoNk(String jigyosyoNk) { this.jigyosyoNk = jigyosyoNk; }

    public String getMailaddress() { return mailaddress; }
    public void setMailaddress(String mailaddress) { this.mailaddress = mailaddress; }

    public String getOyaloginId() { return oyaloginId; }
    public void setOyaloginId(String oyaloginId) { this.oyaloginId = oyaloginId; }

    public Integer getShinsaFlg() { return shinsaFlg; }
    public void setShinsaFlg(Integer shinsaFlg) { this.shinsaFlg = shinsaFlg; }

    public BigDecimal getMitsumoriKin() { return mitsumoriKin; }
    public void setMitsumoriKin(BigDecimal mitsumoriKin) { this.mitsumoriKin = mitsumoriKin; }

    public BigDecimal getKeiyakuKin() { return keiyakuKin; }
    public void setKeiyakuKin(BigDecimal keiyakuKin) { this.keiyakuKin = keiyakuKin; }

    public Integer getShoninFlg() { return shoninFlg; }
    public void setShoninFlg(Integer shoninFlg) { this.shoninFlg = shoninFlg; }

    public BigDecimal getMitsumorishoninKin() { return mitsumorishoninKin; }
    public void setMitsumorishoninKin(BigDecimal mitsumorishoninKin) { this.mitsumorishoninKin = mitsumorishoninKin; }

    public BigDecimal getKeiyakushoninKin() { return keiyakushoninKin; }
    public void setKeiyakushoninKin(BigDecimal keiyakushoninKin) { this.keiyakushoninKin = keiyakushoninKin; }

    public String getSystemriyoFlg() { return systemriyoFlg; }
    public void setSystemriyoFlg(String systemriyoFlg) { this.systemriyoFlg = systemriyoFlg; }

    public String getMaintenanceFlg() { return maintenanceFlg; }
    public void setMaintenanceFlg(String maintenanceFlg) { this.maintenanceFlg = maintenanceFlg; }

    public String getBiko() { return biko; }
    public void setBiko(String biko) { this.biko = biko; }

    public LocalDateTime getCreatedDt() { return createdDt; }
    public void setCreatedDt(LocalDateTime createdDt) { this.createdDt = createdDt; }

    public String getCreatedBy() { return createdBy; }
    public void setCreatedBy(String createdBy) { this.createdBy = createdBy; }

    public LocalDateTime getLastupdateDt() { return lastupdateDt; }
    public void setLastupdateDt(LocalDateTime lastupdateDt) { this.lastupdateDt = lastupdateDt; }

    public String getLastupdateBy() { return lastupdateBy; }
    public void setLastupdateBy(String lastupdateBy) { this.lastupdateBy = lastupdateBy; }
}
