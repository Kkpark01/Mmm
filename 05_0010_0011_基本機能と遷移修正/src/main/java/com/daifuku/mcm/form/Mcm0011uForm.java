/**
 * 【変換元】Mcm0011uScreen.Designer.vb - 入力パネルのコントロール群
 * 元コントロール対応:
 *   NONYUSAKI_CDTextBox    → nonyusakiCd    (Required, MaxLength=12)
 *   NONYUSAKI_NKTextBox    → nonyusakiNk    (Required, MaxLength=80)
 *   KYUNONYUSAKI_NKTextBox → kyunonyusakiNk (MaxLength=80)
 *   NONYUSAKIKANA_KNTextBox→ nonyusakikanaKn(Required, MaxLength=60)
 *   NONYUSAKIEIMEI_ENTextBox→nonyusakieimeiEn(MaxLength=40)
 *   NONYUSAKIKOJO_NKTextBox→ nonyusakikojoNk(MaxLength=80)
 *   YUBIN_NOTextBox        → yubinNo        (MaxLength=8)
 *   KUNI_NKTextBox         → kuniNk         (MaxLength=20)
 *   JUSYO1_NKTextBox       → jusyo1Nk       (Required, MaxLength=80)
 *   JUSYO2_NKTextBox       → jusyo2Nk       (MaxLength=80)
 *   TEL_NOTextBox          → telNo          (MaxLength=50)
 *   FAX_NOTextBox          → faxNo          (MaxLength=50)
 *   BIKOTextBox            → biko           (MaxLength=4000, Multiline)
 *   DTSMasutaRenkeiCheckBox→ dtsrenkeiFlg   (CheckState binding, always disabled)
 *   NonyusakiSakujoCheckBox→ nonyusakiSakujo(delete flag)
 */
package com.daifuku.mcm.form;

import java.math.BigDecimal;

import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.Pattern;
import jakarta.validation.constraints.Size;

import lombok.Data;

@Data
public class Mcm0011uForm implements java.io.Serializable {
    private static final long serialVersionUID = 1L;

    /** 読込時の全カラム比較用。サーバーのセッション内で保持する。 */
    private String revision;


    /** 納入先ID（hidden） — 新規時は0、編集時はDB値 */
    private BigDecimal nonyusakiId;

    /** 納入先コード — Required, MaxLength=12 */
    @NotBlank(message = "納入先コードは必須です。")
    @Size(max = 12, message = "納入先コードは12桁以下で入力してください。")
    @Pattern(regexp = "^[\\x20-\\x7E]*$", message = "納入先コードは半角で入力してください。")
    private String nonyusakiCd;

    /** 納入先名 — Required, MaxLength=80 */
    @NotBlank(message = "納入先名は必須です。")
    @Size(max = 80, message = "納入先名は80桁以下で入力してください。")
    private String nonyusakiNk;

    /** 旧納入先名 — MaxLength=80 */
    @Size(max = 80, message = "旧納入先名は80桁以下で入力してください。")
    private String kyunonyusakiNk;

    /** 納入先カナ名 — Required, MaxLength=60 */
    @NotBlank(message = "納入先カナ名は必須です。")
    @Size(max = 60, message = "納入先カナ名は60桁以下で入力してください。")
    private String nonyusakikanaKn;

    /** 納入先英名 — MaxLength=40 */
    @Size(max = 40, message = "納入先英名は40桁以下で入力してください。")
    private String nonyusakieimeiEn;

    /** 納入先工場名 — MaxLength=80 */
    @Size(max = 80, message = "納入先工場名は80桁以下で入力してください。")
    private String nonyusakikojoNk;

    /** 郵便番号 — MaxLength=8 */
    @Size(max = 8, message = "郵便番号は8桁以下で入力してください。")
    private String yubinNo;

    /** 国名 — MaxLength=20 */
    @Size(max = 20, message = "国名は20桁以下で入力してください。")
    private String kuniNk;

    /** 住所1 — Required, MaxLength=80 */
    @NotBlank(message = "住所1は必須です。")
    @Size(max = 80, message = "住所1は80桁以下で入力してください。")
    private String jusyo1Nk;


    /** 住所2 — MaxLength=80 */
    @Size(max = 80, message = "住所2は80桁以下で入力してください。")
    private String jusyo2Nk;

    /** 電話番号 — MaxLength=50 */
    @Size(max = 50, message = "電話番号は50桁以下で入力してください。")
    private String telNo;

    /** FAX番号 — MaxLength=50 */
    @Size(max = 50, message = "FAX番号は50桁以下で入力してください。")
    private String faxNo;

    /** 備考 — MaxLength=4000, Multiline */
    @Size(max = 4000, message = "備考は4000桁以下で入力してください。")
    private String biko;

    /** DTS連携フラグ — VB: CheckState binding to DTSRENKEI_FLG(Decimal), 画面上は常に無効 */
    private BigDecimal dtsrenkeiFlg;

    /** 納入先削除チェック — VB: NonyusakiSakujoCheckBox */
    private boolean nonyusakiSakujo;

    /** 新規モードフラグ（画面制御用） */
    private boolean newMode;

}
