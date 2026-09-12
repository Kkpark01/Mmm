/**
 * 【変換元】Mcm0010uScreen.Designer.vb - 検索パネルのコントロール群
 * NONYUSAKI_CDTextBox → nonyusakiCd
 * NONYUSAKI_NKTextBox → nonyusakiNk  (旧名・カナなどを含む横断検索)
 * ETCTextBox          → jusyo        (VB: ETCTextBox = その他(住所など))
 * SUPPORT_IDTextBox   → supportId
 * PLANT_NKTextBox     → plantNk
 */
package com.daifuku.mcm.form;

import lombok.Data;

@Data
public class Mcm0010uForm implements java.io.Serializable {
    /** 納入先コード — 元: NONYUSAKI_CDTextBox (MaxLength=12) */
    private String nonyusakiCd;
    /** 納入先名(旧名・カナなど) — 元: NONYUSAKI_NKTextBox */
    private String nonyusakiNk;
    /** その他(住所など) — 元: ETCTextBox — 郵便番号/住所/国名/TEL/FAX/備考を横断検索 */
    private String jusyo;
    /** サポートID — 元: SUPPORT_IDTextBox (MaxLength=7) */
    private String supportId;
    /** プラント名 — 元: PLANT_NKTextBox */
    private String plantNk;
}
