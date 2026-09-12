/**
 * 変換元: MCM_TM_TANKADataGridView (32列)
 * 単価グリッド行Form
 */
package com.daifuku.mcm.form;

import lombok.*;
import java.math.BigDecimal;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class Mcm1004uTankaRow implements java.io.Serializable {

    // Frozen ReadOnly
    private String kikikoseiNk;
    private String seizomakerNk;
    private String kikihinmeiNk;
    private String kikikatashiki;
    private BigDecimal suryoNm;

    // Editable
    private BigDecimal hyojunKin;
    private BigDecimal hyojunkeiKin;
    private BigDecimal sikiriKin;
    private BigDecimal sikirikeiKin;
    private String packFlg;
    private String keiyakunaiyo;
    private String keiyakuNo;
    private BigDecimal torihosyujikanId;
    private String tenkenumu;
    private String hosyuhoho;
    private String servicekeitai;
    private String biko;
    private String tehaiseiban;
    private String shokai;

    // Audit ReadOnly
    private String createdDt;
    private String createdBy;
    private String lastupdateDt;
    private String lastupdateBy;

    // Hidden
    private BigDecimal tmKikikoseiId;
    private BigDecimal tmKikimeisaiId;
    private BigDecimal tmTankaId;
    private BigDecimal tmKikanId;
    private BigDecimal tmIraiId;
    private BigDecimal kikikoseiId;
    private BigDecimal kikimeisaiId;
    private BigDecimal atsukaikikiId;
    private BigDecimal setNm;
    private BigDecimal seizomakerId;
}
