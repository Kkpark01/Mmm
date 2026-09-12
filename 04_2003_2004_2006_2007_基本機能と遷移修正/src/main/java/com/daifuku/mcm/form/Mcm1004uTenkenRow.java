/**
 * 変換元: MCM_TM_TENKENDataGridView (19列)
 * 点検グリッド行Form
 */
package com.daifuku.mcm.form;

import lombok.*;
import java.math.BigDecimal;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class Mcm1004uTenkenRow implements java.io.Serializable {

    // Frozen ReadOnly
    private String kikikoseiNk;
    private String kikihinmeiNk;
    private String kikikatashiki;

    // Editable
    private BigDecimal tenkenkaisu;
    private String tenkenkanoyobi;
    private String yakantaioumu;
    private String biko;

    // Audit ReadOnly
    private String createdDt;
    private String createdBy;
    private String lastupdateDt;
    private String lastupdateBy;

    // Hidden
    private BigDecimal tmTenkenId;
    private BigDecimal tmKikikoseiId;
    private BigDecimal tmKikanId;
    private BigDecimal kikikoseiId;
    private BigDecimal kikimeisaiId;
    private String oyakikibunruiCd;
    private BigDecimal maeHyojijun;
    private BigDecimal mafHyojijun;
}
