/**
 * 変換元: Mcm1004uTabControl.Designer.vb + Mcm1004uTabControl.vb
 * 期間タブForm
 */
package com.daifuku.mcm.form;

import java.math.BigDecimal;
import java.time.LocalDate;
import org.springframework.format.annotation.DateTimeFormat;
import java.util.ArrayList;
import java.util.List;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class Mcm1004uTabForm implements java.io.Serializable {

    private String tabLabel;
    private Long tmKikanId;
    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE)
    private LocalDate kaisiDt;
    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE)
    private LocalDate syuryoDt;
    private String tmMitsumoriNo;
    @DateTimeFormat(iso = DateTimeFormat.ISO.DATE)
    private LocalDate kaitoDt;
    private String biko;

    private BigDecimal hyojungokeiKin;
    private BigDecimal sikirisyokeiKin;
    private BigDecimal syusseinebikiKin;
    private BigDecimal sikirigokeiKin;

    private String createdDt;
    private String createdBy;
    private String lastupdateDt;
    private String lastupdateBy;

    private List<Mcm1004uTankaRow> tankaRows = new ArrayList<>();
    private List<Mcm1004uTenkenRow> tenkenRows = new ArrayList<>();
}
