/**
 * 変換元: Mcm1004uScreen.Designer.vb + Mcm1004uDelivery.vb
 * 取引先見積内容 メインForm
 */
package com.daifuku.mcm.form;

import lombok.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class Mcm1004uForm implements java.io.Serializable {

    // 変換元: Mcm1004uDelivery
    private BigDecimal tmKeiyakujikanId;
    private int seniMotoKbn;

    // ヘッダ部 (ReadOnly) - 変換元: MCM_TM_MITSUMORIBindingSource
    private BigDecimal tmIraiId;
    private String tmIraiNo;
    private String torihikisakiNk;
    private String nonyusakiNk;
    private String keiyakujikantai;
    private String supportId;
    private String mitsumoriDt;
    private String plantNk;
    private String iraijigyosyoNk;
    private String iraitantosya;
    private String jotai;
    private String jotaiDisplay;
    private BigDecimal plantId;
    private BigDecimal torihikisakiId;

    // 監査情報
    private String createdDt;
    private String createdBy;
    private String lastupdateDt;
    private String lastupdateBy;

    // タブデータ
    private List<Mcm1004uTabForm> tabs = new ArrayList<>();

    // 表示制御
    private String revision;
    private int selectedIndex;
    private java.util.Map<String,String> hoshuOptions = new java.util.LinkedHashMap<>();
    private boolean viewFlg;
    private boolean lockReleaseFlg;
    private boolean dataNotFound;
}
