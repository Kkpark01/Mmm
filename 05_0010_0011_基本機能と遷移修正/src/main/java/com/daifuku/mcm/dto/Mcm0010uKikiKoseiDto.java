/**
 * 【変換元】Mcm0010uDataSet.Designer.vb - MCM_MA_KIKIKOSEI テーブル
 * 元SQL: SELECT SUM(MAF.SURYO_NM * MAE.SET_NM) as SURYO,
 *        MAJ.SEIZOMAKER_NK, MAF.KIKIHINMEI_NK, MAF.KIKIKATASHIKI, MAE.PLANT_ID
 *        FROM MCM_MA_KIKIKOSEI MAE, MCM_MA_KIKIMEISAI MAF,
 *             MCM_MA_ATSUKAIKIKI MAH, MCM_MA_SEIZOMAKER MAJ
 *        WHERE ...
 */
package com.daifuku.mcm.dto;

import lombok.Data;
import java.math.BigDecimal;

@Data
public class Mcm0010uKikiKoseiDto {
    /** プラントID — フィルタ用 */
    private BigDecimal plantId;
    /** 製造メーカー名 — 元: DGV_SEIZOMAKER_NK (HeaderText="製造メーカー") */
    private String seizomakerNk;
    /** 品名 — 元: DGV_KIKIHINMEI_NK (HeaderText="品名") */
    private String kikihinmeiNk;
    /** 型式 — 元: DGV_KIKIKATASHIKI (HeaderText="型式") */
    private String kikikatashiki;
    /** 数量 — 元: DGV_SURYO (HeaderText="数量") */
    private BigDecimal suryo;
}
