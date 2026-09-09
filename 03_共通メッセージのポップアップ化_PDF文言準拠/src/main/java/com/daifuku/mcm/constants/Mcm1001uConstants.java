/**
 * 【変換元】Mcm1001uConstant.vb（全体）
 *   元ファイル行数: 約50行
 *   MCM1001U（取引先見積依頼作成検索）の定数定義
 *
 * @author MCM Migration Tool
 * @since v8
 */
package com.daifuku.mcm.constants;

/**
 * MCM1001U 取引先見積依頼作成検索 定数クラス.
 *
 * 【変換元】Mcm1001uConstant.vb
 *   - プラントグリッドのカラムインデックス定数を定義
 *   - VB.NETでは DataGridViewのColumns(name)で参照していたものを定数化
 */
public final class Mcm1001uConstants {

    private Mcm1001uConstants() {
        // インスタンス化禁止
    }

    // ========================================================
    // プラントグリッド カラム名定数
    // 【変換元】Mcm1001uConstant.vb - SUPPORT_ID_NAME 等
    // ========================================================

    /** サポートID列名 (VB: SUPPORT_ID_NAME = "SUPPORT_ID") */
    public static final String COL_SUPPORT_ID = "SUPPORT_ID";

    /** プラント名列名 (VB: PLANT_NK_PLANT = "PLANT_NK") */
    public static final String COL_PLANT_NK = "PLANT_NK";

    /** プラントID列名 (VB: PLANT_ID_PLANT = "PLANT_ID") ※非表示列 */
    public static final String COL_PLANT_ID = "PLANT_ID";

    /** 納入先コード列名 (VB: NONYUSAKI_CD_PLANT = "NONYUSAKI_CD") ※非表示列 */
    public static final String COL_NONYUSAKI_CD = "NONYUSAKI_CD";

    /** 納入先名列名 (VB: NONYUSAKI_NK_PLANT = "NONYUSAKI_NK") ※非表示列 */
    public static final String COL_NONYUSAKI_NK = "NONYUSAKI_NK";

    // ========================================================
    // 納入先グリッド カラム名定数
    // ========================================================

    /** 納入先ID列名 ※非表示列 */
    public static final String COL_NONYUSAKI_ID = "NONYUSAKI_ID";

    /** 納入先カナ名列名 */
    public static final String COL_NONYUSAKIKANA_KN = "NONYUSAKIKANA_KN";

    /** 納入先英名列名 */
    public static final String COL_NONYUSAKIEIMEI_EN = "NONYUSAKIEIMEI_EN";

    /** 納入先工場名列名 */
    public static final String COL_NONYUSAKIKOJO_NK = "NONYUSAKIKOJO_NK";

    /** 旧納入先名列名 */
    public static final String COL_KYUNONYUSAKI_NK = "KYUNONYUSAKI_NK";

    // ========================================================
    // 画面遷移先
    // ========================================================

    /** MCM1002U（需要家見積作成）への遷移先URL */
    public static final String URL_MCM1002U = "/mcm1002u";

    // ========================================================
    // メッセージ
    // ========================================================

    /** 検索結果0件メッセージ */
    public static final String MSG_NO_DATA = "検索結果が1件も存在しません。";

    /** 検索条件未入力メッセージ */
    public static final String MSG_NO_CONDITION = "検索条件は、1項目以上選択して下さい。";
}
