/**
 * 【変換元】Mcm0021uConstant.vb（230行）
 *   MCM0021U プラント付替え 定数クラス
 *   元コード: Public Class Mcm0021uConstant
 */
package com.daifuku.mcm.constants;

/**
 * MCM0021U（プラント付替え）画面の定数クラス
 */
public final class Mcm0021uConstants {

    private Mcm0021uConstants() {}

    // =============================================
    // 画面ID
    // =============================================
    /** 画面ID */
    public static final String SCREEN_ID = "MCM0021U";

    /** 画面タイトル */
    public static final String SCREEN_TITLE = "プラント付替え";

    // =============================================
    // リレーション名
    // 【変換元】Mcm0021uConstant.RELATION_NONYUSAKI_KIKIKOSEI 等
    // =============================================
    /** リレーション 納入先マスタ―機器構成パターン */
    public static final String RELATION_NONYUSAKI_KIKIKOSEI = "FK_MCM_MA_NONYUSAKI_MCM_MA_KIKIKOSEI";

    /** リレーション 機器構成パターン―機器明細 */
    public static final String RELATION_KIKIKOSEI_KIKIMEISAI = "FK_MCM_MA_KIKIKOSEI_MCM_MA_KIKIMEISAI";

    // =============================================
    // DataGridView 列名
    // 【変換元】COL_*（DataGridView列のName属性値）
    // =============================================
    /** プラントID 納入先DGV */
    public static final String COL_PLANT_ID_NONYUSAKI = "PLANT_ID_NONYUSAKI_TextBox";

    /** 移動チェック 付替え元DGV */
    public static final String COL_IDOCHECK_MOTO = "IDOCHECK_KIKIKOSEI_MOTO_CheckBox";

    /** 移行先 付替え元DGV */
    public static final String COL_IKOSEI_MOTO = "IRAISAKI_KIKIKOSEI_MOTO_TextBox";

    /** 機器構成ID ループ用機器構成DGV */
    public static final String COL_KIKIKOSEI_ID_LOOP = "KIKIKOSEI_ID_KIKIKOSEI_LOOP_TextBox";

    /** 機器構成名 ループ用機器構成DGV */
    public static final String COL_KIKIKOSEI_NK_LOOP = "KIKIKOSEINK_KIKIKOSEI_LOOP_TextBox";

    /** セット数 ループ用機器構成DGV */
    public static final String COL_SET_NM_LOOP = "SETNM_KIKIKOSEI_LOOP_TextBox";

    /** 単位 ループ用機器構成DGV */
    public static final String COL_TANI_LOOP = "TANI__KIKIKOSEI_LOOP_TextBox";

    /** 手配製番 ループ用機器構成DGV */
    public static final String COL_TEHAISEIBAN_LOOP = "TEHAISEIBAN__KIKIKOSEI_LOOP_TextBox";

    /** コントローラフラグ ループ用機器構成DGV */
    public static final String COL_CONTROLLER_FLG_LOOP = "CONTROLLERFLG_KIKIKOSEI_LOOP_TextBox";

    /** 備考 ループ用機器構成DGV */
    public static final String COL_BIKO_LOOP = "BIKO__KIKIKOSEI_LOOP_TextBox";

    /** 納入先コード 納入先DGV */
    public static final String COL_NONYUSAKI_CD_NONYUSAKI = "NONYUSAKI_CD_NONYUSAKI_TextBox";

    /** 納入先名 納入先DGV */
    public static final String COL_NONYUSAKI_NK_NONYUSAKI = "NONYUSAKI_NK_NONYUSAKI_TextBox";

    /** 納入先工場名 納入先DGV */
    public static final String COL_NONYUSAKIKOJO_NK_NONYUSAKI = "NONYUSAKIKOJO_NK_NONYUSAKI_TextBox";

    /** サポートID 納入先DGV */
    public static final String COL_SUPPORT_ID_NONYUSAKI = "SUPPORT_ID_NONYUSAKI_TextBox";

    /** プラント名 納入先DGV */
    public static final String COL_PLANT_NK_NONYUSAKI = "PLANT_NK_NONYUSAKI_TextBox";

    // =============================================
    // DB項目名（フィールド名）
    // 【変換元】Mcm0021uConstant.BRANDKOSEI_ID 等
    // =============================================
    /** ブランド構成ID */
    public static final String BRANDKOSEI_ID = "BRANDKOSEI_ID";

    /** プラントID */
    public static final String PLANT_ID = "PLANT_ID";

    /** ブランド詳細名 */
    public static final String BRANDSYOSAI_NK = "BRANDSYOSAI_NK";

    /** 移動先ブランド */
    public static final String IKOSAKI = "IKOSAKI";

    /** 機器構成ID */
    public static final String KIKIKOSEI_ID = "KIKIKOSEI_ID";

    /** 機器構成名 */
    public static final String KIKIKOSEI_NK = "KIKIKOSEI_NK";

    /** 品名 */
    public static final String KIKIHINMEI_NK = "KIKIHINMEI_NK";

    /** 型式 */
    public static final String KIKIKATASHIKI = "KIKIKATASHIKI";

    /** 数量 */
    public static final String SURYO_NM = "SURYO_NM";

    /** セット数 */
    public static final String SET_NM = "SET_NM";

    /** 取扱機器ID */
    public static final String ATSUKAIKIKI_ID = "ATSUKAIKIKI_ID";

    /** 備考 */
    public static final String BIKO = "BIKO";

    /** DATA_MAX */
    public static final String DATA_MAX = "DATA_MAX";

    /** 移動チェック */
    public static final String IDO_CHECK = "IDO_CHECK";

    /** ソート条件（機器明細） */
    public static final String MEISAI_SORT = "HYOJIJUN ASC";

    /** 機器構成IDカウント */
    public static final String COUNT_KIKIKOSEI_ID = "COUNT_KIKIKOSEI_ID";

    // =============================================
    // チェックボックスフラグ値
    // =============================================
    /** チェックON */
    public static final int CHECKBOX_FLAG_ON = 1;

    /** チェックOFF */
    public static final int CHECKBOX_FLAG_OFF = 0;

    // =============================================
    // メッセージ定数
    // 【変換元】CPMessage.xml / Mcm0021uScreen.vb
    // =============================================
    /** 検索条件未入力（MSG_0001） */
    public static final String MSG_SEARCH_EMPTY = "検索条件は、1項目以上選択して下さい。";

    /** 検索結果0件（MSG_0002） */
    public static final String MSG_NO_DATA = "検索結果が1件も存在しません。";

    /** プラント未選択 */
    public static final String MSG_NO_PLANT_SELECTED = "プラントが選択されていません。";
    /** #191: VB版 MSG_0124 */
    public static final String MSG_NOT_SEARCHED = "付替え対象が検索されていません。";

    /** 機器情報なし */
    public static final String MSG_NO_KIKI_DATA = "機器構成情報が存在しません。";

    /** 付替え元未設定 */
    public static final String MSG_MOTO_NOT_SET = "付替え元が設定されていません。";

    /** 付替え先未設定 */
    public static final String MSG_SAKI_NOT_SET = "付替え先が設定されていません。";

    /** 付替え元/先が同一 */
    public static final String MSG_SAME_PLANT = "付替え元と付替え先に同じプラントは設定できません。";

    /** 移行先未選択 */
    public static final String MSG_IKOSAKI_NOT_SELECTED = "チェックされた行の付替え先ブランドを選択してください。";

    /** チェック未選択 */
    public static final String MSG_NO_CHECK = "付替え対象を1件以上チェックしてください。";

    /** セット数不一致 */
    public static final String MSG_SET_MISMATCH = "ブランド[%s]の機器構成[%s]でセット数が一致しないため、付替えできません。";

    /** 契約不整合警告 */
    public static final String MSG_KEIYAKU_WARNING = "付替えを行うと、プラントが複数となる契約が発生します。付替え実行しますか？";

    /** 付替え確認 */
    public static final String MSG_REPLACE_CONFIRM = "付替えを実行しますか？";

    /** 付替え完了 */
    public static final String MSG_REPLACE_COMPLETE = "付替えが完了しました。";
}
