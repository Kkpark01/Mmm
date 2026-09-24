/**
 * 【変換元】Mcm1001pConstant.vb
 *   MCM1001P 見積依頼書 - 定数クラス
 *   元ファイル行数: 約60行
 *   Created by: T.Kajimura (2009/02/12)
 *   Updated by: M.Ohsuka (2010/03/23)
 */
package com.daifuku.mcm.constants;

/**
 * MCM1001P 見積依頼書 Excel帳票出力用定数クラス
 * <p>
 * VB.NETの定数値をそのまま保持（1ベース）。
 * Apache POI使用時は各値から1を減算してください。
 * </p>
 */
public final class Mcm1001pConstants {

    private Mcm1001pConstants() {
        // インスタンス化禁止
    }

    // ========================================
    // ヘッダー行列位置（1ベース ※VB.NET元値）
    // ========================================

    /** 行位置：見積依頼NO */
    public static final int HEADER_ROWINDEX_IRAI_NO = 2;
    /** 列位置：見積依頼NO */
    public static final int HEADER_COLINDEX_IRAI_NO = 48;

    /** 行位置：取引先名 */
    public static final int HEADER_ROWINDEX_TORIHIKISAKI_NK = 1;
    /** 列位置：取引先名 */
    public static final int HEADER_COLINDEX_TORIHIKISAKI_NK = 1;

    /** 行位置：納入先名 */
    public static final int HEADER_ROWINDEX_NONYUSAKI_NK = 3;
    /** 列位置：納入先名 */
    public static final int HEADER_COLINDEX_NONYUSAKI_NK = 7;

    /** 行位置：納入先住所1 */
    public static final int HEADER_ROWINDEX_NONYUSAKIJUSYO1_NK = 4;
    /** 列位置：納入先住所1 */
    public static final int HEADER_COLINDEX_NONYUSAKIJUSYO1_NK = 7;

    /** 行位置：納入先住所2 */
    public static final int HEADER_ROWINDEX_NONYUSAKIJUSYO2_NK = 5;
    /** 列位置：納入先住所2 */
    public static final int HEADER_COLINDEX_NONYUSAKIJUSYO2_NK = 7;

    /** 行位置：回答希望日 */
    public static final int HEADER_ROWINDEX_KAITOKIZITSU_DT = 3;
    /** 列位置：回答希望日 */
    public static final int HEADER_COLINDEX_KAITOKIZITSU_DT = 35;

    /** 行位置：依頼日 */
    public static final int HEADER_ROWINDEX_MITSUMORI_DT = 3;
    /** 列位置：依頼日 */
    public static final int HEADER_COLINDEX_MITSUMORI_DT = 48;

    /** 行位置：契約時間帯 */
    public static final int HEADER_ROWINDEX_KEIYAKUJIKANTAI = 4;
    /** 列位置：契約時間帯 */
    public static final int HEADER_COLINDEX_KEIYAKUJIKANTAI = 35;

    /** 行位置：保守方法 */
    public static final int HEADER_ROWINDEX_HOSYUHOHO = 4;
    /** 列位置：保守方法 */
    public static final int HEADER_COLINDEX_HOSYUHOHO = 48;

    /** 行位置：保守点検 */
    public static final int HEADER_ROWINDEX_HOSYUTENKEN = 5;
    /** 列位置：保守点検 */
    public static final int HEADER_COLINDEX_HOSYUTENKEN = 35;

    /** 行位置：点検可能曜日 */
    public static final int HEADER_ROWINDEX_TENKENKANOYOBI = 5;
    /** 列位置：点検可能曜日 */
    public static final int HEADER_COLINDEX_TENKENKANOYOBI = 43;

    /** 行位置：夜間点検 */
    public static final int HEADER_ROWINDEX_YAKANTENKEN = 5;
    /** 列位置：夜間点検 */
    public static final int HEADER_COLINDEX_YAKANTENKEN = 52;

    // ========================================
    // 一覧行列位置（1ベース ※VB.NET元値）
    // ========================================

    /** 開始行位置：見積依頼書明細 */
    public static final int ICHIRAN_ROWINDEX_DEFAULT = 8;

    /** 列位置：製造メーカー名（開始） */
    public static final int ICHIRAN_COLINDEX_SEIZOMAKER_NK = 1;
    /** 列位置：製造メーカー名（終了） */
    public static final int ICHIRAN_COLINDEX_SEIZOMAKER_NK_END = 6;

    /** 列位置：品名（開始） */
    public static final int ICHIRAN_COLINDEX_KIKIHINMEI_NK = 7;
    /** 列位置：品名（終了） */
    public static final int ICHIRAN_COLINDEX_KIKIHINMEI_NK_END = 20;

    /** 列位置：型式（開始） */
    public static final int ICHIRAN_COLINDEX_KIKIKATASHIKI = 21;
    /** 列位置：型式（終了） */
    public static final int ICHIRAN_COLINDEX_KIKIKATASHIKI_END = 32;

    /** 列位置：数量（開始） */
    public static final int ICHIRAN_COLINDEX_SURYO_NM_START = 33;
    /** 列位置：数量 */
    public static final int ICHIRAN_COLINDEX_SURYO_NM = 34;

    /** 列位置：手配製番（開始） */
    public static final int ICHIRAN_COLINDEX_TEHAISEIBAN = 35;
    /** 列位置：手配製番（終了） */
    public static final int ICHIRAN_COLINDEX_TEHAISEIBAN_END = 38;

    /** 列位置：備考（開始） */
    public static final int ICHIRAN_COLINDEX_BIKO = 39;
    /** 列位置：備考（終了） */
    public static final int ICHIRAN_COLINDEX_BIKO_END = 55;

    // ========================================
    // テンプレート・出力ファイル定数
    // ========================================

    /** テンプレートファイル名 */
    public static final String TEMPLATE_FILE_NAME = "MCM1001P_見積依頼書テンプレート.xls";

    /** 出力ファイル名（サフィックス） */
    public static final String CREATE_FILE_NAME = "見積依頼書.xls";

    /** シート名プレフィックス */
    public static final String SHEET_NK = "見積依頼書_";
}
