package com.daifuku.mcm.constants;

/**
 * 【変換元】Mcm1003pConstant.vb - MCM1003P 契約手続依頼書の定数クラス
 * 【変換元行数】約200行
 * 【作成者】T.Matsui (2009/03/04)
 *
 * Excelテンプレートの行・列位置定数、テンプレートファイル名、
 * 依頼区分文字列、明細区分文字列、タイトル文字列を定義。
 */
public final class Mcm1003pConstants {

    private Mcm1003pConstants() {
        // ユーティリティクラスのためインスタンス化禁止
    }

    // ===================================================================
    // テンプレート・出力ファイル
    // ===================================================================

    /** テンプレートファイル名 */
    public static final String TEMPLATE_FILE_NAME = "MCM1003P_契約手続依頼書.xls";

    /** 作成ファイル名（ダウンロード時のファイル名） */
    public static final String CREATE_FILE_NAME = "契約手続依頼書.xls";

    // ===================================================================
    // 明細区分文字列（更新の場合に明細一覧の備考欄に表示する情報）
    // ===================================================================

    /** 明細区分：契約 */
    public static final String MEISAI_KUBUN_KEIYAKU = "契約";
    /** 明細区分：契約(継続) */
    public static final String MEISAI_KUBUN_KEIZOKU = "契約(継続）";
    /** 明細区分：解約 */
    public static final String MEISAI_KUBUN_KAIYAKU = "解約";

    // ===================================================================
    // 依頼区分
    // ===================================================================

    /** 依頼区分：新規契約 */
    public static final String KUBUN_INSERT = "(■新規契約 □解約 □更新)";
    /** 依頼区分：更新 */
    public static final String KUBUN_UPDATE = "(□新規契約 □解約 ■更新)";
    /** 依頼区分：解約 */
    public static final String KUBUN_DELETE = "(□新規契約 ■解約 □更新)";

    // ===================================================================
    // ヘッダー名（更新時のタイトル表示用）
    // ===================================================================

    /** タイトル：新契約 */
    public static final String TITLE_KEIYAKU = "新契約";
    /** タイトル：旧契約 */
    public static final String TITLE_KAIYAKU = "旧契約";

    // ===================================================================
    // ヘッダー位置（1-based） ※POI使用時は -1 して0-basedに変換
    // ===================================================================

    /** 行位置：区分 */
    public static final int KEIYAKU_ROWINDEX_KUBUN = 1;
    /** 列位置：区分 */
    public static final int KEIYAKU_COLINDEX_KUBUN = 5;
    /** 行位置：作成日 */
    public static final int KEIYAKU_ROWINDEX_CREATED_DT = 1;
    /** 列位置：作成日 */
    public static final int KEIYAKU_COLINDEX_CREATED_DT = 11;
    /** 行位置：ページ番号 */
    public static final int KEIYAKU_ROWINDEX_PAGE_NO = 2;
    /** 列位置：ページ番号 */
    public static final int KEIYAKU_COLINDEX_PAGE_NO = 11;

    // ===================================================================
    // 取引先位置（1-based）
    // ===================================================================

    /** 行位置：取引先名 */
    public static final int KEIYAKU_ROWINDEX_TORIHIKISAKI_NK = 3;
    /** 列位置：取引先名 */
    public static final int KEIYAKU_COLINDEX_TORIHIKISAKI_NK = 3;
    /** 行位置：ご担当者様 */
    public static final int KEIYAKU_ROWINDEX_TORISYUTANTOSYA_NK = 4;
    /** 列位置：ご担当者様 */
    public static final int KEIYAKU_COLINDEX_TORISYUTANTOSYA_NK = 3;

    // ===================================================================
    // 顧客情報位置（1-based）
    // ===================================================================

    /** 行位置：納入先 */
    public static final int KEIYAKU_ROWINDEX_NONYUSAKI_NK = 6;
    /** 列位置：納入先 */
    public static final int KEIYAKU_COLINDEX_NONYUSAKI_NK = 3;
    /** 行位置：住所 */
    public static final int KEIYAKU_ROWINDEX_NONYUSAKIJUSYO1_NK = 7;
    /** 列位置：住所 */
    public static final int KEIYAKU_COLINDEX_NONYUSAKIJUSYO1_NK = 3;
    /** 行位置：電話番号 */
    public static final int KEIYAKU_ROWINDEX_NONYUTEL_NO = 8;
    /** 列位置：電話番号 */
    public static final int KEIYAKU_COLINDEX_NONYUTEL_NO = 3;
    /** 行位置：担当部署 */
    public static final int KEIYAKU_ROWINDEX_NONYUBUSYO_NK = 9;
    /** 列位置：担当部署 */
    public static final int KEIYAKU_COLINDEX_NONYUBUSYO_NK = 3;
    /** 行位置：担当者 */
    public static final int KEIYAKU_ROWINDEX_NONYUTANTOSYA_NK = 9;
    /** 列位置：担当者 */
    public static final int KEIYAKU_COLINDEX_NONYUTANTOSYA_NK = 6;
    /** 行位置：サポートID */
    public static final int KEIYAKU_ROWINDEX_SUPPORT_ID = 10;
    /** 列位置：サポートID */
    public static final int KEIYAKU_COLINDEX_SUPPORT_ID = 3;

    // ===================================================================
    // 契約セクション位置（1-based）
    // ===================================================================

    /** 行位置：契約タイトル */
    public static final int KEIYAKU_ROWINDEX_TITLE = 12;
    /** 列位置：契約タイトル */
    public static final int KEIYAKU_COLINDEX_TITLE = 1;
    /** 行位置：見積金額（仕切） */
    public static final int KEIYAKU_ROWINDEX_SIKIRI_KEIYAKU = 12;
    /** 列位置：見積金額（仕切） */
    public static final int KEIYAKU_COLINDEX_SIKIRI_KEIYAKU = 3;
    /** 行位置：支払月 */
    public static final int KEIYAKU_ROWINDEX_SIHARAI_KEIYAKU = 12;
    /** 列位置：支払月 */
    public static final int KEIYAKU_COLINDEX_SIHARAI_KEIYAKU = 6;
    /** 行位置：契約開始日 */
    public static final int KEIYAKU_ROWINDEX_KEIYAKUKAISI_DT = 13;
    /** 列位置：契約開始日 */
    public static final int KEIYAKU_COLINDEX_KEIYAKUKAISI_DT = 3;
    /** 行位置：見積No */
    public static final int KEIYAKU_ROWINDEX_TM_MITSUMORI_NO = 13;
    /** 列位置：見積No */
    public static final int KEIYAKU_COLINDEX_TM_MITSUMORI_NO = 6;
    /** 行位置：契約時間帯 */
    public static final int KEIYAKU_ROWINDEX_KEIYAKUJIKANTAI = 14;
    /** 列位置：契約時間帯 */
    public static final int KEIYAKU_COLINDEX_KEIYAKUJIKANTAI = 3;
    /** 行位置：手配製番 */
    public static final int KEIYAKU_ROWINDEX_TEHAISEIBAN = 15;
    /** 列位置：手配製番 */
    public static final int KEIYAKU_COLINDEX_TEHAISEIBAN = 3;
    /** 行位置：備考（契約） */
    public static final int KEIYAKU_ROWINDEX_BIKO_KEIYAKU = 16;
    /** 列位置：備考（契約） */
    public static final int KEIYAKU_COLINDEX_BIKO_KEIYAKU = 3;

    // ===================================================================
    // 解約セクション位置（1-based）
    // ===================================================================

    /** 行位置：解約タイトル */
    public static final int KAIYAKU_ROWINDEX_TITLE = 18;
    /** 列位置：解約タイトル */
    public static final int KAIYAKU_COLINDEX_TITLE = 1;
    /** 行位置：見積金額（仕切・解約） */
    public static final int KAIYAKU_ROWINDEX_SIKIRI_KAIYAKU = 18;
    /** 列位置：見積金額（仕切・解約） */
    public static final int KAIYAKU_COLINDEX_SIKIRI_KAIYAKU = 3;
    /** 行位置：支払月（解約） */
    public static final int KAIYAKU_ROWINDEX_SIHARAI_KAIYAKU = 18;
    /** 列位置：支払月（解約） */
    public static final int KAIYAKU_COLINDEX_SIHARAI_KAIYAKU = 6;
    /** 行位置：解約日 */
    public static final int KAIYAKU_ROWINDEX_KAIYAKU_DT = 19;
    /** 列位置：解約日 */
    public static final int KAIYAKU_COLINDEX_KAIYAKU_DT = 3;
    /** 行位置：契約No */
    public static final int KAIYAKU_ROWINDEX_KEIYAKU_NO = 19;
    /** 列位置：契約No */
    public static final int KAIYAKU_COLINDEX_KEIYAKU_NO = 6;
    /** 行位置：備考（解約） */
    public static final int KAIYAKU_ROWINDEX_BIKO_KAIYAKU = 20;
    /** 列位置：備考（解約） */
    public static final int KAIYAKU_COLINDEX_BIKO_KAIYAKU = 3;

    // ===================================================================
    // 返金情報位置（1-based）
    // ===================================================================

    /** 行位置：契約期間 */
    public static final int HENKIN_ROWINDEX_KEIYAKUKIKAN = 21;
    /** 列位置：契約期間 */
    public static final int HENKIN_COLINDEX_KEIYAKUKIKAN = 4;
    /** 行位置：返金期間 */
    public static final int HENKIN_ROWINDEX_HENKINKIKAN = 22;
    /** 列位置：返金期間 */
    public static final int HENKIN_COLINDEX_HENKINKIKAN = 4;
    /** 行位置：金額 */
    public static final int HENKIN_ROWINDEX_KINGAKU = 23;
    /** 列位置：金額 */
    public static final int HENKIN_COLINDEX_KINGAKU = 4;
    /** 行位置：備考（返金） */
    public static final int HENKIN_ROWINDEX_BIKO_HENKIN = 22;
    /** 列位置：備考（返金） */
    public static final int HENKIN_COLINDEX_BIKO_HENKIN = 5;

    // ===================================================================
    // 明細セクション定数（1-based）
    // ===================================================================

    // --- 1シート目（ヘッダー+明細 最大20行） ---

    /** 1シート目 明細開始行（1-based） */
    public static final int STARTINDEX_MEISAIFST = 27;
    /** 1シート目 明細終了行（1-based） */
    public static final int ENDTINDEX_MEISAIFST = 46;
    /** 1シート目 明細行数 */
    public static final int ROWCOUNT_MEISAIFST = 20;

    // --- 2シート目以降（明細のみ 最大50行） ---

    /** 2シート目以降 明細開始行（1-based） */
    public static final int STARTINDEX_MEISAISEC = 3;
    /** 2シート目以降 明細終了行（1-based） */
    public static final int ENDINDEX_MEISAISEC = 52;
    /** 2シート目以降 明細行数 */
    public static final int ROWCOUNT_MEISAISEC = 50;

    // ===================================================================
    // 明細列位置（1-based）
    // ===================================================================

    /** 列位置：表示順 */
    public static final int MEISAI_COLINDEX_HYOJIJUN = 2;
    /** 列位置：機器品名 */
    public static final int MEISAI_COLINDEX_KIKIHINMEI_NK = 3;
    /** 列位置：機器型式 */
    public static final int MEISAI_COLINDEX_KIKIKATASHIKI = 5;
    /** 列位置：数量 */
    public static final int MEISAI_COLINDEX_SURYO_NM = 7;
    /** 列位置：点検回数 */
    public static final int MEISAI_COLINDEX_TENKENKAISU = 8;
    /** 列位置：点検月 */
    public static final int MEISAI_COLINDEX_TENKEN_TSUKI = 9;
    /** 列位置：備考（明細） */
    public static final int MEISAI_COLINDEX_BIKO_MEISAI = 10;

    // ===================================================================
    // 2シート目以降のページ番号位置（1-based）
    // ===================================================================

    /** 2シート目以降 行位置：ページ番号 */
    public static final int MEISAISEC_ROWINDEX_PAGE_NO = 1;
    /** 2シート目以降 列位置：ページ番号 */
    public static final int MEISAISEC_COLINDEX_PAGE_NO = 11;
}
