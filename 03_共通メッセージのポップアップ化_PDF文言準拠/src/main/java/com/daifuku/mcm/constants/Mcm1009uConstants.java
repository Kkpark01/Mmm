package com.daifuku.mcm.constants;

/**
 * 【変換元】Mcm1009uConstant.vb
 *   MCM1009U パック契約・解約 画面の定数クラス
 *   元ファイル行数: 約50行
 */
public final class Mcm1009uConstants {

    private Mcm1009uConstants() {}

    // ========================================
    // 画面ID・タイトル
    // ========================================
    /** 画面ID */
    public static final String SCREEN_ID = "MCM1009U";
    /** 画面タイトル */
    public static final String SCREEN_TITLE = "パック契約・解約";

    // ========================================
    // PACK_FLG 固定条件
    // ========================================
    /** パック契約固定条件: PACK_FLG = 1 */
    public static final int PACK_FLG_ON = 1;

    // ========================================
    // リンク定数 — VB: Mcm1009uConstant
    // ========================================
    /** 依頼NOリンク — VB: TM_IRAI_NO_MCM_1003_V_Link */
    public static final String TM_IRAI_NO_LINK = "tm_irai_no_link";
    /** 契約NOリンク — VB: KEIYAKU_NO_MCM_1003_V_Link */
    public static final String KEIYAKU_NO_LINK = "keiyaku_no_link";
    /** 契約作成リンク — VB: KEIYAKU_KEIYAKU_MCM_1003_V_Link */
    public static final String KEIYAKU_KEIYAKU_LINK = "keiyaku_keiyaku_link";
    /** 契約破棄リンク — VB: KEIYAKU_DEL_MCM_1003_V_Link */
    public static final String KEIYAKU_DEL_LINK = "keiyaku_del_link";

    // ========================================
    // 遷移先画面ID
    // ========================================
    /** MCM1004U（需要家契約一覧） */
    public static final String FORWARD_MCM1004U = "MCM1004U";
    /** MCM1005U（需要家契約管理） */
    public static final String FORWARD_MCM1005U = "MCM1005U";
    /** MCM1006U1（需要家契約照会・機器設定） */
    public static final String FORWARD_MCM1006U1 = "MCM1006U1";

    // ========================================
    // 遷移元区分
    // ========================================
    /** 遷移元区分：見積リンク */
    public static final String SENIMOTO_MITSUMORI_LINK = "MITSUMORI_LINK";
    /** 遷移元区分：契約リンク */
    public static final String SENIMOTO_KEIYAKU_LINK = "KEIYAKU_LINK";
    /** 遷移元区分：新規登録 */
    public static final String SENIMOTO_INSERT = "INSERT";

    // ========================================
    // セッションキー
    // ========================================
    public static final String SESSION_KEY_MCM1004U = "mcm1009u.forward.mcm1004u";
    public static final String SESSION_KEY_MCM1005U = "mcm1009u.forward.mcm1005u";
    public static final String SESSION_KEY_MCM1006U1 = "mcm1009u.forward.mcm1006u1";
    public static final String SESSION_KEY_SEARCH_FORM = "mcm1009u.searchForm";

    // ========================================
    // 状態値
    // ========================================
    public static final String JOTAI_MITSUMORI = "見積";
    public static final String JOTAI_KEIYAKU = "契約";
    public static final String JOTAI_KAIYAKU = "解約";
    public static final String JOTAI_HAKI = "破棄";

    // ========================================
    // メッセージ
    // ========================================
    public static final String MSG_SEARCH_CONDITION_EMPTY = "検索条件は、1項目以上選択して下さい。";
    public static final String MSG_SEARCH_NO_RESULT = "検索結果が1件も存在しません。";
    public static final String MSG_CONFIRM_CANCEL_CONTRACT = "契約を破棄します。よろしいですか？";
    public static final String MSG_PAYMENT_ALREADY_EXISTS = "取引先への支払が発生している為、破棄する事が出来ません。";
    public static final String MSG_CANCEL_COMPLETE = "契約を破棄しました。";
}