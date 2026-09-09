package com.daifuku.mcm.constants;

/**
 * 【変換元】Mcm1006uConstant.vb
 * MCM1006U 取引先契約機器選定 定数クラス
 */
public final class Mcm1006uConstants {

    private Mcm1006uConstants() {}

    // ===== 画面情報 =====
    public static final String SCREEN_ID   = "MCM1006U";
    public static final String SCREEN_NAME = "取引先契約機器選定";

    // ===== Session キー =====
    public static final String SESSION_KEY = "MCM1006U_FORM";

    // ===== 遷移元区分（SeniMotoKbn） =====
    /** 検索一覧から「契約」ボタンによる遷移（新規） */
    public static final int SENIMOTO_INSERT    = 1;
    /** 変更ボタンによる遷移（TAB変更） */
    public static final int SENIMOTO_TAB_HENKO = 2;

    // ===== タブフラグ（TabFlg） =====
    public static final int TAB_FLG_NOW    = 1;  // 現在日付を含む最新タブ
    public static final int TAB_FLG_NON    = 2;  // 最新タブでない
    public static final int TAB_FLG_FUTER  = 3;  // 未来日付の最新タブ
    public static final int TAB_FLG_NEW    = 4;  // 新規追加タブ
    public static final int TAB_FLG_OLD    = 5;  // 過去タブ

    // ===== チェックボックスフラグ =====
    public static final int CHECKBOX_ON  = 1;
    public static final int CHECKBOX_OFF = 0;

    // ===== エラーメッセージ =====
    /** 期間必須 */
    public static final String MSG_ERR_KIKAN_REQUIRED  = "期間（開始日・終了日）を入力してください。";
    /** 期間大小（開始＞終了） */
    public static final String MSG_ERR_KIKAN_RELATION  = "期間の開始日は終了日より前の日付を入力してください。";
    /** 期間1年超過 */
    public static final String MSG_ERR_KIKAN_1YEAR     = "期間は1年以内で設定してください。";
    /** 見積未選択 */
    public static final String MSG_ERR_MITSUMORI_NONE  = "１件もチェックされていません。";
    /** 重複個体管理 */
    public static final String MSG_ERR_KOTAI_DUPLICATE = "選択した見積内に重複した機器が存在します。";
    /** 出精値引き（期間不一致） */
    public static final String MSG_ERR_SYUSSEI_KIKAN   = "選択された見積には、出精値引きが含まれている為、見積の期間と契約の期間が一致する必要があります。\n期間を確認してください。";
    /** 出精値引き（複数期間重複） */
    public static final String MSG_ERR_SYUSSEI_MULTI   = "出精値引きのある見積期間が複数選択されています。";
    /** 出精値引き（合計不一致） */
    public static final String MSG_ERR_SYUSSEI_GOKEI   = "出精値引きの合計金額が一致しません。";

    // ===== 完了メッセージ =====
    public static final String MSG_UPDATE_COMPLETE = "機器選定が完了しました。";
}
