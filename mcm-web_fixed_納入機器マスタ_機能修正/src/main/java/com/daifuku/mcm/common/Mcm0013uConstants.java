/**
 * 【変換元】Mcm0013uConstant.vb（全行）
 * MCM0013U — 納入機器マスタ 定数クラス
 * @since 2026-06-08
 */
package com.daifuku.mcm.common;

public final class Mcm0013uConstants {

    private Mcm0013uConstants() { }

    // ── 画面ID・画面名 ──
    public static final String SCREEN_ID = "MCM0013U";
    public static final String SCREEN_TITLE = "納入機器マスタ";

    // ── 遷移元・遷移先画面 ──
    public static final String CALLER_SCREEN_ID = "MCM0012U";
    public static final String FORWARD_SCREEN_ID = "MCM0014U";

    // ── DataGridView1: 機器構成パターン（MCM_MA_KIKIKOSEI）──
    public static final String COL_KIKIKOSEI_ID      = "KIKIKOSEI_ID";
    public static final String COL_KIKIKOSEI_NK      = "KIKIKOSEI_NK";
    public static final String COL_SET_NM             = "SET_NM";
    public static final String COL_TANI               = "TANI";
    public static final String COL_TEHAISEIBAN        = "TEHAISEIBAN";
    public static final String COL_HYOJIJUN           = "HYOJIJUN";
    public static final String COL_BIKO               = "BIKO";
    public static final String COL_CONTROLLER_FLG     = "CONTROLLER_FLG";
    public static final String COL_SENTAKU            = "SENTAKU";

    // ── DataGridView2: 構成明細（MCM_MA_KIKIMEISAI）── ReadOnly
    public static final String COL_KIKIMEISAI_ID      = "KIKIMEISAI_ID";
    public static final String COL_KIKIHINMEI_NK      = "KIKIHINMEI_NK";
    public static final String COL_KIKIKATASHIKI      = "KIKIKATASHIKI";
    public static final String COL_SURYO_NM           = "SURYO_NM";
    public static final String COL_SEIZOMAKER_NK      = "SEIZOMAKER_NK";
    public static final String COL_KIKIBUNRUI_NK      = "KIKIBUNRUI_NK";
    public static final String COL_KOTAIKANRI_FLG     = "KOTAIKANRI_FLG";
    public static final String COL_NOUNYU_KBN         = "NOUNYU_KBN";

    // ── DataGridView3: 機器個体管理（MCM_MA_KIKIKOTAIKANRI）──
    public static final String COL_KOTAIKANRI_ID      = "KOTAIKANRI_ID";
    public static final String COL_KOTAI_NK           = "KOTAI_NK";
    public static final String COL_SERIAL_NO          = "SERIAL_NO";
    public static final String COL_ITIJINONYU_DT      = "ITIJINONYU_DT";
    public static final String COL_SETCHIBASYO        = "SETCHIBASYO";
    public static final String COL_TEKKYO_DT          = "TEKKYO_DT";
    public static final String COL_KEIYAKUKIGEN_DT    = "KEIYAKUKIGEN_DT";
    public static final String COL_ENCHOKEIYAKUKIGEN_DT = "ENCHOKEIYAKUKIGEN_DT";
    public static final String COL_BRANDKOSEI_ID      = "BRANDKOSEI_ID";
    public static final String COL_ATSUKAIKIKI_NK     = "ATSUKAIKIKI_NK";
    public static final String COL_KATASHIKI           = "KATASHIKI";
    public static final String COL_UPSKOKAN_DT        = "UPSKOKAN_DT";
    public static final String COL_TORIHIKISAKI_NK    = "TORIHIKISAKI_NK";
    public static final String COL_ATSUKAIKIKIKOSEI_ID = "ATSUKAIKIKIKOSEI_ID";
    public static final String COL_KEIYAKUMANRYOYOTEI_DT = "KEIYAKUMANRYOYOTEI_DT";

    // ── 行ステータス ──
    public static final String ROW_STATUS_NEW       = "new";
    public static final String ROW_STATUS_MODIFIED  = "modified";
    public static final String ROW_STATUS_DELETED   = "deleted";
    public static final String ROW_STATUS_UNCHANGED = "unchanged";

    // ── コントローラフラグ ──
    public static final String CONTROLLER_FLG_YES = "1";
    public static final String CONTROLLER_FLG_NO  = "0";

    // ── メッセージ（VB CPMessageConstant 準拠）──
    public static final String MSG_CONFIRM_UPDATE = "登録を実施します。よろしいですか？";
    public static final String MSG_UPDATE_SUCCESS = "登録を完了しました。";
    public static final String MSG_DELETE_REF_ERROR = "この機器構成は他のデータから参照されているため削除できません";
    public static final String MSG_SETNUM_KOTAI_MISMATCH = "機器構成パターン情報のセット数と機器個体情報の数量が一致しません。";
    public static final String MSG_CONTROLLER_DUPLICATE = "コントローラフラグが複数行で設定されています";
    public static final String MSG_KOTAI_REF_TM = "見積または契約情報として、既に使用されている為、削除する事が出来ません。";
    public static final String MSG_KOTAI_REF_UK = "見積または契約情報として、既に使用されている為、削除する事が出来ません。";
    public static final String MSG_DELETE_KOTAI_REF = "機器個体管理テーブルで、既に使用されている為、削除する事が出来ません。";
    public static final String MSG_DELETE_MEISAI_REF = "機器明細テーブルで、既に使用されている為、削除する事が出来ません。";

    // ★ Part 2 追加: D-008 行追加時必須チェック
    // 【変換元】VB CPMessageConstant.MSG_0054
    //   元コード: If IsNullOrEmpty(row.KIKIKOSEI_NK) Or row.SET_NM Is Nothing Or
    //             IsNullOrEmpty(row.TANI) Then ShowMessage(MSG_0054) : Exit Sub
    public static final String MSG_REQUIRED_FIELDS_EMPTY =
        "必須項目を入力して下さい。";

    // ★ Part 2 追加: D-005 セット数チェック
    // 【変換元】VB CPMessageConstant.MSG_0050
    public static final String MSG_SETNUM_INVALID =
        "セット数は1以上の値を入力して下さい。";

    // ★ Part 2 追加: D-005 構成明細・個体の存在チェック
    // 【変換元】VB CPMessageConstant.MSG_0051
    public static final String MSG_KOSEI_MEISAI_KOTAI_REQUIRED =
        "機器構成パターン情報のセット数と機器個体情報の数量が一致しません。";

    // ★ Part 2 追加: コントローラフラグ整合チェック
    // 【変換元】VB CPMessageConstant.MSG_0091
    public static final String MSG_CONTROLLER_FLG_MISMATCH_91 =
        "機器構成情報のコントローラが有効ではありません。";

    // 【変換元】VB CPMessageConstant.MSG_0092
    public static final String MSG_CONTROLLER_FLG_MISMATCH_92 =
        "機器構成情報のコントローラを無効にして下さい。";

    // 【変換元】VB CPMessageConstant.MSG_0093
    public static final String MSG_KOTAI_MEISAI_MISMATCH =
        "機器明細情報と一致していないデータが個体管理情報に存在します。";

    // 【変換元】VB CPMessageConstant.MSG_0055
    public static final String MSG_KOTAI_DELETE_SETNUM =
        "セット数と一致しない為、削除できません。";

    // 【変換元】VB CPMessageConstant.MSG_0009
    public static final String MSG_DELETE_IN_USE_PREFIX =
        "で使用されているため削除できません。";
    public static final String MSG_NO_CHANGE = "値が変更されていません。";
    public static final String MSG_NO_SELECTION = "行が選択されていません。";
    public static final String MSG_CONFIRM_DELETE = "行を削除します。よろしいですか？";
    public static final String MSG_CONFIRM_DISCARD = "データが変更されています。破棄されますがよろしいですか？";
}
