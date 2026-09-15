package com.daifuku.mcm.common;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;

/**
 * 【変換元】CPComponentUtility.vb — DBConcurrencyException 判定に用いていた
 *           ADO.NET TableAdapter の Original_LASTUPDATE_DT 相当
 *
 * 【不具合修正 #296】楽観的排他制御の「排他キー（LASTUPDATE_DT）」を
 *   画面と hidden で往復させるための変換ユーティリティ。
 *
 * <p>移行元VB（ADO.NET TableAdapter）は UPDATE / DELETE の WHERE 句に
 * Original_* 値を並べた楽観的同時実行制御を持っており、更新件数0件のとき
 * DBConcurrencyException → FWM_0009 を表示していた。Java移行時にこの仕組みが
 * 欠落し、同一レコードの同時更新で後勝ち上書き（ロストアップデート）が
 * 発生していたため、同等の仕組みを復活させる。</p>
 *
 * <p>LocalDateTime を th:field でそのまま扱うと Spring の ConversionService の
 * 設定によって書式が変わり得るため、Form 側は String で保持し、書式を本クラスで
 * 固定する。ナノ秒9桁まで持たせ、DBの値（Oracle DATE=秒精度 /
 * SQL Server datetime2=最大7桁）を情報欠落なく往復させる。</p>
 *
 * <p>※ MCM0025U（Mcm0025uController）/ MCM3006U（Mcm3006uController）は
 *   同等の処理を各Controllerに内包しているが、既存画面の挙動へ影響を与えないため
 *   本クラスへの置き換えは行っていない。新規に排他制御を追加する画面
 *   （MCM0011U / MCM0012U）から利用する。</p>
 */
public final class ExclusiveLockKey {

    /** 排他キーを hidden で往復させる際の固定書式 */
    private static final DateTimeFormatter FORMATTER =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss.SSSSSSSSS");

    private ExclusiveLockKey() {
    }

    /** 秒精度のDBでも、更新前と異なる排他キーを保存する。 */
    public static LocalDateTime nextVersion(LocalDateTime previous) {
        LocalDateTime now = LocalDateTime.now().withNano(0);
        return previous == null || now.isAfter(previous) ? now : previous.plusSeconds(1).withNano(0);
    }

    /**
     * DBから読み取った LASTUPDATE_DT を hidden 出力用の文字列へ変換する。
     *
     * @param value DBの LASTUPDATE_DT（未設定の場合は null）
     * @return 固定書式の文字列。null の場合は空文字
     */
    public static String format(LocalDateTime value) {
        return (value == null) ? "" : value.format(FORMATTER);
    }

    /**
     * hidden で往復した排他キーを LocalDateTime へ復元する。
     *
     * <p>{@link #format(LocalDateTime)} が出力した値をそのまま解釈するため、
     * DBの値と等価な LocalDateTime が得られ、DB側の型精度に影響されずに比較できる。</p>
     *
     * <p>空文字・未送信は「DBの LASTUPDATE_DT が未設定」を意味するため null を返す。
     * 書式不正の場合も null を返す。null は {@link #matches(LocalDateTime, LocalDateTime)}
     * で DB現在値（非null）と不一致となり排他エラーになるため、
     * チェックをすり抜けることはない（安全側に倒す）。</p>
     *
     * @param value hidden から受け取った文字列
     * @return 排他キー（未設定・書式不正時は null）
     */
    public static LocalDateTime parse(String value) {
        if (value == null || value.trim().isEmpty()) {
            return null;
        }
        try {
            return LocalDateTime.parse(value.trim(), FORMATTER);
        } catch (DateTimeParseException e) {
            return null;
        }
    }

    /**
     * 排他キー（LASTUPDATE_DT）が一致するか判定する。
     *
     * <p>双方 null（移行データ等で LASTUPDATE_DT が未設定のレコード）の場合は
     * 一致として扱う。この場合は排他判定が実質的に効かないが、
     * 既存の他画面（MCM0022U / MCM0025U / MCM3006U）と同じ挙動に揃える。</p>
     *
     * @param dbValue     DB の現在値
     * @param screenValue 画面が保持していた値
     * @return 一致する場合 true（＝競合していない）
     */
    public static boolean matches(LocalDateTime dbValue, LocalDateTime screenValue) {
        if (dbValue == null) {
            return screenValue == null;
        }
        return dbValue.equals(screenValue);
    }
}
