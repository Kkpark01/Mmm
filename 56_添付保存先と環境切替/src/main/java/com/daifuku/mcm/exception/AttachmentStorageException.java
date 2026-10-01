package com.daifuku.mcm.exception;

/** 共通業務エラーハンドラーで案内する。内部パスはログだけに残す。 */
public class AttachmentStorageException extends McmBusinessException {
    public enum Reason { UNCONFIGURED, ROOT_UNAVAILABLE, ACCESS_DENIED, NOT_FOUND, INVALID_PATH, SAVE_FAILED }
    private final Reason reason;

    public AttachmentStorageException(Reason reason, Throwable cause) {
        super(message(reason), cause);
        this.reason = reason;
    }
    public Reason getReason() { return reason; }

    private static String message(Reason reason) {
        return switch (reason) {
            case UNCONFIGURED -> "添付資料の保存先が未設定です。管理者に確認してください。";
            case ROOT_UNAVAILABLE -> "添付資料の保存先に接続できません。管理者に確認してください。";
            case ACCESS_DENIED -> "添付資料の保存先にアクセスできません。管理者に確認してください。";
            case NOT_FOUND -> "添付資料の実ファイルが見つかりません。管理者に確認してください。";
            case INVALID_PATH -> "添付資料の保存先を確認してください。";
            case SAVE_FAILED -> "添付ファイルを保存できませんでした。管理者に確認してください。";
        };
    }
}
