package com.daifuku.mcm.common;

/**
 * 【変換元】CPValidateUtility.vb (FRAMEWORK/UTILITY/CPValidateUtility.vb 約400行)
 *           CPValidate.xml (FRAMEWORK/XML/CPValidate.xml)
 *
 * バリデーションサービス
 * VB.NETの CPValidateUtility（自動入力チェック）を
 * Spring Service として再実装。
 *
 * 変換マッピング:
 *   CPValidateUtility.IsErrorCell()      → validateField()
 *   CPValidateUtility.ValidateAuto()     → validateForm()
 *   CPValidateUtility.CheckValidate()    → validateGrid()
 *   CPValidate.xml (バリデーション定義)   → Bean Validation + このServiceの補完
 *
 * 注意: Bean Validation (@NotBlank, @Size等) の標準アノテーションを優先し、
 *       VB.NET固有のチェック（半角/全角、バイト長等）はこのServiceで補完する。
 */

import org.springframework.context.MessageSource;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

@Service
public class ValidationService {

    private final MessageSource messageSource;

    public ValidationService(MessageSource messageSource) {
        this.messageSource = messageSource;
    }

    /**
     * 【変換元】CPValidateUtility.vb - IsErrorCell()
     *   元コード: 個別フィールドの入力チェック
     */
    public String validateField(String fieldName, String value,
                                boolean required, String type,
                                int maxLength, int minLength) {
        // 必須チェック
        if (required && CommonUtils.isNullOrEmpty(value)) {
            return getMessage("MSG_0001E", fieldName);
        }
        if (CommonUtils.isNullOrEmpty(value)) return null;

        // 最小桁数チェック
        if (minLength > 0 && value.length() < minLength) {
            return getMessage("MSG_0006E", fieldName, String.valueOf(minLength));
        }
        // 最大桁数チェック（バイト長）
        if (maxLength > 0 && CommonUtils.getByteLength(value) > maxLength) {
            return getMessage("MSG_0007E", fieldName, String.valueOf(maxLength));
        }
        // 型チェック
        if (type != null) {
            switch (type) {
                case AppConstants.VALIDATE_TYPE_NUMERIC:
                    if (!CommonUtils.isNumeric(value)) {
                        return getMessage("MSG_0008E", fieldName);
                    }
                    break;
                case AppConstants.VALIDATE_TYPE_DATE:
                    if (!CommonUtils.isDate(value)) {
                        return getMessage("MSG_0009E", fieldName);
                    }
                    break;
                case AppConstants.VALIDATE_TYPE_HALFWIDTH:
                    if (!CommonUtils.isHalfWidth(value)) {
                        return getMessage("MSG_0010E", fieldName);
                    }
                    break;
                case AppConstants.VALIDATE_TYPE_ALPHANUMERIC:
                    if (!CommonUtils.isAlphaNumeric(value)) {
                        return getMessage("MSG_0011E", fieldName);
                    }
                    break;
                case AppConstants.VALIDATE_TYPE_EMAIL:
                    if (!CommonUtils.isEmail(value)) {
                        return getMessage("MSG_0012E", fieldName);
                    }
                    break;
                default:
                    break;
            }
        }
        return null;
    }

    /** VB CPValidate.xmlの任意数値項目の上下限（境界を含む）。既存validateFieldの動作は変更しない。 */
    public String validateNumericRange(String fieldName, String value,
                                      java.math.BigDecimal minimum, java.math.BigDecimal maximum) {
        String error = validateField(fieldName, value, false, AppConstants.VALIDATE_TYPE_NUMERIC, 0, 0);
        if (error != null || CommonUtils.isNullOrEmpty(value)) return error;
        java.math.BigDecimal number = new java.math.BigDecimal(value.trim());
        if (minimum != null && number.compareTo(minimum) < 0)
            return getMessage("validation.numeric.minimum", fieldName, minimum.toPlainString());
        if (maximum != null && number.compareTo(maximum) > 0)
            return getMessage("validation.numeric.maximum", fieldName, maximum.toPlainString());
        return null;
    }

    /**
     * 【変換元】CPValidateUtility.vb - ValidateAuto()
     *   元コード: DataGridView全行のバリデーション
     */
    public <T> List<String> validateRows(List<T> rows,
                                         RowValidator<T> validator) {
        List<String> errors = new ArrayList<>();
        for (int i = 0; i < rows.size(); i++) {
            List<String> rowErrors = validator.validate(rows.get(i), i + 1);
            errors.addAll(rowErrors);
        }
        return errors;
    }

    /**
     * 行バリデーション用関数型インタフェース
     */
    @FunctionalInterface
    public interface RowValidator<T> {
        List<String> validate(T row, int rowIndex);
    }

    private String getMessage(String code, Object... args) {
        try {
            return messageSource.getMessage(code, args, Locale.JAPAN);
        } catch (Exception e) {
            return code + ": " + String.join(", ", java.util.Arrays.stream(args).map(Object::toString).toList());
        }
    }
}
