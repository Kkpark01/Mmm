package com.daifuku.mcm.common;

/**
 * 【変換元】CPComponentUtility.vb (FRAMEWORK/UTILITY/CPComponentUtility.vb) - HandleException()
 *           CPBaseForm.vb - Try-Catch → MessageBox.Show パターン
 *
 * グローバル例外ハンドラー
 * VB.NETでは Try-Catch → MessageBox.Show で例外を処理していたが、
 * Web版では @ControllerAdvice でグローバルに例外をキャッチし、
 * エラーページまたはFlash Attributeでメッセージを返す。
 */

import com.daifuku.mcm.exception.*;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.servlet.ModelAndView;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import jakarta.servlet.http.HttpServletRequest;
import java.util.stream.Collectors;

@ControllerAdvice
public class GlobalExceptionHandler {

    private static final Logger logger = LoggerFactory.getLogger(GlobalExceptionHandler.class);

    @ExceptionHandler({McmBusinessException.class, ExclusiveControlException.class})
    public String handleBusiness(RuntimeException ex, HttpServletRequest request, RedirectAttributes ra) {
        logger.warn("[業務エラー] {}", ex.getMessage());
        ra.addFlashAttribute("error", ex.getMessage());
        if (request.getAttribute("mcmRecoveryUrl") != null) return recover(request, ra);
        String referer = request.getHeader("Referer");
        return "redirect:" + (referer != null ? referer : "/menu");
    }

    private String recover(HttpServletRequest request, RedirectAttributes ra) {
        Object values = request.getAttribute("mcmSubmittedValues");
        if (values != null) ra.addFlashAttribute("submittedValues", values);
        return "redirect:" + request.getAttribute("mcmRecoveryUrl");
    }

    /**
     * 【変換元】CPComponentUtility.vb - HandleException (InputCheckException)
     *   元コード: MessageBox.Show(ex.ErrorMessages, "入力エラー", Warning)
     */
    @ExceptionHandler(InputCheckException.class)
    public String handleInputCheck(InputCheckException ex,
                                   HttpServletRequest request,
                                   RedirectAttributes ra) {
        logger.warn("[入力エラー] {}", ex.getErrors());
        ra.addFlashAttribute("errors", ex.getErrors());
        if (request.getAttribute("mcmRecoveryUrl") != null) return recover(request, ra);
        String referer = request.getHeader("Referer");
        return "redirect:" + (referer != null ? referer : "/menu");
    }

    /**
     * 【変換元】CPComponentUtility.vb - HandleException (DataNotChangedException)
     *   元コード: MessageBox.Show("データに変更がありません", "情報", Information)
     */
    @ExceptionHandler(DataNotChangedException.class)
    public String handleNoChange(DataNotChangedException ex,
                                 HttpServletRequest request,
                                 RedirectAttributes ra) {
        logger.info("[未変更] {}", ex.getMessage());
        ra.addFlashAttribute("info", ex.getMessage());
        String referer = request.getHeader("Referer");
        return "redirect:" + (referer != null ? referer : "/menu");
    }

    /**
     * 【変換元】CPComponentUtility.vb - HandleException (AuthorityException)
     *   元コード: MessageBox.Show("権限がありません", "権限エラー", Error)
     */
    @ExceptionHandler(AuthorityException.class)
    public String handleAuthority(AuthorityException ex, RedirectAttributes ra) {
        logger.warn("[権限エラー] {}", ex.getMessage());
        ra.addFlashAttribute("error", ex.getMessage());
        return "redirect:/menu";
    }

    /**
     * 【変換元】CPComponentUtility.vb - HandleException (ProcessCanceledException)
     *   元コード: (処理キャンセル - ログのみ)
     */
    @ExceptionHandler(ProcessCanceledException.class)
    public String handleCancel(ProcessCanceledException ex,
                               HttpServletRequest request) {
        logger.info("[処理キャンセル] {}", ex.getMessage());
        String referer = request.getHeader("Referer");
        return "redirect:" + (referer != null ? referer : "/menu");
    }

    /**
     * 【変換元】CPComponentUtility.vb - HandleException (SpecifiedException)
     *   元コード: MessageBox.Show(ex.Message, "エラー", Error)
     */
    @ExceptionHandler(SpecifiedException.class)
    public String handleSpecified(SpecifiedException ex,
                                  HttpServletRequest request,
                                  RedirectAttributes ra) {
        logger.error("[業務エラー] {}", ex.getMessage());
        ra.addFlashAttribute("error", ex.getMessage());
        String referer = request.getHeader("Referer");
        return "redirect:" + (referer != null ? referer : "/menu");
    }

    /**
     * 【変換元】CPComponentUtility.vb - HandleException (Exception - 汎用)
     *   元コード: MessageBox.Show("システムエラーが発生しました", "システムエラー", Error)
     */
    @ExceptionHandler(Exception.class)
    public ModelAndView handleGeneral(Exception ex, HttpServletRequest request) {
        logger.error("[システムエラー] URL: {} - {}", request.getRequestURI(), ex.getMessage(), ex);
        if (request.getAttribute("mcmRecoveryUrl") != null) {
            org.springframework.web.servlet.FlashMap flash = org.springframework.web.servlet.support.RequestContextUtils.getOutputFlashMap(request);
            flash.put("error", "登録処理に失敗しました。入力内容は保持されています。管理者に連絡してください。");
            flash.put("submittedValues", request.getAttribute("mcmSubmittedValues"));
            return new ModelAndView("redirect:" + request.getAttribute("mcmRecoveryUrl"));
        }
        ModelAndView mav = new ModelAndView("error/500");
        mav.addObject("errorMessage", "システムエラーが発生しました。管理者に連絡してください。");
        mav.addObject("errorDetail", ex.getMessage());
        return mav;
    }
}
