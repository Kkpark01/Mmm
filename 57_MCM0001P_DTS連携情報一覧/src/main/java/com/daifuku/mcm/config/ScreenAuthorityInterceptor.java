package com.daifuku.mcm.config;

import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import jakarta.servlet.http.HttpSession;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.servlet.HandlerInterceptor;

import com.daifuku.mcm.common.AppConstants;
import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.common.MessageService;
import com.daifuku.mcm.common.ScreenAuthorityRegistry;
import com.daifuku.mcm.exception.AuthorityException;

/**
 * 画面アクセス時のサーバー側権限チェックを一元的に行う共通Interceptor（FWM_0012対応）。
 *
 * 【不具合内容】
 *   メニュー画面（MCM0001U）では権限のない機能ボタンを非表示にしているが、
 *   これは表示制御のみであり、サーバー側の権限チェックが実装されていない
 *   画面が大半だったため、以下の経路で権限外画面へアクセスできてしまっていた。
 *     - URL直接入力
 *     - 他画面からの不正な画面遷移
 *     - 複数ブラウザ／複数セッション利用時（別ユーザーとして正規ログイン後の直URL入力）
 *
 * 【対応方針】
 *   画面表示制御（{@code MenuService}）とサーバー側アクセス制御（本クラス）の
 *   両方が {@link ScreenAuthorityRegistry} の同じ一覧を参照するようにし、
 *   URL直接入力・フォワード・リダイレクトのいずれの経路でも
 *   Controller に到達する前に必ず権限チェックが行われるようにする。
 *   各 Controller 個別に {@code @PreAuthorize} や
 *   {@code BaseController#requireUpdateAuthority()} を呼び出す実装は、
 *   呼び出し漏れが起きるとそのまま脆弱性になるため、
 *   共通Interceptorへ処理を集約する。
 *
 * 【権限不足時の挙動】
 *   {@link AuthorityException} をスローし、{@code GlobalExceptionHandler} で
 *   捕捉して遷移元（メニュー画面）へリダイレクトする。
 *   メッセージは FWM_0012（「権限がないため遷移できません。」）を使用する。
 */
public class ScreenAuthorityInterceptor implements HandlerInterceptor {

    private static final Logger logger = LoggerFactory.getLogger(ScreenAuthorityInterceptor.class);

    /** URLの先頭パスセグメントから画面IDを取り出す正規表現（例: /mcm0021u/search → mcm0021u） */
    private static final Pattern SCREEN_PATH_PATTERN = Pattern.compile("^/([a-zA-Z0-9]+)(?:/.*)?$");

    private final MessageService messageService;

    public ScreenAuthorityInterceptor(MessageService messageService) {
        this.messageService = messageService;
    }

    @Override
    public boolean preHandle(HttpServletRequest request, HttpServletResponse response, Object handler) {

        String screenId = resolveScreenId(request);
        if (screenId == null) {
            return true;
        }

        // MCM0001PはMCM0020Uの一覧を出力する帳票。
        // 帳票URLへの直接アクセスでも、起動元画面の権限制限を適用する。
        String authorityScreenId = "MCM0001P".equals(screenId) ? "MCM0020U" : screenId;

        HttpSession session = request.getSession(false);
        if (session == null) {
            // 未認証はSpring Securityが先に処理する（ここには到達しない想定）。
            return true;
        }

        String authorityDivision = (String) session.getAttribute(AppConstants.SESSION_AUTHORITY_DIVISION);
        String systemManagementKengen = (String) session.getAttribute(
                AppConstants.SESSION_SYSTEM_MANAGEMENT_KENGEN);
        @SuppressWarnings("unchecked")
        java.util.Set<String> deniedFunctionIds = (java.util.Set<String>) session.getAttribute(
                AppConstants.SESSION_DENIED_FUNCTION_IDS);

        boolean isInspectionOnly = AppConstants.AUTHORITY_DIVISION_INSPECTION.equals(authorityDivision);
        boolean noSystemManagementAuthority = (systemManagementKengen == null
                || McmConstants.RIYOKENGEN_KBN_NASI_CD.equals(systemManagementKengen));

        boolean denied = (isInspectionOnly && ScreenAuthorityRegistry.RESTRICTED_FOR_INSPECTION.contains(authorityScreenId))
                || (noSystemManagementAuthority
                        && ScreenAuthorityRegistry.RESTRICTED_FOR_NO_SYSTEM_MANAGEMENT.contains(authorityScreenId))
                // 【追加】メニュー表示制御の統一対応：MCM_MO_KENGENKOSEIで権限管理対象の
                //   全画面（KENGENBUNRUI_ID=1〜7を含む）について、実効権限が「なし」の
                //   機能IDへのURL直打ちアクセスを拒否する。
                || (deniedFunctionIds != null && (deniedFunctionIds.contains(screenId)
                        || deniedFunctionIds.contains(authorityScreenId)));

        if (denied) {
            String userId = (String) session.getAttribute(AppConstants.SESSION_USER_ID);
            logger.warn("[画面アクセス拒否] screenId={}, userId={}, uri={}", screenId, userId,
                    request.getRequestURI());
            throw new AuthorityException(messageService.getMessage(AppConstants.FWM_UNAUTHORIZED_SCREEN_ACCESS));
        }

        return true;
    }

    /**
     * リクエストパスの先頭セグメントから画面ID（例: MCM0021U）を解決する。
     * 画面IDの命名規則（mcmXXXXu 等）に合致しないパスは null を返す。
     */
    private String resolveScreenId(HttpServletRequest request) {
        String path = request.getRequestURI();
        String contextPath = request.getContextPath();
        if (contextPath != null && !contextPath.isEmpty() && path.startsWith(contextPath)) {
            path = path.substring(contextPath.length());
        }
        Matcher matcher = SCREEN_PATH_PATTERN.matcher(path);
        if (!matcher.matches()) {
            return null;
        }
        return matcher.group(1).toUpperCase(Locale.ROOT);
    }
}
