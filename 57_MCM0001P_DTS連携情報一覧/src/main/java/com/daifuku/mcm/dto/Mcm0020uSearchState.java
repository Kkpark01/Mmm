package com.daifuku.mcm.dto;

import com.daifuku.mcm.form.Mcm0020uForm;
import jakarta.servlet.http.HttpSession;
import java.util.List;
import java.util.Map;

/** MCM0020Uのタブ別検索結果。戻り表示とMCM0001P出力で同じ一覧を使用する。 */
public record Mcm0020uSearchState(Mcm0020uForm form,
                                List<DtslogSearchResultDto> results, boolean linkEnabled) {
    public static final String SESSION_KEY = "mcm0020u_returnStates";

    @SuppressWarnings("unchecked")
    public static Mcm0020uSearchState find(HttpSession session, String token) {
        if (token == null || token.isBlank()) return null;
        synchronized (session) {
            var states = (Map<String, Mcm0020uSearchState>) session.getAttribute(SESSION_KEY);
            return states == null ? null : states.get(token);
        }
    }
}
