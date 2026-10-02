package com.daifuku.mcm.dto;

import jakarta.servlet.http.HttpSession;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

/** 発行単位のデータを保持する。別の発行で上書きせず、取得失敗後も再試行できる。 */
public final class Mcm1001pExportState {
    private static final String KEY = "mcm1001p.exports";
    private static final int MAX_BATCHES = 5;

    private Mcm1001pExportState() {}

    public static String remember(HttpSession session, List<Mcm1001pDeliveryDto> deliveries) {
        if (deliveries == null || deliveries.isEmpty()) {
            throw new IllegalArgumentException("依頼書の出力対象がありません。");
        }
        synchronized (session) {
            var states = states(session);
            String token = UUID.randomUUID().toString();
            states.put(token, List.copyOf(deliveries));
            while (states.size() > MAX_BATCHES) {
                states.remove(states.keySet().iterator().next());
            }
            session.setAttribute(KEY, states);
            return token;
        }
    }

    public static List<Mcm1001pDeliveryDto> find(HttpSession session, String token) {
        synchronized (session) {
            return token == null ? null : states(session).get(token);
        }
    }

    @SuppressWarnings("unchecked")
    private static Map<String, List<Mcm1001pDeliveryDto>> states(HttpSession session) {
        Object value = session.getAttribute(KEY);
        return value instanceof Map<?, ?>
                ? new LinkedHashMap<>((Map<String, List<Mcm1001pDeliveryDto>>) value)
                : new LinkedHashMap<>();
    }
}
