package com.daifuku.mcm.common;

import java.io.Serializable;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.ResolverStyle;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;
import java.util.UUID;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpSession;
import org.springframework.beans.MutablePropertyValues;
import org.springframework.util.SerializationUtils;
import org.springframework.validation.DataBinder;

/** 店舗見積・ユーザ契約画面の入力、画面世代、受け渡しを扱う。 */
public final class CustomerScreenSupport {
    public static final String EXPIRED = "画面の情報が更新されています。画面を開き直してから操作してください。";
    public static final String DENIED = "権限がないため遷移できません。";
    private static final DateTimeFormatter DATE = DateTimeFormatter.ofPattern("uuuu/MM/dd")
        .withResolverStyle(ResolverStyle.STRICT);
    private CustomerScreenSupport() { }

    public static boolean same(BigDecimal a, BigDecimal b) { return a != null && b != null && a.compareTo(b) == 0; }
    public static LocalDate date(String value) {
        try { return LocalDate.parse(value, DATE); }
        catch (RuntimeException e) { throw new IllegalStateException("日付の書式を指定して下さい。(YYYY/MM/DD)"); }
    }
    public static String format(LocalDate value) { return value.format(DATE); }
    public static String token(HttpSession s, String screen) {
        String key = screen + ".workflowToken";
        if (s.getAttribute(key) == null) rotate(s, screen);
        return (String) s.getAttribute(key);
    }
    public static void rotate(HttpSession s, String screen) { s.setAttribute(screen + ".workflowToken", UUID.randomUUID().toString()); }
    public static void check(HttpSession s, String screen, HttpServletRequest request) {
        if (!token(s, screen).equals(request.getParameter("workflowToken"))) throw new IllegalStateException(EXPIRED);
    }
    public static <T extends Serializable> T copy(T value) { return SerializationUtils.clone(value); }

    /** 読み込んだフォームに、画面で編集可能な項目だけを反映する。IDや状態は受け付けない。 */
    public static void bind(HttpServletRequest request, Object target, String... fields) {
        DataBinder binder = new DataBinder(target);
        binder.setAllowedFields(fields);
        binder.setAutoGrowNestedPaths(false);
        binder.bind(new MutablePropertyValues(request.getParameterMap()));
        if (binder.getBindingResult().hasErrors()) throw new IllegalStateException("入力内容を確認してください。");
    }
    /** チェックボックスは、表示した行の番号だけを受け取る。 */
    public static Set<Integer> selections(HttpServletRequest request, String name, int size) {
        Set<Integer> result = new HashSet<>();
        String[] values = request.getParameterValues(name);
        if (values == null) return result;
        for (String value : values) {
            try {
                int i = Integer.parseInt(value);
                if (i < 0 || i >= size) throw new IllegalArgumentException();
                result.add(i);
            } catch (IllegalArgumentException e) { throw new IllegalStateException(EXPIRED); }
        }
        return result;
    }
    public static boolean update(HttpSession s,String screen,boolean role,java.util.function.Supplier<String> lookup) {
        if(!role || AppConstants.AUTHORITY_DIVISION_INSPECTION.equals(s.getAttribute(AppConstants.SESSION_AUTHORITY_DIVISION))
                || AppConstants.AUTHORITY_DIVISION_NONE.equals(s.getAttribute(AppConstants.SESSION_AUTHORITY_DIVISION)))return false;
        Object value=s.getAttribute("authority."+screen);
        if(value==null)try{value=lookup.get();}catch(org.springframework.dao.DataAccessException ex){return false;}
        return "2".equals(value) || "UPDATE".equals(value);
    }
    public static boolean update(HttpSession s, String screen, boolean role) {
        Object value = s.getAttribute("authority." + screen);
        if (value != null) return "2".equals(value.toString()) || "UPDATE".equals(value.toString());
        Object division = s.getAttribute(AppConstants.SESSION_AUTHORITY_DIVISION);
        if (AppConstants.AUTHORITY_DIVISION_INSPECTION.equals(division)) return false;
        return role;
    }
}
