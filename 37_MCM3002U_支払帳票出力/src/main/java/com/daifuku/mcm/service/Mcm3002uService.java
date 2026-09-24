package com.daifuku.mcm.service;

import com.daifuku.mcm.form.Mcm3002uForm;
import com.daifuku.mcm.form.Mcm3002uForm.RowForm;
import com.daifuku.mcm.repository.Mcm3002uRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.List;

/**
 * 【変換元】Mcm3002uScreen.vb OutputButton_Click
 * MCM3002U 支払い明細出力指示サービス
 *
 * MCM_FN_SIHARAIの月配分と支払タイミングの集計をJavaへ移植。
 */
@Service
public class Mcm3002uService {

    @Autowired
    private Mcm3002uRepository repository;

    private static final DateTimeFormatter YYYYMM_FMT = DateTimeFormatter.ofPattern("yyyyMM");

    // ===================================================================
    // 検索（VB版: Fill + 0件チェック に相当）
    // 【変換元】Mcm3002uScreen.vb OutputButton_Click 内の Fill 処理
    //   YYYYMM 文字列を月初日 LocalDate に変換して repository に渡す。
    //   0件は呼び出し側（Controller）でメッセージ表示。
    // ===================================================================

    @Transactional(readOnly = true)
    public List<RowForm> search(Mcm3002uForm form) {
        LocalDate shiharaiDate = parseYearMonth(form.getShiharaiTsuki());
        if (!"年払".equals(form.getSiharaiKbn()) && !"月払".equals(form.getSiharaiKbn())) {
            throw new IllegalArgumentException("支払区分を選択してください。");
        }
        return repository.search(shiharaiDate, form.getSiharaiKbn());
    }

    @Transactional(readOnly = true)
    public List<RowForm> searchForOutput(Mcm3002uForm form) {
        List<RowForm> rows = search(form);
        var target = YearMonth.from(parseYearMonth(form.getShiharaiTsuki()));
        // 同一契約期間が複数行に出る場合も、金額取得は一度だけ。
        record Key(java.math.BigDecimal id, LocalDate start, LocalDate end) {}
        var amounts = new java.util.HashMap<Key, java.math.BigDecimal>();
        for (RowForm row : rows) {
            if (row.getTkKikanId() == null || row.getKaishiDt() == null || row.getSyuryoDt() == null) {
                throw new IllegalArgumentException("契約期間が取得できないため帳票を出力できません。");
            }
            var key = new Key(row.getTkKikanId().stripTrailingZeros(), row.getKaishiDt(), row.getSyuryoDt());
            var amount = amounts.computeIfAbsent(key, k -> calculatePayment(k.start(), k.end(), target,
                    repository.findRatePeriods(k.id()),
                    repository.findPaymentDates(k.id(), k.start().withDayOfMonth(1), k.end().plusDays(1).withDayOfMonth(1))));
            row.setShiharaiKingaku(amount);
        }
        return rows;
    }

    /** database/30_func/mcm_fn_siharai.sqlの配分・端数・初回繰越をそのまま移植。 */
    public static java.math.BigDecimal calculatePayment(LocalDate start, LocalDate end, YearMonth target,
            List<Mcm3002uRepository.RatePeriod> rates, List<LocalDate> payments) {
        if (start == null || end == null || end.isBefore(start)) {
            throw new IllegalArgumentException("契約期間が不正です。");
        }
        var first = YearMonth.from(start);
        var limit = YearMonth.from(end.plusDays(1));
        var monthly = new java.util.TreeMap<YearMonth, java.math.BigDecimal>();
        for (var m = first; m.isBefore(limit); m = m.plusMonths(1)) monthly.put(m, java.math.BigDecimal.ZERO);
        for (var rate : rates) {
            if (rate.start() == null || rate.amount() == null) throw new IllegalArgumentException("見積期間が不正です。");
            if (rate.amount().signum() <= 0) continue;
            var rs = YearMonth.from(rate.start());
            var re = rate.end() == null ? rs.plusMonths(12) : YearMonth.from(rate.end().plusDays(1));
            long months = java.time.temporal.ChronoUnit.MONTHS.between(rs, re);
            if (months < 0) throw new IllegalArgumentException("見積期間が不正です。");
            var remainder = months == 0 ? java.math.BigDecimal.ZERO : rate.amount().remainder(java.math.BigDecimal.valueOf(months));
            var unit = months == 0 ? rate.amount() : rate.amount().subtract(remainder).divide(java.math.BigDecimal.valueOf(months));
            var stop = rate.end() == null || re.isAfter(limit) ? limit : re;
            for (var m = rs.isBefore(first) ? first : rs; m.isBefore(stop); m = m.plusMonths(1)) {
                boolean extra = rate.end() == null
                        ? java.time.temporal.ChronoUnit.MONTHS.between(rs, m) % 12 == 0 : m.equals(rs);
                monthly.put(m, monthly.get(m).add(unit).add(extra ? remainder : java.math.BigDecimal.ZERO));
            }
        }
        var paid = new java.util.HashSet<YearMonth>();
        for (var date : payments) if (!date.isBefore(first.atDay(1)) && date.isBefore(limit.atDay(1))) paid.add(YearMonth.from(date));
        var result = java.math.BigDecimal.ZERO;
        boolean initial = true, found = false;
        for (var entry : monthly.entrySet()) {
            if (paid.contains(entry.getKey())) {
                if (found) break;
                if (initial) initial = false; else result = java.math.BigDecimal.ZERO;
                if (entry.getKey().equals(target)) found = true;
            }
            result = result.add(entry.getValue());
        }
        return found ? result : java.math.BigDecimal.ZERO;
    }

    // ===================================================================
    // YYYYMM文字列 → 月初日 LocalDate
    // 【変換元】VB版: TRUNC(TKV.TSUKI, 'DD') = :{0} のパラメーター生成
    //   入力: "202608" → LocalDate: 2026-08-01
    // ===================================================================

    public LocalDate parseYearMonth(String yyyyMm) {
        if (yyyyMm == null || yyyyMm.isBlank()) {
            throw new IllegalArgumentException("支払い月を入力してください。");
        }
        try {
            if (!yyyyMm.trim().matches("[0-9]{6}")) throw new DateTimeParseException("YYYYMM", yyyyMm, 0);
            YearMonth ym = YearMonth.parse(yyyyMm.trim(), YYYYMM_FMT);
            return ym.atDay(1);
        } catch (DateTimeParseException e) {
            throw new IllegalArgumentException("支払い月は YYYYMM 形式で入力してください（例: 202608）。");
        }
    }
}
