package com.daifuku.mcm.service;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.List;

import org.junit.jupiter.api.Test;

import com.daifuku.mcm.repository.Mcm3002uRepository.RatePeriod;

/**
 * 【変換元】database/30_func/mcm_fn_siharai.sql
 * Mcm3002uService#calculatePayment の月配分・端数・初回繰越ロジックを検証する。
 *
 * 優先度2対応：VB版（MCM_FN_SIHARAI）と同じ計算結果になることを保証する回帰テスト。
 */
class Mcm3002uServiceTest {

    /** 契約期間ちょうど12ヶ月、割り切れる金額（端数なし）。単純な月配分の確認。 */
    @Test
    void evenlyDivisibleAmountIsSplitPerMonth() {
        var start = LocalDate.of(2026, 1, 1);
        var end = LocalDate.of(2026, 12, 31);
        var rates = List.of(new RatePeriod(start, end, new BigDecimal("1200000")));
        // 2月も支払月にして1月分だけを集計する。次の支払月までは対象月に含まれる。
        var payments = List.of(LocalDate.of(2026, 1, 1), LocalDate.of(2026, 2, 1));

        var amount = Mcm3002uService.calculatePayment(start, end, YearMonth.of(2026, 1), rates, payments);

        // 1200000 / 12ヶ月 = 100000（端数なし）。次の支払月の直前までの1ヶ月分。
        assertThat(amount).isEqualByComparingTo("100000");
    }

    /** 端数が発生する場合、SYURYO_DTがある契約は開始月に端数を上乗せする。 */
    @Test
    void remainderIsAddedToFirstMonthWhenEndDateExists() {
        var start = LocalDate.of(2026, 1, 1);
        var end = LocalDate.of(2026, 3, 31); // 3ヶ月
        var rates = List.of(new RatePeriod(start, end, new BigDecimal("1000"))); // 1000/3=333余り1
        var payments = List.of(LocalDate.of(2026, 1, 1), LocalDate.of(2026, 2, 1));

        var amount = Mcm3002uService.calculatePayment(start, end, YearMonth.of(2026, 1), rates, payments);

        // 単位金額333 + 端数1 = 334（初月に乗せる）
        assertThat(amount).isEqualByComparingTo("334");
    }

    /** SYURYO_DTがない（無期限）契約は12ヶ月周期で端数を繰越加算する。 */
    @Test
    void remainderRecursEveryTwelveMonthsWhenEndDateIsNull() {
        var start = LocalDate.of(2026, 1, 1);
        var end = LocalDate.of(2027, 12, 31); // 24ヶ月分を対象期間として観測
        var rates = List.of(new RatePeriod(start, null, new BigDecimal("1000"))); // 終了日なし：12ヶ月換算 1000/12=83余り4
        var payments = List.of(LocalDate.of(2026, 1, 1), LocalDate.of(2027, 1, 1), LocalDate.of(2027, 2, 1));

        var amount = Mcm3002uService.calculatePayment(start, end, YearMonth.of(2027, 1), rates, payments);

        // 2年目の1月は12ヶ月周期の境目（開始から12ヶ月後）なので端数が再度乗る。
        // 2027/1は2回目の支払月なので過去分をリセットし、次の支払月(2月)の直前までを集計する。
        assertThat(amount).isEqualByComparingTo("87"); // 83 + 4
    }

    /** 2回目以降の支払は過去分をリセットし、対象月から次の支払月直前（または契約末尾）まで。 */
    @Test
    void laterPaymentAccumulatesFromTargetToContractEnd() {
        var start = LocalDate.of(2026, 1, 1);
        var end = LocalDate.of(2026, 6, 30); // 6ヶ月
        var rates = List.of(new RatePeriod(start, end, new BigDecimal("600"))); // 600/6=100（端数なし）
        // 1月と4月に支払タイミングがある
        var payments = List.of(LocalDate.of(2026, 1, 1), LocalDate.of(2026, 4, 1));

        var amount = Mcm3002uService.calculatePayment(start, end, YearMonth.of(2026, 4), rates, payments);

        // 4月支払対象分は、4月から契約末尾(6月)までの3ヶ月分=300
        assertThat(amount).isEqualByComparingTo("300");
    }

    /** 対象月が支払タイミングに含まれない場合は0円（未検出）。 */
    @Test
    void returnsZeroWhenTargetMonthIsNotAPaymentMonth() {
        var start = LocalDate.of(2026, 1, 1);
        var end = LocalDate.of(2026, 12, 31);
        var rates = List.of(new RatePeriod(start, end, new BigDecimal("1200000")));
        var payments = List.of(LocalDate.of(2026, 1, 1));

        // 対象月(2月)は支払タイミングに含まれない
        var amount = Mcm3002uService.calculatePayment(start, end, YearMonth.of(2026, 2), rates, payments);

        assertThat(amount).isEqualByComparingTo("0");
    }

    /** 見積金額が0以下の単価情報は無視する（VB: GOKEI_KIN > 0 の条件）。 */
    @Test
    void nonPositiveRateAmountIsIgnored() {
        var start = LocalDate.of(2026, 1, 1);
        var end = LocalDate.of(2026, 12, 31);
        var rates = List.of(
                new RatePeriod(start, end, BigDecimal.ZERO),
                new RatePeriod(start, end, new BigDecimal("1200000")));
        var payments = List.of(LocalDate.of(2026, 1, 1), LocalDate.of(2026, 2, 1));

        var amount = Mcm3002uService.calculatePayment(start, end, YearMonth.of(2026, 1), rates, payments);

        assertThat(amount).isEqualByComparingTo("100000");
    }

    /** 契約期間が不正（終了日が開始日より前）の場合は例外。 */
    @Test
    void invalidContractPeriodThrows() {
        var start = LocalDate.of(2026, 3, 1);
        var end = LocalDate.of(2026, 1, 1);

        assertThatThrownBy(() -> Mcm3002uService.calculatePayment(start, end, YearMonth.of(2026, 1),
                List.of(), List.of()))
                .isInstanceOf(IllegalArgumentException.class);
    }

    /** parseYearMonth: 正常な YYYYMM 文字列を月初日に変換する。 */
    @Test
    void parseYearMonthConvertsToFirstDayOfMonth() {
        var service = new Mcm3002uService();
        assertThat(service.parseYearMonth("202608")).isEqualTo(LocalDate.of(2026, 8, 1));
    }

    /** parseYearMonth: 不正な形式は例外（VB版 FWM_0006 相当）。 */
    @Test
    void parseYearMonthRejectsInvalidFormat() {
        var service = new Mcm3002uService();
        assertThatThrownBy(() -> service.parseYearMonth("2026-08"))
                .isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> service.parseYearMonth(""))
                .isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> service.parseYearMonth(null))
                .isInstanceOf(IllegalArgumentException.class);
    }

    /** VB関数は対象月から次の支払月までを集計する。初月だけ年払いなら年間全額となる。 */
    @Test
    void oneAnnualPaymentIncludesAllMonthsOfTheContract() {
        var start = LocalDate.of(2026, 1, 1);
        var end = LocalDate.of(2026, 12, 31);
        var amount = Mcm3002uService.calculatePayment(start, end, YearMonth.of(2026, 1),
                List.of(new RatePeriod(start, end, new BigDecimal("1200000"))), List.of(start));
        assertThat(amount).isEqualByComparingTo("1200000");
    }

    /** 最初の支払月だけは開始月からの未払い分を含め、次の支払月の直前で区切る。 */
    @Test
    void firstPaymentKeepsEarlierMonthsAndStopsBeforeNextPayment() {
        var start = LocalDate.of(2026, 1, 1);
        var end = LocalDate.of(2026, 8, 31);
        var rates = List.of(new RatePeriod(start, end, new BigDecimal("800")));
        var payments = List.of(LocalDate.of(2026, 3, 1), LocalDate.of(2026, 5, 1));
        assertThat(Mcm3002uService.calculatePayment(start, end, YearMonth.of(2026, 3), rates, payments))
                .isEqualByComparingTo("400"); // 1～4月。次の5月を含めない。
        assertThat(Mcm3002uService.calculatePayment(start, end, YearMonth.of(2026, 5), rates, payments))
                .isEqualByComparingTo("400"); // 5～8月。過去の1～4月を含めない。
    }
}
