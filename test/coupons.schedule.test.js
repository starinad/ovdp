// Standalone, dependency-free regression checks for the coupon scheduling
// logic that drives every coupon, cashflow and analytics number.
// Run with: node --test test/
//
// Mirrors Coupons._buildBondCouponDates (src/coupons.js) — the file can't be
// imported directly (Apps Script globals), so the no-first-coupon-date branch
// is duplicated here; keep the two in sync when touching either.

const test = require('node:test');
const assert = require('node:assert/strict');

const Utils = {
    addMonthsSafe(date, months) {
        const result = new Date(date);
        const targetMonth = result.getMonth() + months;
        const targetDay = result.getDate();
        result.setMonth(targetMonth);
        if (result.getDate() !== targetDay) {
            result.setDate(0); // last day of previous month
        }
        return result;
    },

    daysBetween(a, b) {
        const msPerDay = 86400000;
        const utcA = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
        const utcB = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
        return Math.round((utcB - utcA) / msPerDay);
    },
};

function buildSchedule(maturityDate, firstCouponDate, monthsBetween) {
    if (!firstCouponDate) {
        const dates = [new Date(maturityDate)];
        let current = maturityDate;
        while (true) {
            current = Utils.addMonthsSafe(current, -monthsBetween);
            if (current.getTime() <= new Date(2000, 0, 1).getTime()) break;
            dates.unshift(new Date(current));
        }
        return dates;
    }

    const dates = [];
    let i = 0;
    while (true) {
        const date = Utils.addMonthsSafe(firstCouponDate, monthsBetween * i);
        if (date > maturityDate) break;
        dates.push(date);
        i++;
    }

    const lastDate = dates.length > 0 ? dates[dates.length - 1] : null;
    if (!lastDate) {
        dates.push(new Date(maturityDate));
    } else {
        const gapDays = Utils.daysBetween(lastDate, maturityDate);
        if (gapDays > 7) {
            dates.push(new Date(maturityDate));
        } else if (gapDays > 0) {
            dates[dates.length - 1] = new Date(maturityDate);
        }
    }
    return dates;
}

const iso = (d) => d.toISOString().slice(0, 10);

test('semi-annual schedule from first coupon keeps coupon date, adds maturity stub', () => {
    // Real bond from bonds.example.json: first coupon 17.11.2026, maturity 15.11.2028
    const dates = buildSchedule(
        new Date(2028, 10, 15), // Nov 15, 2028
        new Date(2026, 10, 17), // Nov 17, 2026
        6,
    ).map(iso);
    assert.deepEqual(dates, [
        '2026-11-16', // 17.11.2026 06:00 Europe/Kyiv = UTC 16th; Apps Script runs in Kyiv TZ
        '2027-05-16',
        '2027-11-16',
        '2028-05-16',
        '2028-11-14',
    ]);
    // User-visible dates (local Kyiv): first payment on the 17th, maturity on the 15th
    assert.equal(dates.length, 5);
});

test('no first coupon: backward walk from maturity yields clean 6-month grid', () => {
    const dates = buildSchedule(new Date(2027, 4, 15), null, 6).map(iso);
    assert.deepEqual(dates.slice(-5), [
        '2025-05-14',
        '2025-11-14',
        '2026-05-14',
        '2026-11-14',
        '2027-05-14',
    ]);
    assert.ok(dates.every((d) => d.endsWith('-15') || d.endsWith('-14')), 'all dates aligned');
});

test('quarterly schedule with clamped month-end stays aligned to maturity', () => {
    const dates = buildSchedule(new Date(2027, 5, 30), null, 3).map(iso);
    // Jun 30 -> Mar 30 -> Dec 30 ... (clamped to non-existent Mar 31 / Dec 31)
    assert.deepEqual(dates.slice(-3), ['2026-12-29', '2027-03-29', '2027-06-29']);
});
