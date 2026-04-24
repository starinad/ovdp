// eslint-disable-next-line no-unused-vars
const Utils = {
    addMonthsSafe(date, months) {
        const result = new Date(date);
        const targetMonth = result.getMonth() + months;
        const targetDay = result.getDate();
        result.setMonth(targetMonth);
        // Handle overflow (e.g., Jan 31 + 1 month should be Feb 28, not Mar 3)
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

    yearFraction(periodStart, periodEnd, convention) {
        switch (convention) {
            case 'ACT/365': {
                const days = this.daysBetween(periodStart, periodEnd);
                return { days, divisor: 365 };
            }

            case 'ACT/ACT': {
                const days = this.daysBetween(periodStart, periodEnd);
                const startYear = periodStart.getFullYear();
                const endYear = periodEnd.getFullYear();

                if (startYear === endYear) {
                    return {
                        days,
                        divisor: this.isLeapYear(startYear) ? 366 : 365,
                    };
                }

                const yearBoundary = new Date(endYear, 0, 1);
                const daysInStart = this.daysBetween(periodStart, yearBoundary);
                const daysInEnd = this.daysBetween(yearBoundary, periodEnd);
                const startDenom = this.isLeapYear(startYear) ? 366 : 365;
                const endDenom = this.isLeapYear(endYear) ? 366 : 365;
                const weightedDivisor = Math.round(
                    days / (daysInStart / startDenom + daysInEnd / endDenom),
                );
                return { days, divisor: weightedDivisor };
            }

            case '30/360': {
                let d1 = periodStart.getDate();
                let d2 = periodEnd.getDate();
                const m1 = periodStart.getMonth() + 1;
                const m2 = periodEnd.getMonth() + 1;
                const y1 = periodStart.getFullYear();
                const y2 = periodEnd.getFullYear();

                if (d1 === 31) d1 = 30;
                if (d2 === 31 && d1 >= 30) d2 = 30;

                const days = (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
                return { days, divisor: 360 };
            }

            default:
                return this.yearFraction(periodStart, periodEnd, 'ACT/365');
        }
    },

    isLeapYear(year) {
        return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
    },

    bankersRound(value) {
        const floor = Math.floor(value);
        const decimal = value - floor;
        if (Math.abs(decimal - 0.5) < 1e-10) {
            return floor % 2 === 0 ? floor : floor + 1;
        }
        return Math.round(value);
    },

    normalizeDate(d) {
        if (!d) return null;
        const date = new Date(d);
        date.setHours(0, 0, 0, 0);
        return date;
    },

    formatMonth(date) {
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        return `${y}-${m}`;
    },

    formatUAH(amount) {
        return (
            amount.toLocaleString('uk-UA', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2,
            }) + ' UAH'
        );
    },

    getColumnIndex(headers, name) {
        const idx = headers.findIndex((h) => h.header === name);

        if (idx === -1) {
            throw new Error(`Column "${name}" not found`);
        }

        return idx + 1;
    },
};
