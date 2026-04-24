// eslint-disable-next-line no-unused-vars
const Ladder = {
    refreshLadder(ladderSheet, activeBonds, today) {
        if (ladderSheet.getLastRow() > 1) {
            ladderSheet
                .getRange(
                    2,
                    1,
                    ladderSheet.getLastRow() - 1,
                    Config.LADDER_HEADERS.length,
                )
                .clearContent();
        }

        const buckets = [
            { label: '0-3 months', maxDays: 91, bonds: [] },
            { label: '3-6 months', maxDays: 182, bonds: [] },
            { label: '6-12 months', maxDays: 365, bonds: [] },
            { label: '1-2 years', maxDays: 730, bonds: [] },
            { label: '2-5 years', maxDays: 1825, bonds: [] },
            { label: '5+ years', maxDays: Infinity, bonds: [] },
        ];

        let totalFace = 0;

        activeBonds.forEach((bond) => {
            const daysToMaturity = Utils.daysBetween(today, bond.maturityDate);
            const faceTotal = bond.faceValue * bond.quantity;
            totalFace += faceTotal;

            for (const bucket of buckets) {
                if (daysToMaturity <= bucket.maxDays) {
                    bucket.bonds.push({ isin: bond.isin, face: faceTotal });
                    break;
                }
            }
        });

        const rows = buckets.map((b) => {
            const bucketFace = b.bonds.reduce((sum, x) => sum + x.face, 0);
            const pct =
                totalFace > 0
                    ? Utils.bankersRound((bucketFace / totalFace) * 10000) / 100
                    : 0;
            return [
                b.label,
                b.bonds.length,
                Utils.bankersRound(bucketFace * 100) / 100,
                pct,
                b.bonds.map((x) => x.isin).join(', '),
            ];
        });

        if (rows.length > 0) {
            ladderSheet
                .getRange(2, 1, rows.length, rows[0].length)
                .setValues(rows);
            for (let i = 0; i < rows.length; i++) {
                ladderSheet.getRange(i + 2, 3).setNumberFormat('#,##0.00');
                ladderSheet.getRange(i + 2, 4).setNumberFormat('0.00');
            }
        }
    },
};
