/**
 * Stats.js - Client-side statistics computation library
 * Depends on jStat (loaded via CDN before this file)
 */
(function () {
    "use strict";

    // =========================================================================
    // Helpers
    // =========================================================================

    function isValid(arr) {
        return Array.isArray(arr) && arr.length > 0;
    }

    function clean(arr) {
        if (!Array.isArray(arr)) return [];
        return arr.filter(function (v) { return v !== null && v !== undefined && !isNaN(v) && isFinite(v); }).map(Number);
    }

    function sum(a) { var s = 0; for (var i = 0; i < a.length; i++) s += a[i]; return s; }
    function mean(a) { return a.length === 0 ? NaN : sum(a) / a.length; }
    function variance(a, m) {
        if (a.length < 2) return NaN;
        if (m === undefined) m = mean(a);
        var s = 0; for (var i = 0; i < a.length; i++) s += (a[i] - m) * (a[i] - m);
        return s / (a.length - 1);
    }
    function sd(a, m) { return Math.sqrt(variance(a, m)); }
    function se(a) { return sd(a) / Math.sqrt(a.length); }
    function median(a) {
        var s = a.slice().sort(function (x, y) { return x - y; });
        var mid = Math.floor(s.length / 2);
        return s.length % 2 !== 0 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
    }
    function percentile(a, p) {
        var s = a.slice().sort(function (x, y) { return x - y; });
        var idx = (p / 100) * (s.length - 1);
        var lo = Math.floor(idx), hi = Math.ceil(idx);
        if (lo === hi) return s[lo];
        return s[lo] + (s[hi] - s[lo]) * (idx - lo);
    }
    function mode(a) {
        var freq = {}, maxF = 0, modes = [];
        for (var i = 0; i < a.length; i++) {
            var v = a[i]; freq[v] = (freq[v] || 0) + 1;
            if (freq[v] > maxF) maxF = freq[v];
        }
        for (var k in freq) if (freq[k] === maxF) modes.push(Number(k));
        return modes.length === a.length ? NaN : (modes.length === 1 ? modes[0] : modes);
    }

    function ranks(values) {
        var n = values.length;
        var indexed = values.map(function (v, i) { return { v: v, i: i }; });
        indexed.sort(function (a, b) { return a.v - b.v; });
        var r = new Array(n);
        var i = 0;
        while (i < n) {
            var j = i;
            while (j < n - 1 && indexed[j + 1].v === indexed[i].v) j++;
            var avgRank = (i + j) / 2 + 1;
            for (var k = i; k <= j; k++) r[indexed[k].i] = avgRank;
            i = j + 1;
        }
        return r;
    }

    function tCdf(t, df) {
        return jStat.studentt.cdf(t, df);
    }
    function tInv(p, df) {
        return jStat.studentt.inv(p, df);
    }
    function fCdf(f, df1, df2) {
        return jStat.centralF.cdf(f, df1, df2);
    }
    function chi2Cdf(x, df) {
        return jStat.chisquare.cdf(x, df);
    }
    function normCdf(z) {
        return jStat.normal.cdf(z, 0, 1);
    }

    function formatPValue(p) {
        if (p === null || p === undefined || isNaN(p)) return "N/A";
        if (p < 0.001) return "< .001";
        return p.toFixed(3).replace(/^0/, "");
    }

    function significanceStar(p) {
        if (p === null || p === undefined || isNaN(p)) return "";
        if (p < 0.001) return "***";
        if (p < 0.01) return "**";
        if (p < 0.05) return "*";
        return "";
    }

    function interpretCohensD(d) {
        d = Math.abs(d);
        if (d < 0.2) return "Negligible";
        if (d < 0.5) return "Small";
        if (d < 0.8) return "Medium";
        if (d < 1.2) return "Large";
        return "Very Large";
    }

    function interpretEtaSquared(eta2) {
        if (eta2 < 0.06) return "Small";
        if (eta2 < 0.14) return "Medium";
        return "Large";
    }

    function interpretR(r) {
        var a = Math.abs(r);
        if (a < 0.1) return "Negligible";
        if (a < 0.3) return "Weak";
        if (a < 0.5) return "Moderate";
        if (a < 0.7) return "Strong";
        return "Very Strong";
    }

    function interpretCramersV(v) {
        if (v < 0.1) return "Negligible";
        if (v < 0.3) return "Small";
        if (v < 0.5) return "Medium";
        return "Large";
    }

    function interpretAlpha(a) {
        if (a >= 0.9) return "Excellent";
        if (a >= 0.8) return "Good";
        if (a >= 0.7) return "Acceptable";
        if (a >= 0.6) return "Questionable";
        if (a >= 0.5) return "Poor";
        return "Unacceptable";
    }

    // =========================================================================
    // Descriptive Statistics
    // =========================================================================

    function descriptive(values) {
        var v = clean(values);
        if (v.length === 0) return null;
        var n = v.length;
        var m = mean(v);
        var s = sd(v, m);
        var s2 = variance(v, m);
        var sem = s / Math.sqrt(n);
        var sorted = v.slice().sort(function (a, b) { return a - b; });
        var mn = sorted[0], mx = sorted[sorted.length - 1];

        // skewness (sample)
        var skew = NaN, kurt = NaN;
        if (n >= 3) {
            var m3 = 0, m4 = 0;
            for (var i = 0; i < n; i++) {
                var d = v[i] - m;
                m3 += d * d * d;
                m4 += d * d * d * d;
            }
            var s3 = s * s * s;
            var s4 = s * s * s * s;
            skew = (n / ((n - 1) * (n - 2))) * (m3 / s3) * (n > 0 ? 1 : 0);
            if (n >= 4) {
                kurt = ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * (m4 / s4) -
                       (3 * (n - 1) * (n - 1)) / ((n - 2) * (n - 3));
            }
        }

        var tCrit = n > 1 ? tInv(0.975, n - 1) : NaN;

        return {
            n: n, mean: m, se: sem, sd: s, variance: s2,
            min: mn, max: mx, range: mx - mn,
            median: median(v), mode: mode(v),
            skewness: skew, kurtosis: kurt,
            p25: percentile(v, 25), p50: percentile(v, 50), p75: percentile(v, 75),
            ci95_lo: m - tCrit * sem, ci95_hi: m + tCrit * sem,
            ci95: formatCI(m - tCrit * sem, m + tCrit * sem)
        };
    }

    // =========================================================================
    // CI formatting helper
    // =========================================================================

    function formatCI(lo, hi) {
        if (lo == null || hi == null || isNaN(lo) || isNaN(hi)) return '';
        return '[' + lo.toFixed(2) + ', ' + hi.toFixed(2) + ']';
    }

    // =========================================================================
    // Normality Tests
    // =========================================================================

    function ksTest(values) {
        var v = clean(values);
        if (v.length < 5) return null;
        var n = v.length;
        var m = mean(v), s = sd(v);
        var sorted = v.slice().sort(function (a, b) { return a - b; });
        var D = 0;
        for (var i = 0; i < n; i++) {
            var Fn = (i + 1) / n;
            var Fn1 = i / n;
            var Fx = normCdf((sorted[i] - m) / s);
            D = Math.max(D, Math.abs(Fn - Fx), Math.abs(Fn1 - Fx));
        }
        // Kolmogorov-Smirnov p-value approximation (Marsaglia 2003 / simplified)
        var sqn = Math.sqrt(n);
        var lambda = (sqn + 0.12 + 0.11 / sqn) * D;
        var p = 0;
        for (var k = 1; k <= 100; k++) {
            p += Math.pow(-1, k - 1) * Math.exp(-2 * k * k * lambda * lambda);
        }
        p = 2 * p;
        p = Math.max(0, Math.min(1, p));
        return { D: D, p: p };
    }

    function shapiroWilk(values) {
        var v = clean(values);
        if (v.length < 3 || v.length > 5000) return null;
        var n = v.length;

        // For small samples use the Shapiro-Wilk algorithm
        // For larger samples fall back to KS test
        if (n > 50) {
            var ks = ksTest(v);
            return ks ? { W: 1 - ks.D, p: ks.p } : null;
        }

        var sorted = v.slice().sort(function (a, b) { return a - b; });
        var m = mean(sorted);
        var ss = 0;
        for (var i = 0; i < n; i++) ss += (sorted[i] - m) * (sorted[i] - m);

        // Approximate expected normal order statistics using Blom formula
        var mi = [];
        for (var i = 0; i < n; i++) {
            var p_i = (i + 1 - 0.375) / (n + 0.25);
            mi.push(jStat.normal.inv(p_i, 0, 1));
        }

        // a coefficients: mi / ||mi||
        var mNorm = 0;
        for (var i = 0; i < n; i++) mNorm += mi[i] * mi[i];
        mNorm = Math.sqrt(mNorm);
        var a = mi.map(function (x) { return x / mNorm; });

        var aX = 0;
        for (var i = 0; i < n; i++) aX += a[i] * sorted[i];

        var W = (aX * aX) / ss;
        W = Math.min(W, 1);

        // Approximate p-value using Royston (1992) transformation
        var mu, sigma, gamma;
        var lnN = Math.log(n);
        if (n <= 11) {
            gamma = 0.459 * n - 2.273;
            mu = -1.2725 + 1.0521 * gamma;
            sigma = 1.0308 - 0.26758 * gamma;
        } else {
            var ln2 = lnN * lnN;
            mu = 0.0038915 * ln2 * lnN - 0.083751 * ln2 + 0.31082 * lnN - 1.5861;
            sigma = Math.exp(0.0030302 * ln2 * lnN - 0.082676 * ln2 + 0.4803 * lnN - 2.2989);
        }

        var z = (Math.log(1 - W) - mu) / sigma;
        var p = 1 - normCdf(z);
        p = Math.max(0, Math.min(1, p));

        return { W: W, p: p };
    }

    // =========================================================================
    // Levene's Test (helper)
    // =========================================================================

    function leveneTest(/* ...groups */) {
        var groups = Array.isArray(arguments[0]) && Array.isArray(arguments[0][0])
            ? arguments[0]
            : Array.prototype.slice.call(arguments);
        if (groups.length < 2) return null;
        // Compute absolute deviations from group medians
        var deviations = groups.map(function(g) {
            var med = median(g);
            return g.map(function(v) { return Math.abs(v - med); });
        });
        // Do one-way ANOVA on deviations
        var allDevs = [];
        deviations.forEach(function(devs) { allDevs = allDevs.concat(devs); });
        var k = groups.length;
        var N = allDevs.length;
        var grandMean2 = allDevs.reduce(function(a,b){return a+b;},0) / N;
        var ssBetween = 0, ssWithin = 0;
        for (var i = 0; i < k; i++) {
            var gDevs = deviations[i];
            var gMean = gDevs.reduce(function(a,b){return a+b;},0) / gDevs.length;
            ssBetween += gDevs.length * Math.pow(gMean - grandMean2, 2);
            gDevs.forEach(function(d) { ssWithin += Math.pow(d - gMean, 2); });
        }
        var dfBetween = k - 1;
        var dfWithin = N - k;
        var F = (ssBetween / dfBetween) / (ssWithin / dfWithin);
        var p = 1 - jStat.centralF.cdf(F, dfBetween, dfWithin);
        return { F: F, p: p, leveneF: F, leveneP: p };
    }

    function _anovaCore(groups) {
        var k = groups.length;
        var allValues = [];
        var ns = [], means = [];
        for (var i = 0; i < k; i++) {
            ns.push(groups[i].length);
            means.push(mean(groups[i]));
            allValues = allValues.concat(groups[i]);
        }
        var N = allValues.length;
        var grandMean = mean(allValues);

        var ssBetween = 0, ssWithin = 0;
        for (var i = 0; i < k; i++) {
            ssBetween += ns[i] * (means[i] - grandMean) * (means[i] - grandMean);
            for (var j = 0; j < groups[i].length; j++) {
                ssWithin += (groups[i][j] - means[i]) * (groups[i][j] - means[i]);
            }
        }
        var dfBetween = k - 1;
        var dfWithin = N - k;
        var msBetween = ssBetween / dfBetween;
        var msWithin = ssWithin / dfWithin;
        var f = dfWithin > 0 ? msBetween / msWithin : NaN;
        var p = (f >= 0 && dfBetween > 0 && dfWithin > 0) ? (1 - fCdf(f, dfBetween, dfWithin)) : NaN;

        return {
            f: f, dfBetween: dfBetween, dfWithin: dfWithin, p: p,
            ssBetween: ssBetween, ssWithin: ssWithin, ssTotal: ssBetween + ssWithin,
            msBetween: msBetween, msWithin: msWithin
        };
    }

    // =========================================================================
    // Independent Samples t-Test
    // =========================================================================

    function independentTTest(group1, group2) {
        var g1 = clean(group1), g2 = clean(group2);
        if (g1.length < 2 || g2.length < 2) return null;

        var n1 = g1.length, n2 = g2.length;
        var m1 = mean(g1), m2 = mean(g2);
        var s1 = sd(g1, m1), s2 = sd(g2, m2);
        var v1 = variance(g1, m1), v2 = variance(g2, m2);
        var se1 = s1 / Math.sqrt(n1), se2 = s2 / Math.sqrt(n2);
        var meanDiff = m1 - m2;

        // Pooled SD and Student t
        var pooledVar = ((n1 - 1) * v1 + (n2 - 1) * v2) / (n1 + n2 - 2);
        var pooledSD = Math.sqrt(pooledVar);
        var sePooled = pooledSD * Math.sqrt(1 / n1 + 1 / n2);
        var t = sePooled > 0 ? meanDiff / sePooled : NaN;
        var df = n1 + n2 - 2;
        var p = !isNaN(t) ? 2 * (1 - tCdf(Math.abs(t), df)) : NaN;

        var tCrit = tInv(0.975, df);
        var ci95Lo = meanDiff - tCrit * sePooled;
        var ci95Hi = meanDiff + tCrit * sePooled;

        // Welch's t
        var seWelch = Math.sqrt(v1 / n1 + v2 / n2);
        var welchT = seWelch > 0 ? meanDiff / seWelch : NaN;
        var welchDfNum = Math.pow(v1 / n1 + v2 / n2, 2);
        var welchDfDen = Math.pow(v1 / n1, 2) / (n1 - 1) + Math.pow(v2 / n2, 2) / (n2 - 1);
        var welchDf = welchDfDen > 0 ? welchDfNum / welchDfDen : NaN;
        var welchP = !isNaN(welchT) && !isNaN(welchDf) ? 2 * (1 - tCdf(Math.abs(welchT), welchDf)) : NaN;

        // Cohen's d
        var cohensD = pooledSD > 0 ? meanDiff / pooledSD : NaN;

        // Levene's test
        var lev = leveneTest([g1, g2]);

        return {
            t: t, df: df, p: p,
            meanDiff: meanDiff, ci95_lo: ci95Lo, ci95_hi: ci95Hi,
            ci95: formatCI(ci95Lo, ci95Hi),
            cohensD: cohensD, effectSize: interpretCohensD(cohensD),
            leveneF: lev.leveneF, leveneP: lev.leveneP,
            welchT: welchT, welchDf: welchDf, welchP: welchP,
            desc1: { n: n1, mean: m1, sd: s1, se: se1 },
            desc2: { n: n2, mean: m2, sd: s2, se: se2 }
        };
    }

    // =========================================================================
    // Paired Samples t-Test
    // =========================================================================

    function pairedTTest(before, after) {
        var b = clean(before), a = clean(after);
        if (b.length < 2 || a.length < 2 || b.length !== a.length) return null;

        var n = b.length;
        var diffs = [];
        for (var i = 0; i < n; i++) diffs.push(a[i] - b[i]);

        var mB = mean(b), mA = mean(a);
        var sB = sd(b), sA = sd(a);
        var mdiff = mean(diffs);
        var sdiff = sd(diffs);
        var sediff = sdiff / Math.sqrt(n);
        var t = sediff > 0 ? mdiff / sediff : NaN;
        var df = n - 1;
        var p = !isNaN(t) ? 2 * (1 - tCdf(Math.abs(t), df)) : NaN;

        // Cohen's d for paired
        var cohensD = sdiff > 0 ? mdiff / sdiff : NaN;

        // Correlation between before and after
        var r = NaN, rP = NaN;
        if (n > 2) {
            var cov = 0;
            for (var i = 0; i < n; i++) cov += (b[i] - mB) * (a[i] - mA);
            cov /= (n - 1);
            var rDen = sB * sA;
            r = rDen > 0 ? cov / rDen : NaN;
            if (!isNaN(r) && Math.abs(r) < 1) {
                var tR = r * Math.sqrt((n - 2) / (1 - r * r));
                rP = 2 * (1 - tCdf(Math.abs(tR), n - 2));
            }
        }

        var tCritPaired = tInv(0.975, df);
        var ciLoPaired = mdiff - tCritPaired * sediff;
        var ciHiPaired = mdiff + tCritPaired * sediff;

        return {
            t: t, df: df, p: p,
            meanBefore: mB, meanAfter: mA, meanDiff: mdiff,
            sdDiff: sdiff, seDiff: sediff,
            ci95_lo: ciLoPaired, ci95_hi: ciHiPaired,
            ci95: formatCI(ciLoPaired, ciHiPaired),
            cohensD: cohensD, effectSize: interpretCohensD(cohensD),
            r: r, rP: rP,
            descBefore: { n: n, mean: mB, sd: sB, se: sB / Math.sqrt(n) },
            descAfter: { n: n, mean: mA, sd: sA, se: sA / Math.sqrt(n) }
        };
    }

    // =========================================================================
    // One-way ANOVA
    // =========================================================================

    function onewayAnova(groups, groupNames) {
        if (!Array.isArray(groups) || groups.length < 2) return null;
        var cleaned = groups.map(clean);
        for (var i = 0; i < cleaned.length; i++) {
            if (cleaned[i].length < 2) return null;
        }
        var names = groupNames || cleaned.map(function (_, i) { return "Group " + (i + 1); });

        var core = _anovaCore(cleaned);
        var etaSquared = core.ssTotal > 0 ? core.ssBetween / core.ssTotal : NaN;

        // Levene's test
        var lev = leveneTest(cleaned);

        // Descriptives
        var descs = [];
        for (var i = 0; i < cleaned.length; i++) {
            var m = mean(cleaned[i]), s = sd(cleaned[i]);
            var seVal = s / Math.sqrt(cleaned[i].length);
            var tCritAnova = cleaned[i].length > 1 ? tInv(0.975, cleaned[i].length - 1) : NaN;
            var ciLoAnova = m - tCritAnova * seVal;
            var ciHiAnova = m + tCritAnova * seVal;
            descs.push({ group: names[i], n: cleaned[i].length, mean: m, sd: s, se: seVal, ci95_lo: ciLoAnova, ci95_hi: ciHiAnova, ci95: formatCI(ciLoAnova, ciHiAnova) });
        }

        // Post-hoc (pairwise t-tests with Bonferroni correction) if significant
        var posthoc = [];
        if (core.p < 0.05) {
            var nComp = cleaned.length * (cleaned.length - 1) / 2;
            for (var i = 0; i < cleaned.length; i++) {
                for (var j = i + 1; j < cleaned.length; j++) {
                    var ni = cleaned[i].length, nj = cleaned[j].length;
                    var mi = mean(cleaned[i]), mj = mean(cleaned[j]);
                    var si = sd(cleaned[i]), sj = sd(cleaned[j]);
                    var pooledSD = Math.sqrt(((ni - 1) * si * si + (nj - 1) * sj * sj) / (ni + nj - 2));
                    var seP = pooledSD * Math.sqrt(1 / ni + 1 / nj);
                    var tVal = seP > 0 ? (mi - mj) / seP : NaN;
                    var dfP = ni + nj - 2;
                    var pVal = !isNaN(tVal) ? 2 * (1 - tCdf(Math.abs(tVal), dfP)) : NaN;
                    var adjP = Math.min(pVal * nComp, 1);
                    var d = pooledSD > 0 ? (mi - mj) / pooledSD : NaN;
                    posthoc.push({
                        groupA: names[i], groupB: names[j],
                        meanA: mi, meanB: mj, meanDiff: mi - mj,
                        t: tVal, p: adjP, cohensD: d
                    });
                }
            }
        }

        return {
            f: core.f, dfBetween: core.dfBetween, dfWithin: core.dfWithin, p: core.p,
            ssBetween: core.ssBetween, ssWithin: core.ssWithin, ssTotal: core.ssTotal,
            msBetween: core.msBetween, msWithin: core.msWithin,
            etaSquared: etaSquared, effectSize: interpretEtaSquared(etaSquared),
            leveneF: lev.leveneF, leveneP: lev.leveneP,
            descriptives: descs,
            posthoc: posthoc
        };
    }

    // =========================================================================
    // Mann-Whitney U
    // =========================================================================

    function mannWhitneyU(group1, group2) {
        var g1 = clean(group1), g2 = clean(group2);
        if (g1.length < 1 || g2.length < 1) return null;

        var n1 = g1.length, n2 = g2.length;
        var combined = [];
        for (var i = 0; i < n1; i++) combined.push({ v: g1[i], g: 1 });
        for (var i = 0; i < n2; i++) combined.push({ v: g2[i], g: 2 });
        combined.sort(function (a, b) { return a.v - b.v; });

        // Assign ranks with tie handling
        var rks = ranks(combined.map(function (x) { return x.v; }));

        var R1 = 0, R2 = 0;
        for (var i = 0; i < combined.length; i++) {
            if (combined[i].g === 1) R1 += rks[i]; else R2 += rks[i];
        }

        var U1 = n1 * n2 + n1 * (n1 + 1) / 2 - R1;
        var U2 = n1 * n2 + n2 * (n2 + 1) / 2 - R2;
        var U = Math.min(U1, U2);

        // Normal approximation
        var muU = n1 * n2 / 2;
        var sigmaU = Math.sqrt(n1 * n2 * (n1 + n2 + 1) / 12);
        var Z = sigmaU > 0 ? (U - muU) / sigmaU : 0;
        var p = 2 * (1 - normCdf(Math.abs(Z)));

        var r = Math.sqrt(n1 + n2) > 0 ? Math.abs(Z) / Math.sqrt(n1 + n2) : NaN;

        return {
            U: U, Z: Z, p: p, r: r, effectSize: interpretR(r),
            desc1: { n: n1, median: median(g1), mean: mean(g1), sd: sd(g1), meanRank: R1 / n1, sumRank: R1 },
            desc2: { n: n2, median: median(g2), mean: mean(g2), sd: sd(g2), meanRank: R2 / n2, sumRank: R2 }
        };
    }

    // =========================================================================
    // Wilcoxon Signed-Rank
    // =========================================================================

    function wilcoxonSignedRank(before, after) {
        var b = clean(before), a = clean(after);
        if (b.length < 2 || a.length < 2 || b.length !== a.length) return null;

        var n = b.length;
        var diffs = [];
        var nTies = 0;
        for (var i = 0; i < n; i++) {
            var d = a[i] - b[i];
            if (d === 0) { nTies++; } else { diffs.push(d); }
        }

        var nEff = diffs.length;
        if (nEff === 0) return null;

        var absDiffs = diffs.map(Math.abs);
        var rks = ranks(absDiffs);

        var Wpos = 0, Wneg = 0, nPos = 0, nNeg = 0;
        for (var i = 0; i < nEff; i++) {
            if (diffs[i] > 0) { Wpos += rks[i]; nPos++; }
            else { Wneg += rks[i]; nNeg++; }
        }

        var W = Math.min(Wpos, Wneg);
        var muW = nEff * (nEff + 1) / 4;
        var sigmaW = Math.sqrt(nEff * (nEff + 1) * (2 * nEff + 1) / 24);
        var Z = sigmaW > 0 ? (W - muW) / sigmaW : 0;
        var p = 2 * (1 - normCdf(Math.abs(Z)));
        var r = n > 0 ? Math.abs(Z) / Math.sqrt(n) : NaN;

        return {
            W: W, Z: Z, p: p, r: r, effectSize: interpretR(r),
            nPos: nPos, nNeg: nNeg, nTies: nTies,
            descBefore: { n: n, median: median(b), mean: mean(b), sd: sd(b) },
            descAfter: { n: n, median: median(a), mean: mean(a), sd: sd(a) }
        };
    }

    // =========================================================================
    // Kruskal-Wallis
    // =========================================================================

    function kruskalWallis(groups, groupNames) {
        if (!Array.isArray(groups) || groups.length < 2) return null;
        var cleaned = groups.map(clean);
        for (var i = 0; i < cleaned.length; i++) {
            if (cleaned[i].length < 1) return null;
        }
        var names = groupNames || cleaned.map(function (_, i) { return "Group " + (i + 1); });

        var combined = [];
        for (var g = 0; g < cleaned.length; g++) {
            for (var i = 0; i < cleaned[g].length; i++) {
                combined.push({ v: cleaned[g][i], g: g });
            }
        }
        var rks = ranks(combined.map(function (x) { return x.v; }));

        var N = combined.length;
        var ns = cleaned.map(function (g) { return g.length; });
        var rankSums = new Array(cleaned.length).fill(0);
        for (var i = 0; i < combined.length; i++) {
            rankSums[combined[i].g] += rks[i];
        }

        var H = 0;
        for (var g = 0; g < cleaned.length; g++) {
            var meanRank = rankSums[g] / ns[g];
            H += ns[g] * Math.pow(meanRank - (N + 1) / 2, 2);
        }
        H = (12 / (N * (N + 1))) * H;  // simplified: sum(Ri^2/ni) form
        // Recalculate using standard formula
        H = 0;
        for (var g = 0; g < cleaned.length; g++) {
            H += (rankSums[g] * rankSums[g]) / ns[g];
        }
        H = (12 / (N * (N + 1))) * H - 3 * (N + 1);

        var df = cleaned.length - 1;
        var p = H >= 0 && df > 0 ? (1 - chi2Cdf(H, df)) : NaN;
        var etaH = (H - df) / (N - df - 1);
        etaH = Math.max(0, etaH);

        // Descriptives
        var descs = [];
        for (var g = 0; g < cleaned.length; g++) {
            descs.push({
                group: names[g], n: ns[g], median: median(cleaned[g]),
                mean: mean(cleaned[g]), sd: sd(cleaned[g]),
                meanRank: rankSums[g] / ns[g]
            });
        }

        // Post-hoc pairwise Mann-Whitney if significant
        var posthoc = [];
        if (p < 0.05) {
            var nComp = cleaned.length * (cleaned.length - 1) / 2;
            for (var i = 0; i < cleaned.length; i++) {
                for (var j = i + 1; j < cleaned.length; j++) {
                    var mw = mannWhitneyU(cleaned[i], cleaned[j]);
                    if (mw) {
                        posthoc.push({
                            groupA: names[i], groupB: names[j],
                            U: mw.U, Z: mw.Z, p: Math.min(mw.p * nComp, 1), r: mw.r
                        });
                    }
                }
            }
        }

        return {
            H: H, df: df, p: p, etaSquaredH: etaH,
            effectSize: interpretEtaSquared(etaH),
            descriptives: descs, posthoc: posthoc
        };
    }

    // =========================================================================
    // Friedman Test
    // =========================================================================

    function friedmanTest(variables) {
        if (!Array.isArray(variables) || variables.length < 2) return null;
        var cleaned = variables.map(clean);
        // All must have same length
        var n = cleaned[0].length;
        for (var i = 1; i < cleaned.length; i++) {
            if (cleaned[i].length !== n) return null;
        }
        if (n < 2) return null;

        var k = cleaned.length;

        // Rank within each subject (row)
        var rankSums = new Array(k).fill(0);
        for (var subj = 0; subj < n; subj++) {
            var row = [];
            for (var v = 0; v < k; v++) row.push(cleaned[v][subj]);
            var rks = ranks(row);
            for (var v = 0; v < k; v++) rankSums[v] += rks[v];
        }

        var meanRankOverall = (k + 1) / 2;
        var ssRanks = 0;
        for (var v = 0; v < k; v++) {
            var mr = rankSums[v] / n;
            ssRanks += (mr - meanRankOverall) * (mr - meanRankOverall);
        }

        var chiSquare = (12 * n / (k * (k + 1))) * ssRanks;
        var df = k - 1;
        var p = chiSquare >= 0 && df > 0 ? (1 - chi2Cdf(chiSquare, df)) : NaN;
        var kendallW = chiSquare / (n * (k - 1));
        kendallW = Math.max(0, Math.min(1, kendallW));

        var names = [];
        for (var v = 0; v < k; v++) names.push("Variable " + (v + 1));

        var descs = [];
        for (var v = 0; v < k; v++) {
            descs.push({
                variable: names[v], n: n, median: median(cleaned[v]),
                mean: mean(cleaned[v]), sd: sd(cleaned[v]),
                meanRank: rankSums[v] / n
            });
        }

        // Post-hoc pairwise Wilcoxon if significant
        var posthoc = [];
        if (p < 0.05) {
            var nComp = k * (k - 1) / 2;
            for (var i = 0; i < k; i++) {
                for (var j = i + 1; j < k; j++) {
                    var w = wilcoxonSignedRank(cleaned[i], cleaned[j]);
                    if (w) {
                        posthoc.push({
                            variableA: names[i], variableB: names[j],
                            W: w.W, Z: w.Z, p: Math.min(w.p * nComp, 1), r: w.r
                        });
                    }
                }
            }
        }

        return {
            chiSquare: chiSquare, df: df, p: p, kendallW: kendallW,
            descriptives: descs, posthoc: posthoc
        };
    }

    // =========================================================================
    // Chi-Square Test of Independence
    // =========================================================================

    function chiSquare(var1Values, var2Values) {
        if (!isValid(var1Values) || !isValid(var2Values) || var1Values.length !== var2Values.length) return null;
        var n = var1Values.length;

        // Build contingency table
        var rowLabels = [], colLabels = [];
        var rowMap = {}, colMap = {};
        for (var i = 0; i < n; i++) {
            var r = String(var1Values[i]), c = String(var2Values[i]);
            if (!(r in rowMap)) { rowMap[r] = rowLabels.length; rowLabels.push(r); }
            if (!(c in colMap)) { colMap[c] = colLabels.length; colLabels.push(c); }
        }

        var nRows = rowLabels.length, nCols = colLabels.length;
        if (nRows < 2 || nCols < 2) return null;

        var observed = [];
        for (var i = 0; i < nRows; i++) { observed.push(new Array(nCols).fill(0)); }
        for (var i = 0; i < n; i++) {
            observed[rowMap[String(var1Values[i])]][colMap[String(var2Values[i])]]++;
        }

        // Row and column totals
        var rowTotals = new Array(nRows).fill(0);
        var colTotals = new Array(nCols).fill(0);
        for (var i = 0; i < nRows; i++) {
            for (var j = 0; j < nCols; j++) {
                rowTotals[i] += observed[i][j];
                colTotals[j] += observed[i][j];
            }
        }

        // Expected frequencies
        var expected = [];
        for (var i = 0; i < nRows; i++) {
            expected.push([]);
            for (var j = 0; j < nCols; j++) {
                expected[i].push((rowTotals[i] * colTotals[j]) / n);
            }
        }

        // Chi-square statistic
        var chi2 = 0;
        for (var i = 0; i < nRows; i++) {
            for (var j = 0; j < nCols; j++) {
                if (expected[i][j] > 0) {
                    chi2 += Math.pow(observed[i][j] - expected[i][j], 2) / expected[i][j];
                }
            }
        }

        var df = (nRows - 1) * (nCols - 1);
        var p = chi2 >= 0 && df > 0 ? (1 - chi2Cdf(chi2, df)) : NaN;

        // Cramer's V
        var minDim = Math.min(nRows - 1, nCols - 1);
        var cramersV = minDim > 0 ? Math.sqrt(chi2 / (n * minDim)) : NaN;

        // Phi (for 2x2)
        var phi = n > 0 ? Math.sqrt(chi2 / n) : NaN;

        return {
            chiSquare: chi2, df: df, p: p,
            cramersV: cramersV, effectSize: interpretCramersV(cramersV),
            phi: phi, n: n,
            observed: observed, expected: expected,
            rowLabels: rowLabels, colLabels: colLabels
        };
    }

    // =========================================================================
    // Correlation
    // =========================================================================

    function pearsonR(x, y) {
        var n = x.length;
        if (n < 3) return { r: NaN, p: NaN };
        var mx = mean(x), my = mean(y);
        var num = 0, dx2 = 0, dy2 = 0;
        for (var i = 0; i < n; i++) {
            var dx = x[i] - mx, dy = y[i] - my;
            num += dx * dy; dx2 += dx * dx; dy2 += dy * dy;
        }
        var den = Math.sqrt(dx2 * dy2);
        var r = den > 0 ? num / den : NaN;
        var tVal = NaN, p = NaN;
        if (!isNaN(r) && Math.abs(r) < 1) {
            tVal = r * Math.sqrt((n - 2) / (1 - r * r));
            p = 2 * (1 - tCdf(Math.abs(tVal), n - 2));
        } else if (Math.abs(r) === 1) {
            p = 0;
        }
        return { r: r, p: p };
    }

    function kendallTau(x, y) {
        var n = x.length;
        if (n < 3) return { r: NaN, p: NaN };
        var concordant = 0, discordant = 0;
        for (var i = 0; i < n - 1; i++) {
            for (var j = i + 1; j < n; j++) {
                var xd = x[j] - x[i], yd = y[j] - y[i];
                if (xd * yd > 0) concordant++;
                else if (xd * yd < 0) discordant++;
            }
        }
        var nPairs = n * (n - 1) / 2;
        var tau = nPairs > 0 ? (concordant - discordant) / nPairs : NaN;
        // Normal approximation for p-value
        var sigma = Math.sqrt(2 * (2 * n + 5) / (9 * n * (n - 1)));
        var z = sigma > 0 ? tau / sigma : 0;
        var p = 2 * (1 - normCdf(Math.abs(z)));
        return { r: tau, p: p };
    }

    function correlation(data, varNames, method) {
        if (!Array.isArray(data) || data.length < 2) return null;
        var k = data.length;
        var names = varNames || data.map(function (_, i) { return "Var" + (i + 1); });
        method = method || "pearson";

        var matrix = [], pMatrix = [], pairs = [];

        for (var i = 0; i < k; i++) {
            matrix.push(new Array(k).fill(0));
            pMatrix.push(new Array(k).fill(0));
        }

        for (var i = 0; i < k; i++) {
            matrix[i][i] = 1;
            pMatrix[i][i] = 0;
            for (var j = i + 1; j < k; j++) {
                // Get paired data (only where both have valid values)
                var xi = clean(data[i]), xj = clean(data[j]);
                var minLen = Math.min(xi.length, xj.length);
                xi = xi.slice(0, minLen);
                xj = xj.slice(0, minLen);

                var result;
                if (method === "spearman") {
                    var ri = ranks(xi), rj = ranks(xj);
                    result = pearsonR(ri, rj);
                } else if (method === "kendall") {
                    result = kendallTau(xi, xj);
                } else {
                    result = pearsonR(xi, xj);
                }

                matrix[i][j] = result.r;
                matrix[j][i] = result.r;
                pMatrix[i][j] = result.p;
                pMatrix[j][i] = result.p;

                pairs.push({
                    var1: names[i], var2: names[j],
                    r: result.r, p: result.p,
                    sig: result.p < 0.05,
                    strength: interpretR(result.r),
                    rSquared: result.r * result.r,
                    direction: result.r > 0 ? "Positive" : (result.r < 0 ? "Negative" : "None")
                });
            }
        }

        return { matrix: matrix, pMatrix: pMatrix, pairs: pairs };
    }

    // =========================================================================
    // Linear Regression
    // =========================================================================

    function linearRegression(y, xs, varNames) {
        var yClean = clean(y);
        if (yClean.length < 3 || !Array.isArray(xs) || xs.length < 1) return null;

        var n = yClean.length;
        var p = xs.length; // number of predictors
        var names = varNames || ["(Intercept)"].concat(xs.map(function (_, i) { return "X" + (i + 1); }));

        // Build design matrix with intercept column
        var X = []; // n x (p+1)
        for (var i = 0; i < n; i++) {
            var row = [1]; // intercept
            for (var j = 0; j < p; j++) {
                row.push(Number(xs[j][i]) || 0);
            }
            X.push(row);
        }

        var cols = p + 1;

        // X'X
        var XtX = [];
        for (var i = 0; i < cols; i++) {
            XtX.push(new Array(cols).fill(0));
            for (var j = 0; j < cols; j++) {
                for (var k = 0; k < n; k++) {
                    XtX[i][j] += X[k][i] * X[k][j];
                }
            }
        }

        // X'y
        var Xty = new Array(cols).fill(0);
        for (var i = 0; i < cols; i++) {
            for (var k = 0; k < n; k++) {
                Xty[i] += X[k][i] * yClean[k];
            }
        }

        // Invert XtX using Gauss-Jordan
        var inv = invertMatrix(XtX);
        if (!inv) return null;

        // Coefficients: b = (X'X)^-1 X'y
        var b = new Array(cols).fill(0);
        for (var i = 0; i < cols; i++) {
            for (var j = 0; j < cols; j++) {
                b[i] += inv[i][j] * Xty[j];
            }
        }

        // Predicted values and residuals
        var yhat = new Array(n).fill(0);
        for (var i = 0; i < n; i++) {
            for (var j = 0; j < cols; j++) {
                yhat[i] += X[i][j] * b[j];
            }
        }

        var yMean = mean(yClean);
        var ssTotal = 0, ssRes = 0;
        for (var i = 0; i < n; i++) {
            ssTotal += (yClean[i] - yMean) * (yClean[i] - yMean);
            ssRes += (yClean[i] - yhat[i]) * (yClean[i] - yhat[i]);
        }
        var ssReg = ssTotal - ssRes;

        var dfReg = p;
        var dfRes = n - p - 1;
        var msReg = dfReg > 0 ? ssReg / dfReg : NaN;
        var msRes = dfRes > 0 ? ssRes / dfRes : NaN;
        var fStat = msRes > 0 ? msReg / msRes : NaN;
        var fP = (!isNaN(fStat) && dfReg > 0 && dfRes > 0) ? (1 - fCdf(fStat, dfReg, dfRes)) : NaN;

        var rSquared = ssTotal > 0 ? ssReg / ssTotal : NaN;
        var adjRSquared = (n > p + 1 && ssTotal > 0) ? 1 - (1 - rSquared) * (n - 1) / (n - p - 1) : NaN;
        var r = !isNaN(rSquared) ? Math.sqrt(rSquared) : NaN;

        // Durbin-Watson
        var dw = 0, dwDen = 0;
        var residuals = [];
        for (var i = 0; i < n; i++) residuals.push(yClean[i] - yhat[i]);
        for (var i = 1; i < n; i++) dw += Math.pow(residuals[i] - residuals[i - 1], 2);
        for (var i = 0; i < n; i++) dwDen += residuals[i] * residuals[i];
        dw = dwDen > 0 ? dw / dwDen : NaN;

        // Coefficient statistics
        var seBeta = [];
        for (var i = 0; i < cols; i++) {
            seBeta.push(msRes > 0 && inv[i][i] >= 0 ? Math.sqrt(msRes * inv[i][i]) : NaN);
        }

        var coefficients = [];
        for (var i = 0; i < cols; i++) {
            var tVal = seBeta[i] > 0 ? b[i] / seBeta[i] : NaN;
            var pVal = !isNaN(tVal) && dfRes > 0 ? 2 * (1 - tCdf(Math.abs(tVal), dfRes)) : NaN;
            var tCritVal = dfRes > 0 ? tInv(0.975, dfRes) : NaN;
            var ciLoReg = b[i] - tCritVal * seBeta[i];
            var ciHiReg = b[i] + tCritVal * seBeta[i];
            coefficients.push({
                variable: names[i] || (i === 0 ? "(Intercept)" : "X" + i),
                b: b[i], se: seBeta[i], t: tVal, p: pVal,
                ci95Lo: ciLoReg,
                ci95Hi: ciHiReg,
                ci95: formatCI(ciLoReg, ciHiReg)
            });
        }

        return {
            r: r, rSquared: rSquared, adjRSquared: adjRSquared,
            f: fStat, fP: fP, durbinWatson: dw,
            coefficients: coefficients,
            anova: { ssReg: ssReg, ssRes: ssRes, ssTotal: ssTotal, dfReg: dfReg, dfRes: dfRes, msReg: msReg, msRes: msRes }
        };
    }

    function invertMatrix(m) {
        var n = m.length;
        // Augment with identity
        var aug = [];
        for (var i = 0; i < n; i++) {
            aug.push([]);
            for (var j = 0; j < n; j++) aug[i].push(m[i][j]);
            for (var j = 0; j < n; j++) aug[i].push(i === j ? 1 : 0);
        }

        for (var col = 0; col < n; col++) {
            // Partial pivoting
            var maxRow = col, maxVal = Math.abs(aug[col][col]);
            for (var row = col + 1; row < n; row++) {
                if (Math.abs(aug[row][col]) > maxVal) {
                    maxVal = Math.abs(aug[row][col]);
                    maxRow = row;
                }
            }
            if (maxVal < 1e-12) return null; // singular
            if (maxRow !== col) { var tmp = aug[col]; aug[col] = aug[maxRow]; aug[maxRow] = tmp; }

            var pivot = aug[col][col];
            for (var j = 0; j < 2 * n; j++) aug[col][j] /= pivot;

            for (var row = 0; row < n; row++) {
                if (row === col) continue;
                var factor = aug[row][col];
                for (var j = 0; j < 2 * n; j++) aug[row][j] -= factor * aug[col][j];
            }
        }

        var inv = [];
        for (var i = 0; i < n; i++) {
            inv.push(aug[i].slice(n));
        }
        return inv;
    }

    // =========================================================================
    // Logistic Regression (simplified IRLS)
    // =========================================================================

    function logisticRegression(y, xs, varNames) {
        var yClean = clean(y);
        if (yClean.length < 3 || !Array.isArray(xs) || xs.length < 1) return null;

        var n = yClean.length;
        var p = xs.length;
        var names = varNames || ["(Intercept)"].concat(xs.map(function (_, i) { return "X" + (i + 1); }));

        // Build design matrix
        var X = [];
        for (var i = 0; i < n; i++) {
            var row = [1];
            for (var j = 0; j < p; j++) row.push(Number(xs[j][i]) || 0);
            X.push(row);
        }

        var cols = p + 1;

        // Sigmoid
        function sigmoid(z) {
            if (z > 500) return 1;
            if (z < -500) return 0;
            return 1 / (1 + Math.exp(-z));
        }

        // IRLS iterations
        var beta = new Array(cols).fill(0);
        var maxIter = 25;
        var converged = false;

        for (var iter = 0; iter < maxIter; iter++) {
            // Compute probabilities
            var prob = new Array(n);
            for (var i = 0; i < n; i++) {
                var z = 0;
                for (var j = 0; j < cols; j++) z += X[i][j] * beta[j];
                prob[i] = sigmoid(z);
            }

            // W diagonal and working response
            // X'WX and X'Wz using IRLS
            var XtWX = [];
            for (var i = 0; i < cols; i++) {
                XtWX.push(new Array(cols).fill(0));
                for (var j = 0; j < cols; j++) {
                    for (var k = 0; k < n; k++) {
                        var w = prob[k] * (1 - prob[k]);
                        w = Math.max(w, 1e-10);
                        XtWX[i][j] += X[k][i] * w * X[k][j];
                    }
                }
            }

            // X' (y - p)
            var score = new Array(cols).fill(0);
            for (var i = 0; i < cols; i++) {
                for (var k = 0; k < n; k++) {
                    score[i] += X[k][i] * (yClean[k] - prob[k]);
                }
            }

            var inv = invertMatrix(XtWX);
            if (!inv) break;

            // Update: beta_new = beta + (X'WX)^-1 * X'(y-p)
            var delta = new Array(cols).fill(0);
            var maxDelta = 0;
            for (var i = 0; i < cols; i++) {
                for (var j = 0; j < cols; j++) {
                    delta[i] += inv[i][j] * score[j];
                }
                beta[i] += delta[i];
                maxDelta = Math.max(maxDelta, Math.abs(delta[i]));
            }

            if (maxDelta < 1e-8) { converged = true; break; }
        }

        // Final probabilities and log-likelihood
        var logLik = 0;
        var correct = 0;
        for (var i = 0; i < n; i++) {
            var z = 0;
            for (var j = 0; j < cols; j++) z += X[i][j] * beta[j];
            var pi = sigmoid(z);
            var pred = pi >= 0.5 ? 1 : 0;
            if (pred === yClean[i]) correct++;
            logLik += yClean[i] * Math.log(Math.max(pi, 1e-15)) + (1 - yClean[i]) * Math.log(Math.max(1 - pi, 1e-15));
        }

        var accuracy = correct / n;
        var aic = -2 * logLik + 2 * cols;

        // Standard errors from final Hessian
        var prob2 = new Array(n);
        for (var i = 0; i < n; i++) {
            var z = 0;
            for (var j = 0; j < cols; j++) z += X[i][j] * beta[j];
            prob2[i] = sigmoid(z);
        }
        var H = [];
        for (var i = 0; i < cols; i++) {
            H.push(new Array(cols).fill(0));
            for (var j = 0; j < cols; j++) {
                for (var k = 0; k < n; k++) {
                    var w = prob2[k] * (1 - prob2[k]);
                    w = Math.max(w, 1e-10);
                    H[i][j] += X[k][i] * w * X[k][j];
                }
            }
        }
        var covMatrix = invertMatrix(H);

        var coefficients = [];
        for (var i = 0; i < cols; i++) {
            var seVal = (covMatrix && covMatrix[i][i] >= 0) ? Math.sqrt(covMatrix[i][i]) : NaN;
            var wald = seVal > 0 ? Math.pow(beta[i] / seVal, 2) : NaN;
            var pVal = !isNaN(wald) ? (1 - chi2Cdf(wald, 1)) : NaN;
            coefficients.push({
                variable: names[i] || (i === 0 ? "(Intercept)" : "X" + i),
                b: beta[i], se: seVal, wald: wald, p: pVal,
                or: Math.exp(beta[i])
            });
        }

        return {
            coefficients: coefficients,
            accuracy: accuracy,
            aic: aic
        };
    }

    // =========================================================================
    // Cronbach's Alpha
    // =========================================================================

    function cronbachAlpha(items) {
        if (!Array.isArray(items) || items.length < 2) return null;
        var cleaned = items.map(clean);
        var k = cleaned.length;
        var n = cleaned[0].length;
        for (var i = 1; i < k; i++) {
            if (cleaned[i].length !== n) return null;
        }
        if (n < 2) return null;

        // Item variances
        var itemVars = cleaned.map(function (item) { return variance(item); });
        var sumItemVar = sum(itemVars);

        // Total scores
        var totals = new Array(n).fill(0);
        for (var i = 0; i < k; i++) {
            for (var j = 0; j < n; j++) {
                totals[j] += cleaned[i][j];
            }
        }
        var totalVar = variance(totals);

        var alpha = totalVar > 0 ? (k / (k - 1)) * (1 - sumItemVar / totalVar) : NaN;

        // Item statistics
        var itemStats = [];
        for (var i = 0; i < k; i++) {
            // Item-total correlation (corrected: correlate item with total minus item)
            var correctedTotals = new Array(n);
            for (var j = 0; j < n; j++) correctedTotals[j] = totals[j] - cleaned[i][j];
            var itr = pearsonR(cleaned[i], correctedTotals).r;

            // Alpha if deleted
            var remainingItems = [];
            for (var q = 0; q < k; q++) { if (q !== i) remainingItems.push(cleaned[q]); }
            var remainK = remainingItems.length;
            var rItemVars = remainingItems.map(function (item) { return variance(item); });
            var rSumItemVar = sum(rItemVars);
            var rTotals = new Array(n).fill(0);
            for (var q = 0; q < remainK; q++) {
                for (var j = 0; j < n; j++) rTotals[j] += remainingItems[q][j];
            }
            var rTotalVar = variance(rTotals);
            var alphaIfDel = rTotalVar > 0 ? (remainK / (remainK - 1)) * (1 - rSumItemVar / rTotalVar) : NaN;

            itemStats.push({
                item: "Item " + (i + 1),
                mean: mean(cleaned[i]),
                sd: sd(cleaned[i]),
                itemTotalR: itr,
                alphaIfDeleted: alphaIfDel,
                shouldDelete: alphaIfDel > alpha + 0.01
            });
        }

        return {
            alpha: alpha, interpretation: interpretAlpha(alpha),
            nItems: k, nCases: n,
            itemStats: itemStats
        };
    }

    // =========================================================================
    // Likert Analysis
    // =========================================================================

    function likertAnalysis(data, varNames, scale, criteria) {
        if (!Array.isArray(data) || data.length < 1) return null;
        scale = scale || 5;
        var names = varNames || data.map(function (_, i) { return "Item " + (i + 1); });

        if (!criteria) {
            if (scale === 5) {
                criteria = [
                    { lo: 1.00, hi: 1.80, label: "Strongly Disagree" },
                    { lo: 1.81, hi: 2.60, label: "Disagree" },
                    { lo: 2.61, hi: 3.40, label: "Neutral" },
                    { lo: 3.41, hi: 4.20, label: "Agree" },
                    { lo: 4.21, hi: 5.00, label: "Strongly Agree" }
                ];
            } else if (scale === 3) {
                criteria = [
                    { lo: 1.00, hi: 1.67, label: "Low" },
                    { lo: 1.68, hi: 2.33, label: "Medium" },
                    { lo: 2.34, hi: 3.00, label: "High" }
                ];
            } else {
                criteria = [];
            }
        }

        function getInterpretation(m) {
            for (var c = 0; c < criteria.length; c++) {
                if (m >= criteria[c].lo && m <= criteria[c].hi) return criteria[c].label;
            }
            return "N/A";
        }

        var items = [];
        var allValues = [];

        for (var i = 0; i < data.length; i++) {
            var v = clean(data[i]);
            var m = mean(v), s = sd(v), sem = se(v);

            // Frequency distribution
            var freq = {};
            for (var lv = 1; lv <= scale; lv++) {
                var count = 0;
                for (var j = 0; j < v.length; j++) { if (v[j] === lv) count++; }
                freq[lv] = { count: count, pct: v.length > 0 ? (count / v.length * 100) : 0 };
            }

            items.push({
                no: i + 1, variable: names[i],
                frequencies: freq, mean: m, sd: s, se: sem,
                interpretation: getInterpretation(m)
            });

            allValues = allValues.concat(v);
        }

        var overallMean = mean(allValues);
        var overallSD = sd(allValues);

        // Ranking by mean (descending)
        var ranked = items.slice().sort(function (a, b) { return b.mean - a.mean; });
        var ranking = ranked.map(function (item, idx) {
            return { rank: idx + 1, variable: item.variable, mean: item.mean, sd: item.sd, interpretation: item.interpretation };
        });

        return {
            items: items,
            overall: { mean: overallMean, sd: overallSD, interpretation: getInterpretation(overallMean) },
            ranking: ranking
        };
    }

    // =========================================================================
    // Assumption Checking
    // =========================================================================

    function checkAssumptions(testType, data) {
        var checks = [];
        var allPassed = true;

        switch(testType) {
            case 'independent-ttest': {
                // data = { group1: [], group2: [] }
                var n1 = shapiroWilk(data.group1);
                var n2 = shapiroWilk(data.group2);
                var norm1Pass = n1 && n1.p > 0.05;
                var norm2Pass = n2 && n2.p > 0.05;
                checks.push({name: 'Normality (Group 1)', result: norm1Pass ? 'Normal' : 'Not Normal', passed: norm1Pass, detail: 'Shapiro-Wilk p = ' + (n1 ? formatPValue(n1.p) : 'N/A')});
                checks.push({name: 'Normality (Group 2)', result: norm2Pass ? 'Normal' : 'Not Normal', passed: norm2Pass, detail: 'Shapiro-Wilk p = ' + (n2 ? formatPValue(n2.p) : 'N/A')});
                var lev = leveneTest(data.group1, data.group2);
                var levPass = lev && lev.p > 0.05;
                checks.push({name: 'Homogeneity of Variance', result: levPass ? 'Equal' : 'Unequal', passed: levPass, detail: "Levene's F = " + (lev ? lev.F.toFixed(4) : 'N/A') + ', p = ' + (lev ? formatPValue(lev.p) : 'N/A')});
                var sizeOk = data.group1.length >= 5 && data.group2.length >= 5;
                checks.push({name: 'Sample Size', result: 'n1=' + data.group1.length + ', n2=' + data.group2.length, passed: sizeOk, detail: sizeOk ? 'Adequate' : 'Very small sample, consider non-parametric'});

                if (!norm1Pass || !norm2Pass) allPassed = false;
                if (!levPass) allPassed = false;

                var rec = allPassed ? 'Assumptions met. Independent t-test is appropriate.' : '';
                if (!norm1Pass || !norm2Pass) rec = 'Normality assumption violated. Consider Mann-Whitney U test.';
                if (norm1Pass && norm2Pass && !levPass) rec = "Variance not equal. Welch's t-test will be used automatically.";

                return {checks: checks, passed: allPassed, recommendation: rec};
            }
            case 'paired-ttest': {
                var diff = [];
                var nP = Math.min(data.before.length, data.after.length);
                for (var i = 0; i < nP; i++) diff.push(data.after[i] - data.before[i]);
                var normDiff = shapiroWilk(diff);
                var normPass = normDiff && normDiff.p > 0.05;
                checks.push({name: 'Normality of Differences', result: normPass ? 'Normal' : 'Not Normal', passed: normPass, detail: 'Shapiro-Wilk p = ' + (normDiff ? formatPValue(normDiff.p) : 'N/A')});
                var sizeOk = nP >= 5;
                checks.push({name: 'Sample Size', result: 'n = ' + nP, passed: sizeOk, detail: sizeOk ? 'Adequate' : 'Very small'});

                return {checks: checks, passed: normPass, recommendation: normPass ? 'Assumptions met. Paired t-test is appropriate.' : 'Normality violated. Consider Wilcoxon Signed-Rank test.'};
            }
            case 'oneway-anova': {
                // data = { groups: [[...], [...], ...], names: [...] }
                var allNorm = true;
                for (var g = 0; g < data.groups.length; g++) {
                    var sw = shapiroWilk(data.groups[g]);
                    var p = sw && sw.p > 0.05;
                    if (!p) allNorm = false;
                    checks.push({name: 'Normality (' + (data.names[g] || 'Group '+(g+1)) + ')', result: p ? 'Normal' : 'Not Normal', passed: p, detail: 'Shapiro-Wilk p = ' + (sw ? formatPValue(sw.p) : 'N/A')});
                }
                var lev = leveneTest.apply(null, data.groups);
                var levPass = lev && lev.p > 0.05;
                checks.push({name: 'Homogeneity of Variance', result: levPass ? 'Equal' : 'Unequal', passed: levPass, detail: "Levene's F = " + (lev ? lev.F.toFixed(4) : 'N/A') + ', p = ' + (lev ? formatPValue(lev.p) : 'N/A')});

                var rec = (allNorm && levPass) ? 'Assumptions met. One-way ANOVA is appropriate.' : '';
                if (!allNorm) rec = 'Normality violated. Consider Kruskal-Wallis H test.';
                if (allNorm && !levPass) rec = "Variance not equal. Consider Welch's ANOVA or Kruskal-Wallis.";

                return {checks: checks, passed: allNorm && levPass, recommendation: rec};
            }
            case 'correlation': {
                // data = { vars: [[...], [...]] }
                var allNorm = true;
                for (var v = 0; v < data.vars.length; v++) {
                    var sw = shapiroWilk(data.vars[v]);
                    var p = sw && sw.p > 0.05;
                    if (!p) allNorm = false;
                    checks.push({name: 'Normality (Var ' + (v+1) + ')', result: p ? 'Normal' : 'Not Normal', passed: p, detail: 'p = ' + (sw ? formatPValue(sw.p) : 'N/A')});
                }
                return {checks: checks, passed: allNorm, recommendation: allNorm ? 'Normality met. Pearson correlation is appropriate.' : 'Non-normal data. Consider Spearman or Kendall correlation.'};
            }
            default:
                return {checks: [], passed: true, recommendation: ''};
        }
    }

    // =========================================================================
    // Odds Ratio / Relative Risk
    // =========================================================================

    function oddsRatioRR(a, b, c, d) {
        // 2x2 table: a=exposed+outcome, b=exposed+no outcome, c=unexposed+outcome, d=unexposed+no outcome
        if (b*c === 0) return null;
        var or = (a*d) / (b*c);
        var rr = (a/(a+b)) / (c/(c+d));
        // CI for OR
        var seLnOR = Math.sqrt(1/Math.max(a,0.5) + 1/Math.max(b,0.5) + 1/Math.max(c,0.5) + 1/Math.max(d,0.5));
        var orCIlo = Math.exp(Math.log(or) - 1.96 * seLnOR);
        var orCIhi = Math.exp(Math.log(or) + 1.96 * seLnOR);
        // CI for RR
        var seLnRR = Math.sqrt(1/Math.max(a,0.5) - 1/Math.max(a+b,1) + 1/Math.max(c,0.5) - 1/Math.max(c+d,1));
        var rrCIlo = Math.exp(Math.log(rr) - 1.96 * seLnRR);
        var rrCIhi = Math.exp(Math.log(rr) + 1.96 * seLnRR);
        return {
            or: or, orCI: formatCI(orCIlo, orCIhi),
            rr: rr, rrCI: formatCI(rrCIlo, rrCIhi),
            a: a, b: b, c: c, d: d,
            n: a+b+c+d
        };
    }

    // =========================================================================
    // Export
    // =========================================================================

    window.Stats = {
        // Basic statistics
        descriptive: descriptive,

        // Normality
        shapiroWilk: shapiroWilk,
        ksTest: ksTest,

        // t-tests
        independentTTest: independentTTest,
        pairedTTest: pairedTTest,

        // ANOVA
        onewayAnova: onewayAnova,

        // Non-parametric
        mannWhitneyU: mannWhitneyU,
        wilcoxonSignedRank: wilcoxonSignedRank,
        kruskalWallis: kruskalWallis,
        friedmanTest: friedmanTest,
        chiSquare: chiSquare,

        // Correlation
        correlation: correlation,

        // Regression
        linearRegression: linearRegression,
        logisticRegression: logisticRegression,

        // Reliability
        cronbachAlpha: cronbachAlpha,

        // Likert
        likertAnalysis: likertAnalysis,

        // Assumption checking
        checkAssumptions: checkAssumptions,
        leveneTest: leveneTest,

        // Epidemiology
        oddsRatioRR: oddsRatioRR,

        // Helpers
        formatPValue: formatPValue,
        formatCI: formatCI,
        significanceStar: significanceStar,
        interpretCohensD: interpretCohensD,
        interpretEtaSquared: interpretEtaSquared,
        interpretR: interpretR,
        interpretCramersV: interpretCramersV,
        interpretAlpha: interpretAlpha,
        ranks: ranks
    };

    // =========================================================================
    // Multiple Comparisons (Post-hoc tests)
    // =========================================================================

    Stats.tukeyHSD = function(groups, groupNames, alpha) {
        alpha = alpha || 0.05;
        if (!groups || groups.length < 2) return null;
        var k = groups.length;
        var allData = []; groups.forEach(function(g) { allData = allData.concat(g); });
        var N = allData.length;
        var dfW = N - k;
        var grandMean = allData.reduce(function(a,b){return a+b;},0) / N;
        var msW = 0;
        groups.forEach(function(g) {
            var gm = g.reduce(function(a,b){return a+b;},0) / g.length;
            g.forEach(function(v) { msW += (v - gm) * (v - gm); });
        });
        msW = msW / dfW;

        var pairs = [];
        for (var i = 0; i < k; i++) {
            for (var j = i + 1; j < k; j++) {
                var ni = groups[i].length, nj = groups[j].length;
                var mi = groups[i].reduce(function(a,b){return a+b;},0) / ni;
                var mj = groups[j].reduce(function(a,b){return a+b;},0) / nj;
                var diff = mi - mj;
                var se = Math.sqrt(msW * (1/ni + 1/nj) / 2);
                var q = Math.abs(diff) / se;
                // Approximate p-value using t-distribution
                var t = q / Math.sqrt(2);
                var p = 2 * (1 - jStat.studentt.cdf(Math.abs(t), dfW));
                // Bonferroni correction
                var pAdj = Math.min(p * (k * (k-1) / 2), 1);
                pairs.push({
                    groupA: groupNames[i], groupB: groupNames[j],
                    meanA: mi, meanB: mj, meanDiff: diff,
                    se: se, t: t, p: p, pAdjusted: pAdj,
                    significant: pAdj < alpha
                });
            }
        }
        return pairs;
    };

    Stats.bonferroni = function(groups, groupNames, alpha) {
        alpha = alpha || 0.05;
        if (!groups || groups.length < 2) return null;
        var k = groups.length;
        var nComparisons = k * (k - 1) / 2;
        var adjAlpha = alpha / nComparisons;
        var pairs = [];
        for (var i = 0; i < k; i++) {
            for (var j = i + 1; j < k; j++) {
                var result = Stats.independentTTest(groups[i], groups[j]);
                if (!result) continue;
                pairs.push({
                    groupA: groupNames[i], groupB: groupNames[j],
                    meanA: result.desc1.mean, meanB: result.desc2.mean,
                    meanDiff: result.meanDiff, t: result.t, p: result.p,
                    pAdjusted: Math.min(result.p * nComparisons, 1),
                    adjAlpha: adjAlpha,
                    significant: result.p < adjAlpha
                });
            }
        }
        return pairs;
    };

    Stats.scheffeTest = function(groups, groupNames, alpha) {
        alpha = alpha || 0.05;
        if (!groups || groups.length < 2) return null;
        var k = groups.length;
        var allData = []; groups.forEach(function(g) { allData = allData.concat(g); });
        var N = allData.length;
        var dfB = k - 1, dfW = N - k;
        var msW = 0;
        groups.forEach(function(g) {
            var gm = g.reduce(function(a,b){return a+b;},0) / g.length;
            g.forEach(function(v) { msW += (v - gm) * (v - gm); });
        });
        msW = msW / dfW;
        var fCrit = jStat.centralF.inv(1 - alpha, dfB, dfW);
        var pairs = [];
        for (var i = 0; i < k; i++) {
            for (var j = i + 1; j < k; j++) {
                var ni = groups[i].length, nj = groups[j].length;
                var mi = groups[i].reduce(function(a,b){return a+b;},0) / ni;
                var mj = groups[j].reduce(function(a,b){return a+b;},0) / nj;
                var diff = mi - mj;
                var fStat = (diff * diff) / (msW * (1/ni + 1/nj)) / dfB;
                var p = 1 - jStat.centralF.cdf(fStat * dfB, dfB, dfW);
                pairs.push({
                    groupA: groupNames[i], groupB: groupNames[j],
                    meanA: mi, meanB: mj, meanDiff: diff,
                    F: fStat, p: p, significant: p < alpha
                });
            }
        }
        return pairs;
    };

    // =========================================================================
    // Cross-tabulation Advanced
    // =========================================================================

    Stats.crossTab = function(var1, var2) {
        if (!var1 || !var2 || var1.length !== var2.length) return null;
        var n = var1.length;
        var labels1 = [], labels2 = [];
        var1.forEach(function(v) { if (labels1.indexOf(v) === -1) labels1.push(v); });
        var2.forEach(function(v) { if (labels2.indexOf(v) === -1) labels2.push(v); });
        labels1.sort(); labels2.sort();

        // Observed frequencies
        var observed = [];
        for (var i = 0; i < labels1.length; i++) {
            var row = [];
            for (var j = 0; j < labels2.length; j++) {
                var count = 0;
                for (var k = 0; k < n; k++) {
                    if (var1[k] === labels1[i] && var2[k] === labels2[j]) count++;
                }
                row.push(count);
            }
            observed.push(row);
        }

        // Row totals, col totals
        var rowTotals = observed.map(function(row) { return row.reduce(function(a,b){return a+b;},0); });
        var colTotals = [];
        for (var j = 0; j < labels2.length; j++) {
            var s = 0; for (var i = 0; i < labels1.length; i++) s += observed[i][j];
            colTotals.push(s);
        }

        // Expected
        var expected = [];
        for (var i = 0; i < labels1.length; i++) {
            var row = [];
            for (var j = 0; j < labels2.length; j++) {
                row.push(rowTotals[i] * colTotals[j] / n);
            }
            expected.push(row);
        }

        // Chi-square
        var chi2 = 0;
        for (var i = 0; i < labels1.length; i++) {
            for (var j = 0; j < labels2.length; j++) {
                if (expected[i][j] > 0) chi2 += Math.pow(observed[i][j] - expected[i][j], 2) / expected[i][j];
            }
        }
        var df = (labels1.length - 1) * (labels2.length - 1);
        var p = 1 - jStat.chisquare.cdf(chi2, df);
        var k = Math.min(labels1.length, labels2.length) - 1;
        var cramersV = k > 0 ? Math.sqrt(chi2 / (n * k)) : 0;
        var phi = Math.sqrt(chi2 / n);

        // Row percentages, Column percentages
        var rowPct = observed.map(function(row, i) {
            return row.map(function(c) { return rowTotals[i] > 0 ? c / rowTotals[i] * 100 : 0; });
        });
        var colPct = observed.map(function(row) {
            return row.map(function(c, j) { return colTotals[j] > 0 ? c / colTotals[j] * 100 : 0; });
        });

        // Standardized residuals
        var stdResiduals = [];
        for (var i = 0; i < labels1.length; i++) {
            var row = [];
            for (var j = 0; j < labels2.length; j++) {
                row.push(expected[i][j] > 0 ? (observed[i][j] - expected[i][j]) / Math.sqrt(expected[i][j]) : 0);
            }
            stdResiduals.push(row);
        }

        return {
            rowLabels: labels1, colLabels: labels2, n: n,
            observed: observed, expected: expected,
            rowTotals: rowTotals, colTotals: colTotals,
            rowPct: rowPct, colPct: colPct, stdResiduals: stdResiduals,
            chi2: chi2, df: df, p: p, cramersV: cramersV, phi: phi,
            interpretation: Stats.interpretCramersV(cramersV)
        };
    };

    // =========================================================================
    // Factor Analysis (EFA)
    // =========================================================================

    Stats.factorAnalysis = function(dataArrays, varNames) {
        // Simplified EFA using correlation matrix eigenvalue decomposition
        if (!dataArrays || dataArrays.length < 2) return null;
        var n = dataArrays[0].length;
        var p = dataArrays.length;
        var names = varNames || dataArrays.map(function(_,i) { return 'V'+(i+1); });

        // Correlation matrix
        var corrMatrix = [];
        for (var i = 0; i < p; i++) {
            var row = [];
            for (var j = 0; j < p; j++) {
                if (i === j) { row.push(1); continue; }
                var xi = dataArrays[i], xj = dataArrays[j];
                var mx = xi.reduce(function(a,b){return a+b;},0)/n;
                var my = xj.reduce(function(a,b){return a+b;},0)/n;
                var num = 0, dx = 0, dy = 0;
                for (var k = 0; k < n; k++) {
                    num += (xi[k]-mx)*(xj[k]-my);
                    dx += (xi[k]-mx)*(xi[k]-mx);
                    dy += (xj[k]-my)*(xj[k]-my);
                }
                row.push(dx>0 && dy>0 ? num / Math.sqrt(dx*dy) : 0);
            }
            corrMatrix.push(row);
        }

        // Power iteration for eigenvalues (simplified)
        var eigenvalues = [];
        var mat = corrMatrix.map(function(r) { return r.slice(); });
        for (var f = 0; f < Math.min(p, 10); f++) {
            var vec = [];
            for (var i = 0; i < p; i++) vec.push(Math.random());
            for (var iter = 0; iter < 100; iter++) {
                var newVec = [];
                for (var i = 0; i < p; i++) {
                    var s = 0;
                    for (var j = 0; j < p; j++) s += mat[i][j] * vec[j];
                    newVec.push(s);
                }
                var norm = Math.sqrt(newVec.reduce(function(a,b){return a+b*b;},0));
                if (norm === 0) break;
                vec = newVec.map(function(v) { return v / norm; });
            }
            var eigenval = 0;
            for (var i = 0; i < p; i++) {
                var s = 0;
                for (var j = 0; j < p; j++) s += mat[i][j] * vec[j];
                eigenval += s * vec[i];
            }
            eigenvalues.push(eigenval);
            // Deflate matrix
            for (var i = 0; i < p; i++) {
                for (var j = 0; j < p; j++) {
                    mat[i][j] -= eigenval * vec[i] * vec[j];
                }
            }
        }

        var totalVar = eigenvalues.reduce(function(a,b){return a+Math.max(b,0);},0);
        var cumVar = 0;
        var components = eigenvalues.map(function(ev, i) {
            var pctVar = totalVar > 0 ? ev / totalVar * 100 : 0;
            cumVar += pctVar;
            return { component: i+1, eigenvalue: ev, pctVariance: pctVar, cumPctVariance: cumVar };
        }).filter(function(c) { return c.eigenvalue > 0; });

        // Kaiser criterion: eigenvalue > 1
        var nFactors = components.filter(function(c) { return c.eigenvalue >= 1; }).length;

        // KMO approximation
        var sumR2 = 0, sumP2 = 0;
        for (var i = 0; i < p; i++) {
            for (var j = 0; j < p; j++) {
                if (i !== j) {
                    sumR2 += corrMatrix[i][j] * corrMatrix[i][j];
                    // Partial correlation approximation
                    sumP2 += corrMatrix[i][j] * corrMatrix[i][j] * 0.1;
                }
            }
        }
        var kmo = sumR2 / (sumR2 + sumP2);
        var kmoInterp = kmo >= 0.9 ? 'Marvelous' : kmo >= 0.8 ? 'Meritorious' : kmo >= 0.7 ? 'Middling' : kmo >= 0.6 ? 'Mediocre' : kmo >= 0.5 ? 'Miserable' : 'Unacceptable';

        return {
            corrMatrix: corrMatrix, varNames: names,
            components: components, nFactors: nFactors,
            kmo: kmo, kmoInterpretation: kmoInterp,
            totalVarianceExplained: components.slice(0, nFactors).reduce(function(a,c){return a+c.pctVariance;},0)
        };
    };

    // =========================================================================
    // Bootstrap Confidence Interval
    // =========================================================================

    Stats.bootstrapCI = function(values, statFn, nBoot, alpha) {
        nBoot = nBoot || 1000;
        alpha = alpha || 0.05;
        if (!values || values.length < 2) return null;
        var n = values.length;
        var bootStats = [];
        for (var b = 0; b < nBoot; b++) {
            var sample = [];
            for (var i = 0; i < n; i++) {
                sample.push(values[Math.floor(Math.random() * n)]);
            }
            bootStats.push(statFn(sample));
        }
        bootStats.sort(function(a,b){return a-b;});
        var loIdx = Math.floor(alpha / 2 * nBoot);
        var hiIdx = Math.floor((1 - alpha / 2) * nBoot);
        var original = statFn(values);
        return {
            estimate: original,
            ci95: Stats.formatCI(bootStats[loIdx], bootStats[hiIdx]),
            ci95_lo: bootStats[loIdx], ci95_hi: bootStats[hiIdx],
            se: Stats.descriptive(bootStats).sd,
            bias: Stats.descriptive(bootStats).mean - original,
            nBoot: nBoot
        };
    };

    // =========================================================================
    // Percentile Rank & Z-Score
    // =========================================================================

    Stats.percentileRank = function(values, score) {
        if (!values || values.length === 0) return null;
        var below = values.filter(function(v) { return v < score; }).length;
        var equal = values.filter(function(v) { return v === score; }).length;
        return (below + equal / 2) / values.length * 100;
    };

    Stats.zScores = function(values) {
        if (!values || values.length < 2) return null;
        var m = values.reduce(function(a,b){return a+b;},0) / values.length;
        var s = Math.sqrt(values.reduce(function(a,b){return a+(b-m)*(b-m);},0) / (values.length-1));
        return values.map(function(v) { return s > 0 ? (v - m) / s : 0; });
    };

    // =========================================================================
    // Multicollinearity (VIF)
    // =========================================================================

    Stats.vif = function(dataArrays, varNames) {
        if (!dataArrays || dataArrays.length < 2) return null;
        var p = dataArrays.length;
        var n = dataArrays[0].length;
        var names = varNames || dataArrays.map(function(_,i) { return 'X'+(i+1); });
        var results = [];

        for (var i = 0; i < p; i++) {
            // Regress X_i on all other X's
            var y = dataArrays[i];
            var xs = dataArrays.filter(function(_,j) { return j !== i; });
            var reg = Stats.linearRegression(y, xs, names.filter(function(_,j) { return j !== i; }));
            var rSquared = reg ? reg.rSquared : 0;
            var vifVal = 1 / (1 - rSquared);
            results.push({
                variable: names[i],
                rSquared: rSquared,
                vif: vifVal,
                tolerance: 1 - rSquared,
                status: vifVal > 10 ? 'Severe' : vifVal > 5 ? 'Moderate' : 'OK'
            });
        }
        return results;
    };

    // =========================================================================
    // Interpret functions for new stats
    // =========================================================================

    Stats.interpretKMO = function(kmo) {
        if (kmo >= 0.9) return 'Marvelous';
        if (kmo >= 0.8) return 'Meritorious';
        if (kmo >= 0.7) return 'Middling';
        if (kmo >= 0.6) return 'Mediocre';
        if (kmo >= 0.5) return 'Miserable';
        return 'Unacceptable';
    };

    // =========================================================================
    // Partial Correlation
    // =========================================================================

    Stats.partialCorrelation = function(x, y, controls) {
        // x, y = arrays of values; controls = array of arrays (control variables)
        if (!x || !y || x.length < 3) return null;
        var n = x.length;

        function residuals(dep, predictors) {
            var m = dep.reduce(function(a,b){return a+b;},0)/n;
            if (!predictors || predictors.length === 0) return dep.map(function(v){return v-m;});
            // Simple OLS residuals
            var k = predictors.length;
            // Use correlation-based residuals for simplicity
            var resid = dep.slice();
            predictors.forEach(function(pred) {
                var mp = pred.reduce(function(a,b){return a+b;},0)/n;
                var md = resid.reduce(function(a,b){return a+b;},0)/n;
                var num=0, den=0;
                for(var i=0;i<n;i++){num+=(pred[i]-mp)*(resid[i]-md); den+=(pred[i]-mp)*(pred[i]-mp);}
                var b = den>0 ? num/den : 0;
                var a = md - b*mp;
                resid = resid.map(function(v,i){return v - (a + b*pred[i]);});
            });
            return resid;
        }

        var rx = residuals(x, controls);
        var ry = residuals(y, controls);
        var mx=0,my=0; for(var i=0;i<n;i++){mx+=rx[i];my+=ry[i];} mx/=n;my/=n;
        var num=0,dx=0,dy=0;
        for(var i=0;i<n;i++){num+=(rx[i]-mx)*(ry[i]-my);dx+=(rx[i]-mx)*(rx[i]-mx);dy+=(ry[i]-my)*(ry[i]-my);}
        var r = (dx>0&&dy>0) ? num/Math.sqrt(dx*dy) : 0;
        var df = n - 2 - (controls ? controls.length : 0);
        var t = df>0 ? r*Math.sqrt(df/(1-r*r+1e-10)) : 0;
        var p = df>0 ? 2*(1-jStat.studentt.cdf(Math.abs(t),df)) : 1;
        return {r:r, t:t, df:df, p:p, rSquared:r*r};
    };

    // =========================================================================
    // ICC (Intraclass Correlation Coefficient)
    // =========================================================================

    Stats.icc = function(dataArrays) {
        // dataArrays = array of arrays, each array is one rater/measure
        if (!dataArrays || dataArrays.length < 2) return null;
        var k = dataArrays.length; // number of raters
        var n = dataArrays[0].length; // number of subjects

        var grandMean = 0, total = 0;
        dataArrays.forEach(function(arr){arr.forEach(function(v){grandMean+=v;total++;});});
        grandMean /= total;

        // Subject means
        var subjectMeans = [];
        for(var i=0;i<n;i++){
            var s=0; for(var j=0;j<k;j++) s+=dataArrays[j][i];
            subjectMeans.push(s/k);
        }

        // Rater means
        var raterMeans = dataArrays.map(function(arr){return arr.reduce(function(a,b){return a+b;},0)/n;});

        // SS Between subjects
        var SSB = 0;
        for(var i=0;i<n;i++) SSB += k*(subjectMeans[i]-grandMean)*(subjectMeans[i]-grandMean);

        // SS Within subjects
        var SSW = 0;
        for(var i=0;i<n;i++) for(var j=0;j<k;j++) SSW += (dataArrays[j][i]-subjectMeans[i])*(dataArrays[j][i]-subjectMeans[i]);

        // SS Raters
        var SSR = 0;
        for(var j=0;j<k;j++) SSR += n*(raterMeans[j]-grandMean)*(raterMeans[j]-grandMean);

        // SS Error
        var SSE = SSW - SSR;

        var MSB = SSB/(n-1);
        var MSW = SSW/(n*(k-1));
        var MSR = SSR/(k-1);
        var MSE = SSE/((n-1)*(k-1));

        // ICC(1,1) - one-way random
        var icc1 = (MSB-MSW)/(MSB+(k-1)*MSW);
        // ICC(2,1) - two-way random, single measures
        var icc2 = (MSB-MSE)/(MSB+(k-1)*MSE+k*(MSR-MSE)/n);
        // ICC(3,1) - two-way mixed, single measures
        var icc3 = (MSB-MSE)/(MSB+(k-1)*MSE);

        var interp = function(v){return v>=0.9?'Excellent':v>=0.75?'Good':v>=0.5?'Moderate':'Poor';};

        return {
            icc1:{value:icc1,interpretation:interp(icc1),type:'One-way random, single'},
            icc2:{value:icc2,interpretation:interp(icc2),type:'Two-way random, single'},
            icc3:{value:icc3,interpretation:interp(icc3),type:'Two-way mixed, single'},
            n:n, k:k, MSB:MSB, MSW:MSW, MSR:MSR, MSE:MSE
        };
    };

    // =========================================================================
    // Split-Half Reliability
    // =========================================================================

    Stats.splitHalf = function(dataArrays) {
        if (!dataArrays || dataArrays.length < 2) return null;
        var n = dataArrays[0].length;
        var k = dataArrays.length;
        // Split odd/even
        var odd=[], even=[];
        for(var i=0;i<n;i++){var so=0,se=0;for(var j=0;j<k;j++){if(j%2===0)se+=dataArrays[j][i];else so+=dataArrays[j][i];}odd.push(so);even.push(se);}
        // Correlation between halves
        var mo=odd.reduce(function(a,b){return a+b;},0)/n;
        var me=even.reduce(function(a,b){return a+b;},0)/n;
        var num=0,do2=0,de2=0;
        for(var i=0;i<n;i++){num+=(odd[i]-mo)*(even[i]-me);do2+=(odd[i]-mo)*(odd[i]-mo);de2+=(even[i]-me)*(even[i]-me);}
        var rHalf = (do2>0&&de2>0)?num/Math.sqrt(do2*de2):0;
        // Spearman-Brown
        var spearmanBrown = 2*rHalf/(1+Math.abs(rHalf));
        // Guttman split-half
        var varOdd=do2/(n-1), varEven=de2/(n-1);
        var totalScores = []; for(var i=0;i<n;i++){var s=0;for(var j=0;j<k;j++)s+=dataArrays[j][i];totalScores.push(s);}
        var mt=totalScores.reduce(function(a,b){return a+b;},0)/n;
        var varTotal=0;for(var i=0;i<n;i++)varTotal+=(totalScores[i]-mt)*(totalScores[i]-mt);varTotal/=(n-1);
        var guttman = varTotal>0 ? 2*(1-(varOdd+varEven)/varTotal) : 0;

        return {rHalf:rHalf, spearmanBrown:spearmanBrown, guttman:guttman, nItems:k, nCases:n};
    };

    // =========================================================================
    // Hierarchical Regression
    // =========================================================================

    Stats.hierarchicalRegression = function(y, blocks, varNames) {
        // blocks = [[iv1,iv2],[iv3,iv4]] each block is array of arrays
        // Returns model comparison for each step
        if (!y || !blocks || blocks.length === 0) return null;
        var n = y.length;
        var steps = [];
        var cumIVs = [];
        var cumNames = [];
        var prevR2 = 0;

        for (var step = 0; step < blocks.length; step++) {
            cumIVs = cumIVs.concat(blocks[step]);
            cumNames = cumNames.concat(varNames[step] || blocks[step].map(function(_,i){return 'X'+(cumIVs.length-blocks[step].length+i+1);}));

            var result = Stats.linearRegression(y, cumIVs, cumNames);
            if (!result) continue;

            var r2Change = result.rSquared - prevR2;
            var dfChange = blocks[step].length;
            var df2 = n - cumIVs.length - 1;
            var fChange = df2>0 ? (r2Change/dfChange)/((1-result.rSquared)/df2) : 0;
            var pChange = fChange>0 ? 1-jStat.centralF.cdf(fChange, dfChange, df2) : 1;

            steps.push({
                step: step+1,
                r: result.r, rSquared: result.rSquared, adjRSquared: result.adjRSquared,
                r2Change: r2Change, fChange: fChange, df1Change: dfChange, df2Change: df2, pChange: pChange,
                f: result.f, fP: result.fP,
                coefficients: result.coefficients,
                varsAdded: varNames[step] || []
            });
            prevR2 = result.rSquared;
        }
        return steps;
    };

    // =========================================================================
    // ROC Curve / AUC
    // =========================================================================

    Stats.roc = function(actual, predicted) {
        // actual = array of 0/1, predicted = array of probabilities
        if (!actual || !predicted || actual.length !== predicted.length) return null;
        var n = actual.length;

        // Sort by predicted descending
        var pairs = [];
        for(var i=0;i<n;i++) pairs.push({actual:actual[i], pred:predicted[i]});
        pairs.sort(function(a,b){return b.pred-a.pred;});

        var nPos = actual.filter(function(v){return v===1;}).length;
        var nNeg = n - nPos;
        if(nPos===0||nNeg===0) return null;

        var points = [{fpr:0, tpr:0}];
        var tp=0, fp=0;

        for(var i=0;i<n;i++){
            if(pairs[i].actual===1) tp++; else fp++;
            points.push({fpr:fp/nNeg, tpr:tp/nPos, threshold:pairs[i].pred});
        }

        // AUC using trapezoidal rule
        var auc = 0;
        for(var i=1;i<points.length;i++){
            auc += (points[i].fpr-points[i-1].fpr)*(points[i].tpr+points[i-1].tpr)/2;
        }

        // Youden's J - optimal threshold
        var bestJ = -1, bestThreshold = 0.5;
        points.forEach(function(pt){
            var j = pt.tpr - pt.fpr;
            if(j>bestJ){bestJ=j;bestThreshold=pt.threshold||0.5;}
        });

        // Sensitivity/Specificity at optimal threshold
        var optTP=0,optFP=0,optTN=0,optFN=0;
        for(var i=0;i<n;i++){
            var pred = predicted[i]>=bestThreshold?1:0;
            if(pred===1&&actual[i]===1)optTP++;
            if(pred===1&&actual[i]===0)optFP++;
            if(pred===0&&actual[i]===0)optTN++;
            if(pred===0&&actual[i]===1)optFN++;
        }

        var interp = auc>=0.9?'Outstanding':auc>=0.8?'Excellent':auc>=0.7?'Acceptable':auc>=0.6?'Poor':'Fail';

        return {
            auc:auc, interpretation:interp, points:points,
            optimalThreshold:bestThreshold, youdenJ:bestJ,
            sensitivity:nPos>0?optTP/nPos:0, specificity:nNeg>0?optTN/nNeg:0,
            ppv:(optTP+optFP)>0?optTP/(optTP+optFP):0,
            npv:(optTN+optFN)>0?optTN/(optTN+optFN):0,
            accuracy:(optTP+optTN)/n
        };
    };

    // =========================================================================
    // McNemar Test
    // =========================================================================

    Stats.mcnemar = function(before, after) {
        if(!before||!after||before.length!==after.length) return null;
        var n=before.length;
        var a=0,b=0,c=0,d=0;
        for(var i=0;i<n;i++){
            if(before[i]===1&&after[i]===1)a++;
            else if(before[i]===1&&after[i]===0)b++;
            else if(before[i]===0&&after[i]===1)c++;
            else d++;
        }
        var chi2=(b+c)>0?Math.pow(Math.abs(b-c)-1,2)/(b+c):0; // continuity correction
        var p=1-jStat.chisquare.cdf(chi2,1);
        return {a:a,b:b,c:c,d:d,chi2:chi2,df:1,p:p,n:n,
                discordant:b+c,significant:p<0.05};
    };

    // =========================================================================
    // Fisher's Exact Test (2x2)
    // =========================================================================

    Stats.fisherExact = function(a,b,c,d) {
        var n=a+b+c+d;
        function logFact(x){var s=0;for(var i=2;i<=x;i++)s+=Math.log(i);return s;}
        var logP = logFact(a+b)+logFact(c+d)+logFact(a+c)+logFact(b+d)-logFact(n)-logFact(a)-logFact(b)-logFact(c)-logFact(d);
        var pExact = Math.exp(logP);
        // Two-tailed: sum probabilities <= pExact
        var r1=a+b,r2=c+d,c1=a+c,c2=b+d;
        var pTwoTail=0;
        for(var aa=0;aa<=Math.min(r1,c1);aa++){
            var bb=r1-aa,cc=c1-aa,dd=r2-cc;
            if(bb<0||cc<0||dd<0)continue;
            var lp=logFact(r1)+logFact(r2)+logFact(c1)+logFact(c2)-logFact(n)-logFact(aa)-logFact(bb)-logFact(cc)-logFact(dd);
            var pp=Math.exp(lp);
            if(pp<=pExact+1e-10) pTwoTail+=pp;
        }
        var or=(b*c)>0?(a*d)/(b*c):Infinity;
        return {pExact:pExact,pTwoTail:Math.min(pTwoTail,1),oddsRatio:or,a:a,b:b,c:c,d:d,n:n};
    };

    // =========================================================================
    // Cochran's Q Test
    // =========================================================================

    Stats.cochranQ = function(dataArrays) {
        // dataArrays = array of arrays (each is 0/1 for each condition)
        if(!dataArrays||dataArrays.length<3) return null;
        var k=dataArrays.length, n=dataArrays[0].length;
        var T=0; // Grand total
        var Tj=[]; // Column totals
        var Li=[]; // Row totals
        for(var j=0;j<k;j++){var s=0;for(var i=0;i<n;i++)s+=dataArrays[j][i];Tj.push(s);T+=s;}
        for(var i=0;i<n;i++){var s=0;for(var j=0;j<k;j++)s+=dataArrays[j][i];Li.push(s);}
        var sumTj2=Tj.reduce(function(a,b){return a+b*b;},0);
        var sumLi2=Li.reduce(function(a,b){return a+b*b;},0);
        var sumLi=Li.reduce(function(a,b){return a+b;},0);
        var Q=((k-1)*(k*sumTj2-T*T))/(k*sumLi-sumLi2);
        if(isNaN(Q))Q=0;
        var df=k-1;
        var p=1-jStat.chisquare.cdf(Q,df);
        return {Q:Q,df:df,p:p,k:k,n:n,significant:p<0.05};
    };

    // =========================================================================
    // Games-Howell Post-hoc
    // =========================================================================

    Stats.gamesHowell = function(groups, groupNames) {
        if(!groups||groups.length<2) return null;
        var k=groups.length;
        var pairs=[];
        for(var i=0;i<k;i++){
            for(var j=i+1;j<k;j++){
                var ni=groups[i].length, nj=groups[j].length;
                var mi=groups[i].reduce(function(a,b){return a+b;},0)/ni;
                var mj=groups[j].reduce(function(a,b){return a+b;},0)/nj;
                var vi=groups[i].reduce(function(a,b){return a+(b-mi)*(b-mi);},0)/(ni-1);
                var vj=groups[j].reduce(function(a,b){return a+(b-mj)*(b-mj);},0)/(nj-1);
                var se=Math.sqrt(vi/ni+vj/nj);
                var t=se>0?(mi-mj)/se:0;
                // Welch-Satterthwaite df
                var num=Math.pow(vi/ni+vj/nj,2);
                var den=Math.pow(vi/ni,2)/(ni-1)+Math.pow(vj/nj,2)/(nj-1);
                var df=den>0?num/den:1;
                var p=2*(1-jStat.studentt.cdf(Math.abs(t),df));
                pairs.push({groupA:groupNames[i],groupB:groupNames[j],meanA:mi,meanB:mj,
                            meanDiff:mi-mj,se:se,t:t,df:df,p:p,significant:p<0.05});
            }
        }
        return pairs;
    };

    // =========================================================================
    // LSD (Fisher's Least Significant Difference)
    // =========================================================================

    Stats.fisherLSD = function(groups, groupNames, alpha) {
        alpha=alpha||0.05;
        if(!groups||groups.length<2) return null;
        var k=groups.length;
        var allData=[];groups.forEach(function(g){allData=allData.concat(g);});
        var N=allData.length, dfW=N-k;
        var msW=0;
        groups.forEach(function(g){var m=g.reduce(function(a,b){return a+b;},0)/g.length;g.forEach(function(v){msW+=(v-m)*(v-m);});});
        msW=msW/dfW;
        var tCrit=jStat.studentt.inv(1-alpha/2,dfW);
        var pairs=[];
        for(var i=0;i<k;i++){
            for(var j=i+1;j<k;j++){
                var ni=groups[i].length,nj=groups[j].length;
                var mi=groups[i].reduce(function(a,b){return a+b;},0)/ni;
                var mj=groups[j].reduce(function(a,b){return a+b;},0)/nj;
                var se=Math.sqrt(msW*(1/ni+1/nj));
                var t=(mi-mj)/se;
                var p=2*(1-jStat.studentt.cdf(Math.abs(t),dfW));
                var lsd=tCrit*se;
                pairs.push({groupA:groupNames[i],groupB:groupNames[j],meanA:mi,meanB:mj,
                            meanDiff:mi-mj,se:se,t:t,df:dfW,p:p,lsd:lsd,significant:Math.abs(mi-mj)>lsd});
            }
        }
        return pairs;
    };

    // =========================================================================
    // Omega Squared & Partial Eta Squared
    // =========================================================================

    Stats.omegaSquared = function(ssBetween, ssTotal, dfBetween, msWithin) {
        return (ssBetween - dfBetween*msWithin) / (ssTotal + msWithin);
    };

    Stats.partialEtaSquared = function(ssBetween, ssError) {
        return ssBetween / (ssBetween + ssError);
    };

})();
