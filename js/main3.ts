import chroma from "chroma-js";
import dayjs from "dayjs";
import duration from "dayjs/plugin/duration";
import quarterOfYear from "dayjs/plugin/quarterOfYear";
// Seasonality
dayjs.extend(duration);
dayjs.extend(quarterOfYear);

import { init as initEChart } from "echarts";
import type { EChartsType } from "echarts";
import Handsontable from "handsontable";
import { CellMeta } from "handsontable/settings";

import lz77 from "./lz77";


/** 
 * Typescript type and interface definitions
 */

declare global {
  function closeCookieBanner(choice: string): void;
  function redrawEcharts(): void;
}

type DataValue = {
  order: number;
  x: string; // Date in YYYY-MM-DD format
  value: number;
  status: DataStatus;
  seasonalFactor?: number;
};

enum DataStatus {
  NORMAL = 0,
  RUN_OF_EIGHT_EXCEPTION = 1, // 8 on one side
  FOUR_NEAR_LIMIT_EXCEPTION = 2, // 3 out of 4 on the extreme quarters
  NPL_EXCEPTION = 3, // out of the NPL limit lines.
}

function dataStatusColor(d: DataStatus): string {
  switch (d) {
    case DataStatus.RUN_OF_EIGHT_EXCEPTION:
      return "blue";
    case DataStatus.FOUR_NEAR_LIMIT_EXCEPTION:
      return "orange";
    case DataStatus.NPL_EXCEPTION:
      return "red";
    default:
      return "black";
  }
}

function dataLabelsStatusColor(d: DataStatus): string {
  switch (d) {
    case DataStatus.RUN_OF_EIGHT_EXCEPTION:
      return "blue";
    case DataStatus.FOUR_NEAR_LIMIT_EXCEPTION:
      return "#be5504"; // ginger. contrast: 4.68
    case DataStatus.NPL_EXCEPTION:
      return "#e3242b"; // rose. contrast 4.61
    default:
      return "black";
  }
}

// We use bits to represent locked limit status
// If we modify unpl and lnpl and locked, the locked limit status will be 0111
enum LockedLimitStatus {
  UNLOCKED = 0,
  LOCKED = 1,
  UNPL_MODIFIED = 2, // 0010
  LNPL_MODIFIED = 4, // 0100
  AVGX_MODIFIED = 8, // 1000
}

function isAvgXModified(s: LockedLimitStatus): boolean {
  return (
    (s & LockedLimitStatus.AVGX_MODIFIED) == LockedLimitStatus.AVGX_MODIFIED
  );
}

function isUnplModified(s: LockedLimitStatus): boolean {
  return (
    (s & LockedLimitStatus.UNPL_MODIFIED) == LockedLimitStatus.UNPL_MODIFIED
  );
}

function isLnplModified(s: LockedLimitStatus): boolean {
  return (
    (s & LockedLimitStatus.LNPL_MODIFIED) == LockedLimitStatus.LNPL_MODIFIED
  );
}

// Divider Type is the backing type for the divider line.
interface DividerType {
  id: string;
  x: number;
}

// This contains the state for a segmented section of both graphs (after
// a divider has been inserted)
type LineValueType = {
  xLeft: number; // date in unix milliseconds
  xRight: number; // date in unix milliseconds
  avgX: number;
  avgMovement?: number;
  UNPL?: number; // upper natural process limit
  LNPL?: number; // lower natural process limit
  URL?: number; // upper range limit
  lowerQuartile?: number;
  upperQuartile?: number;
};

type _Stats = {
  xchartMin: number;
  xchartMax: number;
  mrchartMax: number;
  lineValues: LineValueType[];
  xdataPerRange: DataValue[][];
  movementsPerRange: DataValue[][];
};

type SeasonalityPeriod = "year" | "quarter" | "month" | "week";
type SeasonalityGrouping = "week" | "month" | "quarter";

/**
 * Global constants and variables
 */
const MAX_LINK_LENGTH = 2000; // 2000 characters in url: https://stackoverflow.com/questions/417142/what-is-the-maximum-length-of-a-url-in-different-browsers
const NPL_SCALING = 2.66;
const URL_SCALING = 3.268;
const USE_MEDIAN_AVG = false;
const USE_MEDIAN_MR = false;
const MEDIAN_NPL_SCALING = 3.145;
const MEDIAN_URL_SCALING = 3.865;
// PADDING_FROM_EXTREMES is the percentage from the data limits (max/min) that we use as chart limits.
const PADDING_FROM_EXTREMES = 0.1;
const DECIMAL_POINT = 2;
const LIMIT_LINE_WIDTH = 2;
const DIVIDER_LINE_WIDTH = 4;
const INACTIVE_LOCKED_LIMITS = {
  avgX: 0,
  LNPL: Infinity,
  UNPL: -Infinity,
  avgMovement: 0,
  URL: -Infinity,
} as LineValueType;
const MEAN_SHAPE_COLOR = "red";
const LIMIT_SHAPE_COLOR = "steelblue";
const INITIAL_VALUES = [
  5045, 4350, 4350, 3975, 4290, 4430, 4485, 4285, 3980, 3925, 3645, 3760, 3300,
  3685, 3463, 5200,
];
const DUMMY_DATA: DataValue[] = INITIAL_VALUES.map(function (el, i) {
  return {
    order: i,
    x: `2020-01-${(i < 9 ? "0" : "") + (i + 1)}`,
    value: el,
    status: DataStatus.NORMAL,
  };
});


// This is the state of the app.
const state = {
  // tableData, xLabel and yLabel is set by the handsontable (the table rendered onscreen).
  // for handsontable to be reactive, use hot.updateSettings({ ... }). Don't just use assignment.
  tableData: [] as DataValue[], // handsontable data. Might contain null values.
  lockedLimitBaseData: [] as DataValue[], // data for the locked limit lines
  xLabel: "Date",
  yLabel: "Value",

  // xdata, movements are bound to the charts.
  // So, all update must be done in place (e.g. use updateInPlace) to maintain reactiveness.
  xdata: [] as DataValue[], // data passed to plot(...). Will not contain null values.
  movements: [] as DataValue[],
  dividerLines: [] as DividerType[],
  // this is the state of the locked limit lines, either calculated from lockedLimitBaseData or overwritten by user.
  // Only written when the user press "Lock limits" in the popup.
  // Locked limits should be considered active if some values are non-zero.
  lockedLimits: structuredClone(INACTIVE_LOCKED_LIMITS),
  lockedLimitStatus: LockedLimitStatus.UNLOCKED,

  // De-sesonalisation
  seasonalFactorData: [] as DataValue[], // data to generate seasonal factors from
  seasonalFactorTableData: [] as number[][], // seasonal factors to apply
  deSeasonalisePeriod: "year" as SeasonalityPeriod, // de-seasonalisation period
  deSeasonaliseGrouping: "none" as "none" | SeasonalityGrouping, //de-seasonalisation grouping
  deSeasonalisedData: [] as DataValue[], // data after applying seasonal factors
  isShowingDeseasonalisedData: false, // flag that toggles between normal data and deseasonalised data
  hasMissingPeriods: false,

  // Trends
  trendData: [] as DataValue[],
  trendLines: {
    centreLine: [],
    unpl: [],
    reducedUnpl: [],
    lnpl: [],
    reducedLnpl: [],
    lowerQtl: [],
    upperQtl: [],
  } as { [line: string]: DataValue[] },
  regressionStats: { m: 0, c: 0, avgMR: 0 } as RegressionStats,
  isShowingTrendLines: false,
};

type AppState = typeof state;

let xplot: EChartsType;
if (document.getElementById("xplot")) {
  xplot = initEChart(document.getElementById("xplot") as HTMLDivElement);
  xplot.on("click", (params) => {
    if (params.componentType === "title") {
      promptNewColumnName();
    }
  });
}
let mrplot: EChartsType;
if (document.getElementById("mrplot")) {
  mrplot = initEChart(document.getElementById("mrplot") as HTMLDivElement);
  mrplot.on("click", (params) => {
    if (params.componentType === "title") {
      promptNewColumnName();
    }
  });
}

// Limit checks
function checkRunOfEight(data: DataValue[], avg: number | DataValue[]) {
  if (typeof avg !== "number" && avg.length < 1) return;

  const avgValue = (i) => (typeof avg !== "number" ? avg[i].value : avg);

  let isApplied = false;
  if (data.length < 8) {
    return;
  }
  let aboveOrBelow = 0; // We use an 8-bit integer. Bit is set to 0 if below avg, or 1 otherwise
  for (let i = 0; i < 7; i++) {
    if (data[i].value > avgValue(i)) {
      aboveOrBelow |= 1 << i % 8;
    }
  }
  for (let i = 7; i < data.length; i++) {
    if (data[i].value > avgValue(i)) {
      // set bit to 1
      aboveOrBelow |= 1 << i % 8;
    } else {
      // set bit to 0
      aboveOrBelow &= ~(1 << i % 8);
    }
    if (aboveOrBelow == 0 || aboveOrBelow == 255) {
      for (let j = i - 7; j <= i; j++) {
        data[j].status = DataStatus.RUN_OF_EIGHT_EXCEPTION;
      }
    }
  }
  return;
}

/**
 * Detection Checks for quarter lines. If 3 out of 4 consecutive values are above or below a quarter line, it will be marked
 * @param data
 * @param lowerQuartile Pass in Infinity to disable the use of the lower quartile of the data for checks
 * @param upperQuartile Pass in -Infinity to disable the use of the upper quartile of the data for checks
 */
function checkFourNearLimit(
  data: DataValue[],
  lowerQuartile: number | DataValue[],
  upperQuartile: number | DataValue[]
) {
  if (typeof lowerQuartile !== "number" && lowerQuartile.length < 1) return;
  if (typeof upperQuartile !== "number" && upperQuartile.length < 1) return;

  const upperQuartileValue = (i) =>
    typeof upperQuartile !== "number" ? upperQuartile[i].value : upperQuartile;
  const lowerQuartileValue = (i) =>
    typeof lowerQuartile !== "number" ? lowerQuartile[i].value : lowerQuartile;

  if (data.length < 4) {
    return;
  }

  let belowQuartile = 0;
  let aboveQuartile = 0;
  // setup sliding window
  for (let i = 0; i < 3; i++) {
    if (data[i].value < lowerQuartileValue(i)) {
      belowQuartile += 1;
    } else if (data[i].value > upperQuartileValue(i)) {
      aboveQuartile += 1;
    }
  }

  for (let i = 3; i < data.length; i++) {
    // set value for the current window
    if (data[i].value < lowerQuartileValue(i)) {
      belowQuartile += 1;
    } else if (data[i].value > upperQuartileValue(i)) {
      aboveQuartile += 1;
    }

    if (belowQuartile >= 3 || aboveQuartile >= 3) {
      for (let j = i - 3; j <= i; j++) {
        data[j].status = DataStatus.FOUR_NEAR_LIMIT_EXCEPTION;
      }
    }

    // reset value to prepare for next window
    if (data[i - 3].value < lowerQuartileValue(i - 3)) {
      belowQuartile -= 1;
    } else if (data[i - 3].value > upperQuartileValue(i - 3)) {
      aboveQuartile -= 1;
    }
  }
}

function checkOutsideLimit(
  data: DataValue[],
  lowerLimit: number | DataValue[],
  upperLimit: number | DataValue[]
) {
  if (typeof lowerLimit !== "number" && lowerLimit.length < 1) return;
  if (typeof upperLimit !== "number" && upperLimit.length < 1) return;

  const upperLimitValue = (i) =>
    typeof upperLimit !== "number" ? upperLimit[i].value : upperLimit;
  const lowerLimitValue = (i) =>
    typeof lowerLimit !== "number" ? lowerLimit[i].value : lowerLimit;

  data.forEach((dv, i) => {
    if (dv.value < lowerLimitValue(i)) {
      dv.status = DataStatus.NPL_EXCEPTION;
    } else if (dv.value > upperLimitValue(i)) {
      dv.status = DataStatus.NPL_EXCEPTION;
    }
  });
}

function sum(ns: number[]) {
  return ns.reduce((t, n) => t + n);
}
function average(ns: number[]) {
  return ns.reduce((t, n) => t + n, 0) / ns.length;
}

function determinePeriodicity(xdata: DataValue[]) {
  // Get deltas (as number of days)
  const xValues = xdata.map((d) => dayjs(d.x));
  let deltas = [];
  for (let i = 1; i < xValues.length; i++) {
    deltas.push(xValues[i].diff(xValues[i - 1], "days"));
  }
  // Count frequency of delta
  const diffCounts = deltas.reduce((acc, diff) => {
    acc[diff] = (acc[diff] || 0) + 1;
    return acc;
  }, {});
  // Determine most common interval
  const mostCommonDiff = Object.keys(diffCounts).reduce((a, b) =>
    diffCounts[a] > diffCounts[b] ? a : b
  );
  // Map intervals to possible periodicities
  const interval = parseFloat(mostCommonDiff);
  if (interval < 7) {
    return "day";
  } else if (interval < 28) {
    return "week";
  } else if (interval < 90) {
    return "month";
  } else if (interval < 365) {
    return "quarter";
  } else {
    return "year";
  }
}

function periodiseData(
  initialX: string,
  xdata: DataValue[],
  period: SeasonalityPeriod = "year"
): DataValue[][] {
  const dataInterval = determinePeriodicity(xdata);

  if (dataInterval === "unknown") {
    console.error("Data has an irregular interval");
    return [];
  }

  const lastDate = dayjs(xdata[xdata.length - 1].x);

  const xDataMap = xdata.reduce((acc, d) => {
    acc[dayjs(d.x).toISOString()] = d;
    return acc;
  }, {});

  const periodisedData: DataValue[][] = [];

  let periodStart = dayjs(initialX).startOf(period).startOf(dataInterval);
  let periodEnd = dayjs(initialX).endOf(period).endOf(dataInterval);
  let periodDuration = periodEnd.diff(periodStart, dataInterval);

  let d = periodStart;
  while (d.isBefore(lastDate.add(1, "day"))) {
    const currPeriod = [];
    for (let i = 0; i <= periodDuration; i++) {
      d = periodStart.add(i, dataInterval);
      if (!d.isBefore(lastDate.add(1, "day"))) {
        continue;
      }
      const dataPoint = xDataMap[d.toISOString()] ?? null;
      currPeriod.push(dataPoint);
    }

    periodisedData.push(currPeriod);

    // Move to the next period and seek to the start/end of the data interval
    periodStart = periodStart.add(1, period).startOf(dataInterval);
    periodEnd = periodEnd.add(1, period).endOf(dataInterval);
    periodDuration = periodEnd.diff(periodStart, dataInterval);
  }

  return periodisedData;
}

function periodiseDataGrouped(
  initialX: string,
  xdata: DataValue[],
  period: SeasonalityPeriod = "year",
  grouping: SeasonalityGrouping
): DataValue[][] {
  // define auxiliary variables for tracking period and sub-period
  let periodEnd = dayjs(initialX).endOf(period);
  let subPeriodEnd = dayjs(initialX).endOf(grouping);

  if (subPeriodEnd.isAfter(periodEnd)) {
    throw "Invalid parameters: sub-period duration must be more than period duration!";
  }

  let currSubPeriodValues: DataValue[] = [];
  const periodisedData: DataValue[][] = [[]];

  xdata.forEach((d) => {
    const date = dayjs(d.x);
    if (date.isValid() && !date.isBefore(subPeriodEnd)) {
      periodisedData[0].unshift(
        currSubPeriodValues.reduce(
          (acc, x) => ({ ...acc, value: acc.value + x.value }),
          {
            order: periodisedData.length,
            value: 0,
            status: DataStatus.NORMAL,
            x: toDateStr(subPeriodEnd.startOf(grouping).toDate()),
          }
        )
      );
      // move to the next sub-period
      subPeriodEnd = subPeriodEnd.add(1, grouping).endOf(grouping);
      currSubPeriodValues = [];
    }
    if (date.isValid() && !date.isBefore(periodEnd)) {
      // move to the next period
      periodEnd = periodEnd.add(1, period).endOf(period);
      periodisedData[0].reverse();
      periodisedData.unshift([]);
    }
    // add current data value to the current sub-period
    currSubPeriodValues.push(d);
  });

  periodisedData[0].unshift(
    currSubPeriodValues.reduce(
      (acc, x) => ({ ...acc, value: acc.value + x.value }),
      {
        order: periodisedData.length,
        value: 0,
        status: DataStatus.NORMAL,
        x: toDateStr(subPeriodEnd.startOf(grouping).toDate()),
      }
    )
  );
  periodisedData[0].reverse();
  periodisedData.reverse();

  return periodisedData;
}

function calculateSeasonalFactors(
  xData: DataValue[],
  seasonalData: DataValue[],
  period: SeasonalityPeriod = "year",
  grouping: SeasonalityGrouping | "none" = "none"
) {
  const isGrouped = grouping !== "none";

  // periodise data
  let periodisedData: DataValue[][];
  if (isGrouped) {
    periodisedData = periodiseDataGrouped(
      xData[0].x,
      seasonalData,
      period,
      grouping
    );
  } else {
    periodisedData = periodiseData(xData[0].x, seasonalData, period);
  }

  if (!periodisedData.every((p) => p.length === periodisedData[0].length)) {
    console.warn("Seasons have varying number of points");
  }

  let hasMissingSubPeriods = false;
  hasMissingSubPeriods =
    !periodisedData.every((p) => p.length === periodisedData[0].length) &&
    isGrouped;

  const subPeriodCount = Math.max(...periodisedData.map((p) => p.length));

  const subPeriodAggregates: number[] = [];
  const aggregationStrategy = isGrouped ? sum : average;

  for (let i = 0; i < subPeriodCount; i++) {
    subPeriodAggregates.push(
      aggregationStrategy(
        periodisedData
          .map((p) => p[i]) // get subperiod values
          .filter((v) => v != null) // ignore periods with missing subperiods
          .map((d) => d.value) // map to raw values
      )
    );
  }

  // Derive overall average and calculate seasonal factors
  const overallAvg = average(
    isGrouped
      ? subPeriodAggregates
      : seasonalData.map((d) => d.value).filter((v) => v != null)
  );

  const seasonalFactors = subPeriodAggregates.map((v) =>
    isNaN(v) ? 1 : v / overallAvg
  );

  return [seasonalFactors, hasMissingSubPeriods] as const;
}

function applySeasonalFactorsGrouped(
  periodisedData: DataValue[][],
  seasonalFactors: number[],
  grouping: SeasonalityGrouping
) {
  console.assert(periodisedData.length > 0);
  let subPeriodEnd = dayjs(periodisedData[0][0].x).endOf(grouping);

  let newXData: DataValue[] = [];
  let currSubPeriodSum = 0;

  periodisedData.forEach((period) => {
    period.forEach((d, i) => {
      const date = dayjs(d.x);
      if (!date.isBefore(subPeriodEnd)) {
        // add new data value
        console.assert(i < seasonalFactors.length);

        newXData.push({
          value: parseFloat((currSubPeriodSum / seasonalFactors[i]).toFixed(1)),
          x: toDateStr(subPeriodEnd.startOf(grouping).toDate()),
          order: newXData.length,
          status: DataStatus.NORMAL,
        });
        currSubPeriodSum = 0;
        // move to the next sub-period
        subPeriodEnd = subPeriodEnd.add(1, grouping).endOf(grouping);
      }
      currSubPeriodSum += d.value;
    });
  });

  return newXData;
}

function applySeasonalFactors(
  xData: DataValue[],
  seasonalFactors: number[],
  grouping: SeasonalityGrouping | "none",
  period: SeasonalityPeriod = "year"
) {
  const periodisedData = periodiseData(xData[0].x, xData, period);

  const isGrouped = grouping !== "none";
  if (isGrouped) {
    return applySeasonalFactorsGrouped(
      periodisedData,
      seasonalFactors,
      grouping as SeasonalityGrouping
    );
  }

  // build a map of date values to their SFs
  const dateSfMap = {};
  periodisedData.forEach((period) => {
    for (let i = 0; i < period.length; i++) {
      const d = period[i];
      if (d === null) continue;
      dateSfMap[dayjs(d.x).toISOString()] = [i, seasonalFactors[i]];
    }
  });

  // build a set of all dates included in the periodised data
  const periodSet = new Set(
    periodisedData
      .reduce((acc, p) => acc.concat(p), [])
      .filter((d) => d !== null)
      .map((d) => dayjs(d.x).toISOString())
  );

  return deepClone(xData).map((d) => {
    const isoDate = dayjs(d.x).toISOString();
    if (periodSet.has(isoDate)) {
      // apply seasonal factor
      const [sfIndex, sf] = dateSfMap[isoDate];
      if (sf !== 0) d.value /= sf;
      // set sf index
      d.seasonalFactor = sfIndex + 1;
    }
    return d;
  });
}

// Trends
type RegressionStats = {
  m: number;
  c: number;
  avgMR: number;
};

function calculateRegressionStats(data: DataValue[]): RegressionStats {
  const { m, c } = linearRegression(data);
  const mr = getMovements(data).map((d) => d.value);
  const avgMR = average(mr);
  return {
    m,
    c,
    avgMR,
  };
}

function createTrendlines(
  { m, c, avgMR }: RegressionStats,
  dataValues: DataValue[]
) {
  let centreLine: DataValue[] = [];
  let unpl: DataValue[] = [];
  let lnpl: DataValue[] = [];
  let lowerQtl: DataValue[] = [];
  let upperQtl: DataValue[] = [];

  let reducedUnpl: DataValue[] = [];
  let reducedLnpl: DataValue[] = [];

  dataValues.forEach((d, i) => {
    const centreLineValue = i * m + c;
    const unplValue = centreLineValue + avgMR * NPL_SCALING;
    const lnplValue = centreLineValue - avgMR * NPL_SCALING;
    const lowerQtlValue = (lnplValue + centreLineValue) / 2;
    const upperQtlValue = (unplValue + centreLineValue) / 2;

    const reducedUnplValue = centreLineValue + (avgMR - m) * NPL_SCALING;
    const reducedLnplValue = centreLineValue - (avgMR - m) * NPL_SCALING;

    const createDataValue = (value: number) => ({
      order: i,
      x: d.x,
      value: round(value),
      status: DataStatus.NORMAL,
    });

    centreLine.push(createDataValue(centreLineValue));
    unpl.push(createDataValue(unplValue));
    lnpl.push(createDataValue(lnplValue));
    lowerQtl.push(createDataValue(lowerQtlValue));
    upperQtl.push(createDataValue(upperQtlValue));

    reducedUnpl.push(createDataValue(reducedUnplValue));
    reducedLnpl.push(createDataValue(reducedLnplValue));
  });

  return {
    centreLine,
    unpl,
    lnpl,
    lowerQtl,
    upperQtl,
    reducedUnpl,
    reducedLnpl,
  };
}

function linearRegression(yValues: DataValue[]): { m: number; c: number } | null {
  if (yValues.length < 2) {
    console.error(
      "At least two data points are required for linear regression."
    );
    return null;
  }

  // normalize data
  let validXData = removeAllNulls(state.xdata)
  let firstData = fromDateStr(validXData[0].x)
  let base = fromDateStr(validXData[1].x) - firstData;
  let normalizedValue: { x: number, y: number }[] = yValues.map(d => {
    return {
      x: (fromDateStr(d.x) - firstData) / base,
      y: d.value,
    }
  })

  const n = yValues.length;

  let sumX = 0;
  let sumY = 0;
  let sumXY = 0;
  let sumX2 = 0;

  for (let i = 0; i < n; i++) {
    const x = normalizedValue[i].x; // 1-indexed x values
    const y = normalizedValue[i].y;

    sumX += x;
    sumY += y;
    sumXY += x * y;
    sumX2 += x * x;
  }

  const numerator = n * sumXY - sumX * sumY;
  const denominator = n * sumX2 - sumX * sumX;

  if (denominator === 0) {
    console.error(
      "Denominator is zero; check the data points for possible vertical alignment."
    );
    return null;
  }

  const m = numerator / denominator;
  const c = (sumY - m * sumX) / n;

  return { m, c };
}

/**
 * Add a new divider line to the chart
 * @returns
 */
function addDividerLine() {
  // checks if limit of divider lines is reached
  // we only allow max 3 divider lines (in addition to the two invisible
  // ones we create)
  const dividerCount = state.dividerLines.length - 2;
  if (dividerCount >= 3) {
    return;
  }

  // This is the date value (type: number) where it should be
  // And this calculates the insertion point (25%, 50%, etc) for each
  // divider point
  let [xPosition, _] = xplot.convertFromPixel("grid", [
    (xplot.getWidth() * (dividerCount + 1)) / 4,
    0,
  ]);
  // trick: dividerLine might coincides with a data point, so we move it slightly to the right
  if (xPosition % 10) {
    xPosition += 1;
  }

  let dividerLine = {
    id: `divider-${dividerCount + 1}`,
    x: xPosition,
  };
  state.dividerLines.push(dividerLine);

  redraw("addDividerLine");
}

// redrawDividerButtons style "add/remove divider" buttons
function redrawDividerButtons() {
  if (state.dividerLines.length > 2) {
    document
      .querySelector("#remove-divider")
      .classList.remove("text-slate-400");
  } else {
    document.querySelector("#remove-divider").classList.add("text-slate-400");
  }
  if (state.dividerLines.length < 5) {
    document.querySelector("#add-divider").classList.remove("bg-slate-700");
    document
      .querySelector("#add-divider")
      .classList.remove("hover:bg-slate-600");
  } else {
    document.querySelector("#add-divider").classList.add("bg-slate-700");
    document.querySelector("#add-divider").classList.add("hover:bg-slate-600");
  }
}

// reflow charts to stack and fill in the entire screen width if there are a lot of data points
function reflowCharts() {
  let div = document.querySelector("#charts-container > div");
  const x = document.querySelector("#xplot-div");
  const mr = document.querySelector("#mrplot-div");
  if (state.xdata.length > 31) {
    div.classList.remove("lg:flex-nowrap");
    x.classList.remove("lg:w-1/2");
    mr.classList.remove("lg:w-1/2");
  } else {
    div.classList.add("lg:flex-nowrap");
    x.classList.add("lg:w-1/2");
    mr.classList.add("lg:w-1/2");
  }
}

function redraw(e: string = ""): _Stats {
  let stats = wrangleData();
  redrawDividerButtons();
  reflowCharts();
  doEChartsThings(stats);
  updateChartsYMinMax(stats);
  return stats;
}
window.redrawEcharts = redraw;

/**
 * Renders a vertical divider line
 * @param stats
 */
function renderDividerLine(dividerLine: DividerType, stats: _Stats) {
  if (isShadowDividerLine(dividerLine)) {
    // empty or undefined id means it is the shadow divider lines, so we don't render it
    return;
  }

  // Convert domain from data to pixel dimension
  const [xplotYMin, xplotYMax] = getXChartYAxisMinMax(stats);
  const p1 = xplot.convertToPixel("grid", [dividerLine.x, xplotYMin]);
  const p2 = xplot.convertToPixel("grid", [dividerLine.x, xplotYMax]);

  xplot.setOption({
    graphic: [
      {
        type: "line",
        id: dividerLine.id,
        shape: {
          x1: p1[0],
          y1: p1[1],
          x2: p2[0],
          y2: p2[1],
        },
        z: 99, // ensure the divider renders above all other lines
        style: {
          lineWidth: DIVIDER_LINE_WIDTH,
          lineDash: "solid",
          stroke: "purple",
        },
        draggable: "horizontal",
        ondragend: (dragEvent) => {
          for (let d of state.dividerLines) {
            if (d.id == dragEvent.target.id) {
              const translatedPt = xplot.convertFromPixel("grid", [
                dragEvent.offsetX,
                0,
              ]);
              d.x = translatedPt[0];
              break;
            }
          }
          redraw("moveDividerLine");
        },
        cursor: "ew-resize",
      },
    ],
  });
}
/**
 * Renders horizontal limit lines
 * @param stats
 */
function renderLimitLines(stats: _Stats) {
  // Create a bunch of series for each split
  let xSeries = [];
  let mrSeries = [];

  // Iterate through all splits (created by dividers) and draw their limit lines
  for (let i = 0; i < stats.lineValues.length; i++) {
    const isLastSegment = i === stats.lineValues.length - 1;

    let lv = stats.lineValues[i];

    let strokeWidth = LIMIT_LINE_WIDTH;
    let lineType = "dashed";
    let options = {
      useUpperQuartile: true,
      useLowerQuartile: true,
    };

    if (i == 0 && isLockedLimitsActive()) {
      options = shouldUseQuartile();
      // augment xLeft and xRight data because it is not included in the calculation
      state.lockedLimits.xLeft = lv.xLeft;
      state.lockedLimits.xRight = lv.xRight;
      lv = state.lockedLimits;
      strokeWidth = 3;
      lineType = "solid";
    }

    const createHorizontalLimitLineSeries = ({
      name,
      lineStyle,
      statisticY,
      showLabel = false,
    }: {
      name: string;
      lineStyle: any;
      statisticY: number;
      showLabel?: boolean;
    }) => ({
      name: name,
      type: "line",
      markLine: {
        symbol: ["none", "none"],
        lineStyle,
        tooltip: {
          show: true,
          formatter: `${statisticY}`,
          textStyle: {
            fontWeight: "bold",
          },
        },
        label: {
          formatter: showLabel && `${round(statisticY)}`,
          fontSize: 11,
          color: "#000",
        },
        emphasis: {
          disabled: true,
        },
        data: [
          [
            {
              xAxis: lv.xLeft,
              yAxis: statisticY,
            },
            {
              // add one day if we are at the last segment (extend line slightly beyond last value)
              xAxis:
                i == stats.lineValues.length - 1
                  ? dayjs(lv.xRight).add(1, "day").valueOf()
                  : lv.xRight,
              yAxis: statisticY,
            },
          ],
        ],
      },
    });

    mrSeries = mrSeries.concat([
      createHorizontalLimitLineSeries({
        name: `${i}-avgmovement`,
        lineStyle: {
          color: MEAN_SHAPE_COLOR,
          width: strokeWidth,
          type: lineType,
        },
        statisticY: lv.avgMovement ?? 0,
        showLabel: isLastSegment,
      }),
      createHorizontalLimitLineSeries({
        name: `${i}-URL`,
        lineStyle: {
          color: LIMIT_SHAPE_COLOR,
          width: strokeWidth,
          type: lineType,
        },
        statisticY: lv.URL ?? 0,
        showLabel: isLastSegment,
      }),
    ]);

    if (state.isShowingTrendLines && state.dividerLines.length <= 2) {
      // Don't show limit lines at all when trend lines are active and no dividers active.
      break;
    }
    if (state.isShowingTrendLines && i === 0) {
      // Don't show normal limit lines in the FIRST segment when trend lines are active.
      continue;
    }

    xSeries = xSeries.concat([
      options.useLowerQuartile &&
      createHorizontalLimitLineSeries({
        name: `${i}-low-Q`,
        lineStyle: {
          color: "gray",
          type: "dashed",
          dashOffset: 15,
          width: 1,
        },
        statisticY: lv.lowerQuartile ?? 0,
      }),
      options.useUpperQuartile &&
      createHorizontalLimitLineSeries({
        name: `${i}-upp-Q`,
        lineStyle: {
          color: "gray",
          type: "dashed",
          dashOffset: 15,
          width: 1,
        },
        statisticY: lv.upperQuartile ?? 0,
      }),
      createHorizontalLimitLineSeries({
        name: `${i}-avg`,
        lineStyle: {
          color: MEAN_SHAPE_COLOR,
          width: strokeWidth,
          type: lineType,
        },
        statisticY: lv.avgX ?? 0,
        showLabel: isLastSegment,
      }),
      createHorizontalLimitLineSeries({
        name: `${i}-unpl`,
        lineStyle: {
          color: LIMIT_SHAPE_COLOR,
          width: strokeWidth,
          type: lineType,
        },
        statisticY: lv.UNPL ?? 0,
        showLabel: isLastSegment,
      }),
      createHorizontalLimitLineSeries({
        name: `${i}-lnpl`,
        lineStyle: {
          color: LIMIT_SHAPE_COLOR,
          width: strokeWidth,
          type: lineType,
        },
        statisticY: lv.LNPL ?? 0,
        showLabel: isLastSegment,
      }),
    ]);
  }

  xplot.setOption({
    series: xSeries,
  });
  mrplot.setOption({
    series: mrSeries,
  });
}

function removeDividerLine() {
  // remove the last added annotation
  let id = `divider-${state.dividerLines.length - 2}`;
  state.dividerLines = state.dividerLines.filter((d) => d.id != id);
  redraw("removeDividerLine");
}

function updateChartsYMinMax(stats: _Stats) {
  // We have calculated the max of the limit lines before, now we compare with the value of each data point
  // since some of them might be outside the limit and we want it to be in view
  let allValues = deepClone(state.xdata);
  if (state.isShowingTrendLines) {
    allValues = allValues.concat(state.trendLines["lnpl"]);
    allValues = allValues.concat(state.trendLines["unpl"]);
  }

  removeAllNulls(allValues).forEach((dv) => {
    if (dv.value > stats.xchartMax) {
      stats.xchartMax = dv.value;
    }
    if (dv.value < stats.xchartMin) {
      stats.xchartMin = dv.value;
    }
  });
  state.movements.forEach((dv) => {
    if (dv.value > stats.mrchartMax) {
      stats.mrchartMax = dv.value;
    }
  });
  // we now compare with the user-set locked limit lines
  if (isLockedLimitsActive()) {
    stats.xchartMax = Math.max(stats.xchartMax, state.lockedLimits.UNPL);
    stats.xchartMin = Math.min(stats.xchartMin, state.lockedLimits.LNPL);
    stats.mrchartMax = Math.max(stats.mrchartMax, state.lockedLimits.URL);
  }
}

// HELPER FUNCTIONS
// Does 2 things:
// (1) Filters data based on divider lines and calculates SPC statistics per range
// By default,`dividerLines` contains two 'invisible' dividerlines at both ends
// with no Line associated with it (so it doesn't get rendered in the chart)
// (2) Checks against the 3 XMR rules for each range and color data points accordingly
function wrangleData(): _Stats {
  let dividerLines = state.dividerLines;
  // need to make sure dividerLines are sorted
  dividerLines.sort((a, b) => a.x - b.x);

  console.assert(
    dividerLines.length >= 2,
    "dividerLines should contain at least two divider lines"
  );

  // make sure state.xdata only contains valid data (i.e. have both x and value columns set)
  // and it is sorted by date (x) ascending
  let tableData = state.tableData.filter(
    (dv) => dv.x && (dv.value || dv.value == 0)
  );
  sortDataValues(tableData);
  updateInPlace(state.xdata, tableData);

  // Seasonality
  updateDeseasonalisedData();
  if (state.isShowingDeseasonalisedData) {
    state.xdata = deepClone(state.deSeasonalisedData);
  }

  // Since a user might paste in data that falls beyond either limits of the previous x-axis range
  // we need to update our "shadow" divider lines so that the filteredXdata will always get all data
  let { min: xdataXmin, max: xdataXmax } = findExtremesX(state.xdata);
  dividerLines[0].x = xdataXmin;
  dividerLines[dividerLines.length - 1].x = xdataXmax;

  // chartMin is the lowest y-value that needs to be drawn in the chart
  // chartMax is the highest y-value that needs to be drawn in the chart
  // xdataPerRange groups state.xdata based on ranges between divider lines. Initially, all xdata is in 1 range.
  // movementsPerRange groups movements based on ranges between divider lines. Initially, all movements is in 1 range.
  const stats = {
    xchartMin: Infinity,
    xchartMax: -Infinity,
    mrchartMax: -Infinity,
    lineValues: [] as LineValueType[],
    xdataPerRange: [] as DataValue[][],
    movementsPerRange: [] as DataValue[][],
  };

  let xdataWithStatus: DataValue[] = [];
  let mrdataWithStatus: DataValue[] = [];
  for (let i = 0; i < dividerLines.length - 1; i++) {
    const xLeft = dividerLines[i].x;
    const xRight = dividerLines[i + 1].x;
    // We ignore the edge case where the dividerlines is precisely at T00:00.
    const filteredXdata = state.xdata.filter((d) => {
      return fromDateStr(d.x) >= xLeft && fromDateStr(d.x) <= xRight;
    });
    const filteredMovements = getMovements(filteredXdata);
    // if no data in range, skip
    if (filteredXdata.length === 0) {
      console.log(`No data in range ${xLeft} - ${xRight}`);
      continue;
    }
    const { avgX, avgMovement, UNPL, LNPL, URL, lowerQuartile, upperQuartile } =
      calculateLimits(filteredXdata);
    // append line values
    let lv: LineValueType = {
      xLeft,
      xRight,
      avgX,
      avgMovement,
      UNPL,
      LNPL,
      URL,
      lowerQuartile,
      upperQuartile,
    };
    stats.lineValues.push(lv);
    stats.xchartMin = Math.min(stats.xchartMin, LNPL);
    stats.xchartMax = Math.max(stats.xchartMax, UNPL);
    stats.mrchartMax = Math.max(stats.mrchartMax, URL);

    // check for process exceptions
    //
    // We need to first reset all status to normal first before calculating
    // to prevent cached status from previous checks
    filteredXdata.forEach((dv) => (dv.status = DataStatus.NORMAL));
    if (i == 0 && isLockedLimitsActive()) {
      // locked limits only apply to the left-most section
      let opts = shouldUseQuartile();
      checkRunOfEight(filteredXdata, state.lockedLimits.avgX);

      checkFourNearLimit(
        filteredXdata,
        opts.useLowerQuartile ? state.lockedLimits.lowerQuartile : -Infinity,
        opts.useUpperQuartile ? state.lockedLimits.upperQuartile : Infinity
      );
      checkOutsideLimit(
        filteredXdata,
        state.lockedLimits.LNPL,
        state.lockedLimits.UNPL
      );
    } else if (i == 0 && state.isShowingTrendLines) {
      // trend lines, like locked limits only apply to the left-most section
      checkRunOfEight(filteredXdata, state.trendLines.centreLine);
      checkFourNearLimit(
        filteredXdata,
        state.trendLines.lowerQtl,
        state.trendLines.upperQtl
      );
      checkOutsideLimit(
        filteredXdata,
        state.trendLines.lnpl,
        state.trendLines.unpl
      );
    } else {
      checkRunOfEight(filteredXdata, avgX);
      checkFourNearLimit(filteredXdata, lowerQuartile, upperQuartile);
      checkOutsideLimit(filteredXdata, LNPL, UNPL);
    }
    xdataWithStatus = xdataWithStatus.concat(filteredXdata);
    stats.xdataPerRange.push(filteredXdata);

    // check for movement exceptions
    filteredMovements.forEach((dv) => (dv.status = DataStatus.NORMAL));
    if (i == 0 && isLockedLimitsActive()) {
      checkOutsideLimit(filteredMovements, 0, state.lockedLimits.URL);
    } else {
      checkOutsideLimit(filteredMovements, 0, URL);
    }
    mrdataWithStatus = mrdataWithStatus.concat(filteredMovements);
    stats.movementsPerRange.push(filteredMovements);
  }

  updateInPlace(state.xdata, xdataWithStatus);
  updateInPlace(state.movements, mrdataWithStatus);

  updateChartsYMinMax(stats);
  return stats;
}

/**
 * Tests and parses the CSV file input string.
 * @param str: string of the CSV file
 * @param delimiter: delimiter of the CSV file, default is comma
 * @returns an object containing the test result, multiplier, xLabel, yLabel, and xdata
 */
function csvTestingParser(str: string, delimiter = ",") {
  // self-defined variables
  let xLabel = "";
  let yLabel = "";
  let multiplier = 0;
  const xdata: DataValue[] = [];

  // testing axes label inputs/empty CSV
  const firstBreak = str.indexOf("\r\n");

  if (firstBreak == -1) {
    const errorMsg = document.getElementById("file-error") as HTMLDivElement;
    errorMsg.style.display = "block";
    errorMsg.innerText = "Missing CSV labels and/or data.";
    return { passed: false, multiplier, xLabel, yLabel, xdata };
  }
  const labels = str.slice(0, firstBreak).split(delimiter);
  if (labels.length < 2) {
    const errorMsg = document.getElementById("file-error") as HTMLDivElement;
    errorMsg.style.display = "block";
    errorMsg.innerText = "First row of CSV must have 2 columns.";
    return { passed: false, multiplier, xLabel, yLabel, xdata };
  } else if (labels[0] === "" || labels[1] === "") {
    const errorMsg = document.getElementById("file-error") as HTMLDivElement;
    errorMsg.style.display = "block";
    errorMsg.innerText = "Missing CSV label(s).";
    return { passed: false, multiplier, xLabel, yLabel, xdata };
  } else if (labels[0].toLowerCase() != "date") {
    const errorMsg = document.getElementById("file-error") as HTMLDivElement;
    errorMsg.style.display = "block";
    errorMsg.innerText = "First column of CSV must be 'Date'.";
    return { passed: false, multiplier, xLabel, yLabel, xdata };
  }
  // set x and y labels
  xLabel = "Date";
  yLabel = labels[1];

  // testing data inputs
  const rows = str.slice(firstBreak + 2).split("\r\n");
  if (rows.length == 0) {
    const errorMsg = document.getElementById("file-error") as HTMLDivElement;
    errorMsg.style.display = "block";
    errorMsg.innerText = "Missing CSV data.";
    return { passed: false, multiplier, xLabel, yLabel, xdata };
  }
  // This is the actual parising of the csv.
  for (let i = 0; i < rows.length; i++) {
    const values = rows[i].split(delimiter);
    if (values[0] === "" || values[1] === "") {
      // split apart the remaining rows of the array string and check if any contain values
      // if yes, display error; else, return passed test and multiplier
      const remainingContent = rows.slice(i + 1).join(delimiter);
      // if remaining content is not a string of only commas, return failed test, else return passed test
      if (remainingContent.replace(/,/g, "") !== "") {
        const errorMsg = document.getElementById(
          "file-error"
        ) as HTMLDivElement;
        errorMsg.style.display = "block";
        errorMsg.innerText = "Fragmented CSV data.";
        return { passed: false, multiplier: 0, xLabel, yLabel, xdata };
      } else {
        return { passed: true, multiplier, xLabel, yLabel, xdata };
      }
    }
    const parsedDate = Date.parse(values[0]);
    if (!parsedDate) {
      const errorMsg = document.getElementById("file-error") as HTMLDivElement;
      errorMsg.style.display = "block";
      errorMsg.innerText =
        "Please input date in YYYY-MM-DD format (if on Excel, change date format settings).";
      return { passed: false, multiplier, xLabel, yLabel, xdata };
    }
    let parsedVal = Number(values[1]);
    if (isNaN(parsedVal)) {
      const errorMsg = document.getElementById("file-error") as HTMLDivElement;
      errorMsg.style.display = "block";
      errorMsg.innerText = "Values must be numbers.";
      return { passed: false, multiplier, xLabel, yLabel, xdata };
    } else {
      // check if too many digits, and increment multiplier until its happy
      while (Math.abs(parsedVal) / 10 ** multiplier >= 10000) {
        multiplier += 3;
      }
      xdata.push({
        order: i,
        x: values[0],
        value: parsedVal,
        status: DataStatus.NORMAL,
      });
    }
  }
  return { passed: true, multiplier, xLabel, yLabel, xdata };
}

function downloadCSV(state, isLockedLimitsActive: boolean) {
  // only download lines with at least one non-null value.
  let csvContent =
    `${state.xLabel},${state.yLabel}\r\n` +
    state.xdata
      .filter((dv) => dv.x || dv.value)
      .map((d) => `${d.x || ""},${round(d.value) || ""}`)
      .join("\r\n");
  if (isLockedLimitsActive) {
    csvContent +=
      `\r\n\r\nlimit_lines,value\r\n` +
      [
        `avg_x,${state.lockedLimits.avgX}`,
        `LNPL,${state.lockedLimits.LNPL}`,
        `UNPL,${state.lockedLimits.UNPL}`,
        `avg_movement,${state.lockedLimits.avgMovement}`,
        `URL,${state.lockedLimits.URL}`,
      ].join("\r\n");
  }
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const dataUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.setAttribute("href", dataUrl);
  link.setAttribute("download", "xmr-data.csv");
  link.click();
  // Clean up the anchor element
  link.remove();
  // Revoke the object URL to free up memory
  URL.revokeObjectURL(dataUrl);
}

/**
 * Returns an array of movements given an array of data values
 * @param xdata: array of data values
 * @returns an array of MR values
 */
function getMovements(xdata: DataValue[]): DataValue[] {
  const movements = [];
  for (let i = 1; i < xdata.length; i++) {
    const diff = round(Math.abs(xdata[i].value - xdata[i - 1].value));
    movements.push({ order: xdata[i].order, x: xdata[i].x, value: diff });
  }
  return movements;
}

// window.onresize doesn't play well with mobile browsers
// we use matchMedia to catch when we cross the 767px boundary.
// https://www.cocomore.com/blog/dont-use-window-onresize
const mql = window.matchMedia("(max-width: 767px)");
mql.addEventListener("change", (e) => {
  redraw("windowResize");
});

screen.orientation.addEventListener("change", (e) => {
  redraw("changeOrientation");
});

let hot: Handsontable;
let lockedLimitHot: Handsontable;
let deseasonaliseHot: Handsontable;
let seasonalFactorsHot: Handsontable;
let trendHot: Handsontable;

let dTableColors = [];

function updateDTableSeasonalFactorMapping() {
  // given row Index, we need to figure out what color to use
  const periodisedData = periodiseData(
    state.xdata[0].x,
    state.deSeasonalisedData,
    state.deSeasonalisePeriod
  );

  const periodSize = Math.max(...periodisedData.map((p) => p.length));

  // generate N different colors
  dTableColors = chroma.scale(["white", "#bfdbfe"]).colors(periodSize);

  // update season numbers in season dialog table
  state.seasonalFactorData.forEach((d, i) => {
    if (state.deSeasonalisedData[i]) {
      d.seasonalFactor = state.deSeasonalisedData[i].seasonalFactor;
    }
  });

  if (deseasonaliseHot) {
    deseasonaliseHot.updateData(state.seasonalFactorData);
  }
}

const _seasonalDataCellRenderer = (
  instance,
  td,
  row,
  col,
  prop,
  value,
  cellProperties
) => {
  // This renders the text
  Handsontable.renderers.TextRenderer(
    instance,
    td,
    row,
    col,
    prop,
    value,
    cellProperties
  );

  if (
    row < state.seasonalFactorData.length - 1 &&
    state.seasonalFactorData[row]
  ) {
    const sfIndex = state.seasonalFactorData[row].seasonalFactor - 1;
    td.style.background = dTableColors[sfIndex] ?? "salmon";
    td.style.borderBottomColor = dTableColors[sfIndex];
  }
};

const _seasonalFactorCellRenderer = (
  instance,
  td,
  row,
  col,
  prop,
  value,
  cellProperties
) => {
  // This renders the text
  Handsontable.renderers.TextRenderer(
    instance,
    td,
    row,
    col,
    prop,
    value,
    cellProperties
  );

  td.style.background = dTableColors[row];
  td.style.borderBottomColor = dTableColors[row];
};

// maps function to a lookup string
Handsontable.renderers.registerRenderer(
  "seasonalDataRenderer",
  _seasonalDataCellRenderer
);
Handsontable.renderers.registerRenderer(
  "seasonalFactorRenderer",
  _seasonalFactorCellRenderer
);

function lockLimits() {
  const lockedLimits = calculateLockedLimits(); // calculate locked limits from the table
  let obj = structuredClone(INACTIVE_LOCKED_LIMITS);

  document
    .querySelectorAll(".lock-limit-input")
    .forEach((el: HTMLInputElement) => {
      obj[el.dataset.limit] =
        el.value !== "" ? Number(el.value) : lockedLimits[el.dataset.limit];
    });
  obj.lowerQuartile = round((obj.avgX + obj.LNPL) / 2);
  obj.upperQuartile = round((obj.avgX + obj.UNPL) / 2);

  // validate user input
  if (obj.avgX < obj.LNPL || obj.avgX > obj.UNPL || obj.avgMovement > obj.URL) {
    alert(
      "Please ensure that the following limits are satisfied:\n" +
      "1. Average X is between Lower Natural Process Limit (LNPL) and Upper Natural Process Limit (UNPL)\n" +
      "2. Average Movement is less than or equal to Upper Range Limit (URL)"
    );
    return;
  }
  if (lockedLimits.avgX != obj.avgX) {
    state.lockedLimitStatus |= LockedLimitStatus.AVGX_MODIFIED;
  }
  if (lockedLimits.LNPL != obj.LNPL) {
    state.lockedLimitStatus |= LockedLimitStatus.LNPL_MODIFIED;
  }
  if (lockedLimits.UNPL != obj.UNPL) {
    state.lockedLimitStatus |= LockedLimitStatus.UNPL_MODIFIED;
  }
  state.lockedLimits = obj; // set state
  state.lockedLimitStatus |= LockedLimitStatus.LOCKED; // set to locked

  redraw("addLockLimit");
}

function initialiseHandsOnTable() {
  const seasonalFactorsDataTable = document.querySelector(
    "#seasonal-factors-dataTable"
  ) as HTMLDivElement;
  const trendDataTable = document.querySelector(
    "#trend-dataTable"
  ) as HTMLDivElement;
  const deseasonaliseDataTable = document.querySelector(
    "#deseason-dataTable"
  ) as HTMLDivElement;
  const lockedLimitDataTable = document.querySelector(
    "#lock-limit-dataTable"
  ) as HTMLDivElement;
  const table = document.querySelector("#dataTable") as HTMLDivElement;

  trendHot = trendDataTable ? new Handsontable(trendDataTable, {
    licenseKey: "non-commercial-and-evaluation", // for non-commercial use only
    data: state.trendData,
    dataSchema: { x: null, value: null },
    columns: [
      {
        data: "x",
        type: "date",
        dateFormat: "YYYY-MM-DD",
        validator: (value, callback) => {
          if (!value) {
            // if null, "", or undefined
            callback(true);
            return;
          }

          let d = new Date(fromDateStr(value));
          // https://stackoverflow.com/questions/1353684/detecting-an-invalid-date-date-instance-in-javascript#1353711
          // callback(true) if date != 'Invalid Date'
          callback(d instanceof Date && !isNaN(d.getTime()));
        },
      },
      { data: "value", type: "numeric" },
    ],
    colHeaders: [state.xLabel, state.yLabel],
    // Show context menu to enable removing rows.
    contextMenu: true,
    allowRemoveColumn: false,
    minSpareRows: 1,
    height: "auto",
    stretchH: "all",
    fillHandle: {
      autoInsertRow: true,
      direction: "vertical",
    },
    beforeAutofill(selectionData, sourceRange, targetRange, direction) {
      return autofillTable(selectionData, sourceRange, targetRange, direction);
    },
    beforePaste(data, coords) {
      return beforePasteTable(data, coords);
    },
    afterChange(changes, source) {
      if (source === "loadData") {
        return;
      }
      calculateTrendStats();
    },
    afterValidate(isValid, value, row, prop, source) {
      const errorMsg = document.getElementById("data-table-error");
      if (isValid) {
        errorMsg.classList.add("hidden");
        return;
      }
      errorMsg.classList.remove("hidden");
      return false;
    },
  }) : undefined;

  seasonalFactorsHot = seasonalFactorsDataTable ? new Handsontable(seasonalFactorsDataTable, {
    colHeaders: ["Seasonal Factors"],
    data: state.seasonalFactorTableData,
    rowHeaders: true,
    // colHeaders: sfCols,
    // Show context menu to enable removing rows.
    contextMenu: true,
    allowRemoveColumn: false,
    height: "auto",
    stretchH: "all",
    allowInsertRow: false,
    allowInsertColumn: false,
    beforeAutofill(selectionData, sourceRange, targetRange, direction) {
      return autofillTable(selectionData, sourceRange, targetRange, direction);
    },
    beforePaste(data, coords) {
      return beforePasteTable(data, coords);
    },
    afterChange(changes, source) {
      if (source === "edit" && state.isShowingDeseasonalisedData) {
        redraw("updateDeseasonalisedData");
      }
    },
    afterValidate(isValid, value, row, prop, source) {
      const errorMsg = document.getElementById("data-table-error");
      if (isValid) {
        errorMsg.classList.add("hidden");
        return;
      }
      errorMsg.classList.remove("hidden");
      return false;
    },
    // the `cells` options overwrite all other options
    cells(row, col, prop) {
      const cellProperties: CellMeta = {
        renderer: "seasonalFactorRenderer",
      };
      if (row === 0) {
        cellProperties.type = "numeric";
      }
      return cellProperties;
    },
    licenseKey: "non-commercial-and-evaluation", // for non-commercial use only
  }) : undefined;

  deseasonaliseHot = deseasonaliseDataTable ? new Handsontable(deseasonaliseDataTable, {
    data: state.seasonalFactorData,
    dataSchema: { x: null, value: null },
    columns: [
      {
        data: "x",
        type: "date",
        dateFormat: "YYYY-MM-DD",
        validator: (value, callback) => {
          if (!value) {
            // if null, "", or undefined
            callback(true);
            return;
          }

          let d = new Date(fromDateStr(value));
          // https://stackoverflow.com/questions/1353684/detecting-an-invalid-date-date-instance-in-javascript#1353711
          // callback(true) if date != 'Invalid Date'
          callback(d instanceof Date && !isNaN(d.getTime()));
        },
      },
      { data: "value", type: "numeric" },
      { data: "seasonalFactor", type: "numeric" },
    ],
    colHeaders: [state.xLabel, state.yLabel, "Season"],
    // Show context menu to enable removing rows.
    contextMenu: true,
    allowRemoveColumn: false,
    minSpareRows: 1,
    height: "auto",
    stretchH: "all",
    fillHandle: {
      autoInsertRow: true,
      direction: "vertical",
    },
    beforeAutofill(selectionData, sourceRange, targetRange, direction) {
      return autofillTable(selectionData, sourceRange, targetRange, direction);
    },
    afterChange(changes, source) {
      // Check range of values entered
      // get range of data in deseason table
      const data = sortDataValues(
        state.seasonalFactorData.filter((d) => d.x != null)
      );
      const startDate = dayjs(data[0].x);
      const endDate = dayjs(data[data.length - 1].x);

      // Get selected period
      const selectedPeriod = (document.querySelector("#period-select") as HTMLSelectElement).value;

      // Calculate difference based on selected period
      const periodDiff = Math.abs(startDate.diff(endDate, selectedPeriod as any, true));

      // Show warning if less than one period
      if (periodDiff < 1) {
        showElement(document.querySelector("#deseason-warn-1"));
      } else {
        hideElement(document.querySelector("#deseason-warn-1"));
      }

      // Show warning if exactly one period
      if (periodDiff === 1) {
        showElement(document.querySelector("#deseason-warn-2"));
      } else {
        hideElement(document.querySelector("#deseason-warn-2"));
      }

      if (source === "edit") {
        updateSeasonalFactors();
      }
    },
    afterValidate(isValid, value, row, prop, source) {
      const errorMsg = document.getElementById("data-table-error");
      if (isValid) {
        errorMsg.classList.add("hidden");
        return;
      }
      errorMsg.classList.remove("hidden");
      return false;
    },
    cells(row, col) {
      const cellProperties: CellMeta = {};

      cellProperties.renderer = "seasonalDataRenderer"; // uses lookup map

      return cellProperties;
    },
    licenseKey: "non-commercial-and-evaluation", // for non-commercial use only
  }) : undefined;

  lockedLimitHot = lockedLimitDataTable ? new Handsontable(lockedLimitDataTable, {
    data: state.lockedLimitBaseData,
    dataSchema: { x: null, value: null },
    columns: [
      {
        data: "x",
        type: "date",
        dateFormat: "YYYY-MM-DD",
        validator: (value, callback) => {
          if (!value) {
            // if null, "", or undefined
            callback(true);
            return;
          }

          let d = new Date(fromDateStr(value));
          // https://stackoverflow.com/questions/1353684/detecting-an-invalid-date-date-instance-in-javascript#1353711
          // callback(true) if date != 'Invalid Date'
          callback(d instanceof Date && !isNaN(d.getTime()));
        },
      },
      { data: "value", type: "numeric" },
    ],
    colHeaders: [state.xLabel, state.yLabel],
    // Show context menu to enable removing rows.
    contextMenu: true,
    allowRemoveColumn: false,
    minSpareRows: 1,
    height: "auto",
    stretchH: "all",
    fillHandle: {
      autoInsertRow: true,
      direction: "vertical",
    },
    beforeAutofill(selectionData, sourceRange, targetRange, direction) {
      return autofillTable(selectionData, sourceRange, targetRange, direction);
    },
    beforePaste(data, coords) {
      return beforePasteTable(data, coords);
    },
    afterChange(changes, source) {
      if (source === "loadData") {
        return;
      }
      setLockedLimitInputs(true);
    },
    afterValidate(isValid, value, row, prop, source) {
      const errorMsg = document.getElementById("data-table-error");
      if (isValid) {
        errorMsg.classList.add("hidden");
        return;
      }
      errorMsg.classList.remove("hidden");
      return false;
    },
    licenseKey: "non-commercial-and-evaluation", // for non-commercial use only
  }) : undefined;

  hot = new Handsontable(table, {
    data: state.tableData,
    dataSchema: { x: null, value: null },
    columns: [
      {
        data: "x",
        type: "date",
        dateFormat: "YYYY-MM-DD",
        validator: (value, callback) => {
          if (!value) {
            // if null, "", or undefined
            callback(true);
            return;
          }

          let d = new Date(fromDateStr(value));
          // https://stackoverflow.com/questions/1353684/detecting-an-invalid-date-date-instance-in-javascript#1353711
          // callback(true) if date != 'Invalid Date'
          callback(d instanceof Date && !isNaN(d.getTime()));
        },
      },
      { data: "value", type: "numeric" },
    ],
    colHeaders: [
      state.xLabel ?? "Date",
      state.yLabel === "Value" ? "Value (✏️)" : state.yLabel,
    ],
    // Editable column header: https://github.com/handsontable/handsontable/issues/1980
    afterOnCellMouseDown: function (e, coords) {
      if (coords.row !== -1) {
        return;
      }
      let newColName = prompt(
        "Insert a new column name",
        this.getColHeader()[coords.col]
      );
      if (newColName) {
        let colHeaders = this.getColHeader();
        colHeaders[coords.col] = newColName;
        this.updateSettings({
          colHeaders,
        });
        if (coords.col == 0) {
          state.xLabel = newColName;
        } else {
          state.yLabel = newColName;
        }
        // Redraw
        redraw("editColumnName");
      }
    },
    // Show context menu to enable removing rows.
    contextMenu: true,
    allowRemoveColumn: false,
    minSpareRows: 1,
    height: "auto",
    stretchH: "all",
    fillHandle: {
      autoInsertRow: true,
      direction: "vertical",
    },
    beforeAutofill(selectionData, sourceRange, targetRange, direction) {
      return autofillTable(selectionData, sourceRange, targetRange, direction);
    },
    beforePaste(data, coords) {
      return beforePasteTable(data, coords);
    },
    beforeChange(changes, source) {
      let xOnly = state.xdata.map((d) => d.x);
      changes.forEach(([row, prop, oldVal, newVal], idx) => {
        if (prop == "x") {
          xOnly[row] = newVal;
        }
        if (prop == "value") {
          // force float by removing special characters
          changes[idx][3] = forceFloat(newVal);
        }
      });
      checkDuplicatesInTable(xOnly, this);
    },
    afterChange(changes, source) {
      if (source == "loadData") {
        return;
      }
      redraw("editTableData");
    },
    afterValidate(isValid, value, row, prop, source) {
      const errorMsg = document.getElementById("data-table-error");
      if (isValid) {
        errorMsg.classList.add("hidden");
        return;
      }
      errorMsg.classList.remove("hidden");
      return false;
    },
    licenseKey: "non-commercial-and-evaluation", // for non-commercial use only
  });
}

// LOGIC ON PAGE LOAD
document.addEventListener("DOMContentLoaded", async function (_e) {
  const seasonalFactorsRemoveButtons = document.querySelectorAll(
    ".deseason-remove-btn"
  ) as NodeListOf<HTMLButtonElement>;
  const trendRemoveButtons = document.querySelectorAll(
    ".trend-remove-btn"
  ) as NodeListOf<HTMLButtonElement>;
  const lockLimitButton = document.querySelector(
    "#lock-limit-btn"
  ) as HTMLButtonElement;
  const seasonalityPeriodSelect = document.querySelector(
    "#period-select"
  ) as HTMLSelectElement;
  const seasonalityPeriodOption = (
    opt: "year" | "quarter" | "month" | "week"
  ) => document.querySelector(`#period-option-${opt}`) as HTMLOptionElement;
  const seasonalFactorsButton = document.querySelector(
    "#de-seasonalise-btn"
  ) as HTMLButtonElement;
  const trendDialogOpenButton = document.querySelector(
    "#trend-open-btn"
  ) as HTMLButtonElement;

  const pageParams = extractDataFromUrl();

  let isDeseasonalised = false;
  let isShowingTrend = false;

  if (pageParams.has("d")) {
    // Load data
    let version = pageParams.get("v") || "0";
    let separator = pageParams.get("s") || "";
    let ll = pageParams.get("l") || "";
    let sf = pageParams.get("p") || "";
    let tl = pageParams.get("t") || "";

    isShowingTrend = tl !== "";
    isDeseasonalised = sf !== "";

    let {
      xLabel,
      yLabel,
      data,
      dividerLines,
      lockedLimits,
      lockedLimitStatus,
      seasonalFactors,
      regressionStats,
    } = await decodeShareLink(
      version,
      pageParams.get("d")!,
      separator,
      ll,
      sf,
      tl
    );

    state.xLabel = xLabel;
    state.yLabel = yLabel;
    state.tableData = data;
    state.dividerLines = dividerLines;
    state.lockedLimits = lockedLimits;
    state.lockedLimitStatus = lockedLimitStatus;

    if (isDeseasonalised) {
      const period = pageParams.get("p0");
      let periodIsValid = period !== null;

      if (["year", "quarter", "month", "week"].indexOf(period) === -1) {
        periodIsValid = false;
      }

      if (periodIsValid) {
        state.deSeasonalisePeriod = period as SeasonalityPeriod;
        seasonalityPeriodSelect.value = state.deSeasonalisePeriod;
      }

      state.seasonalFactorTableData = seasonalFactors.map((f) => [f]);
      updateDeseasonalisedData();
      state.isShowingDeseasonalisedData = true;
      seasonalFactorsRemoveButtons.forEach(showElement);

      // disable trends
      trendDialogOpenButton.disabled = true;
    }

    if (isShowingTrend) {
      state.regressionStats = regressionStats;
      state.isShowingTrendLines = true;
      trendRemoveButtons.forEach(showElement);
      // disable seasonality
      seasonalFactorsButton.disabled = true;
      // disable locked limits
      lockLimitButton.disabled = true;
    }
  } else {
    state.tableData = DUMMY_DATA;
  }

  // initialise locked limit base data with default data
  state.lockedLimitBaseData = deepClone(state.tableData);

  // initialise seasonal factor data with default data
  state.seasonalFactorData = deepClone(state.tableData);
  if (!isDeseasonalised) {
    updateSeasonalFactors();
  }

  state.trendData = deepClone(state.tableData);

  renderCharts("init");

  // Divider Buttons
  const addDividerButton = document.querySelector(
    "#add-divider"
  ) as HTMLButtonElement;
  const removeDividerButton = document.querySelector(
    "#remove-divider"
  ) as HTMLButtonElement;

  addDividerButton.addEventListener("click", addDividerLine);
  removeDividerButton.addEventListener("click", removeDividerLine);

  // Initialize all modals
  initializeModal("lock-limit-dialog", "lock-limit-backdrop", "lock-limit-close");
  initializeModal("deseason-dialog", "deseason-backdrop", "deseason-close");
  initializeModal("trend-dialog", "trend-backdrop", "trend-close-btn");

  // Seasonal Factors
  const seasonalFactorsDialog = document.querySelector(
    "#deseason-dialog"
  ) as HTMLDialogElement;
  const seasonalFactorsDialogCloseButton = document.querySelector(
    "#deseason-close"
  ) as HTMLButtonElement;
  const seasonalFactorsDialogApplyButton = document.querySelector(
    "#deseason-add"
  ) as HTMLButtonElement;
  const seasonalFactorsDialogResetButton = document.querySelector(
    "#deseason-reset-data"
  ) as HTMLButtonElement;
  const seasonalFactorsDialogResetFactors = document.querySelector(
    "#deseason-reset-factors"
  ) as HTMLButtonElement;
  const seasonalityGroupingSelect = document.querySelector(
    "#grouping-select"
  ) as HTMLSelectElement;

  seasonalFactorsDialogResetButton?.addEventListener(
    "click",
    () => {
      resetDeseasonalisedDialogTable();
      if (state.isShowingDeseasonalisedData) {
        redraw("resetDeseasonalisedData");
      }
    }
  );

  seasonalityGroupingSelect?.addEventListener("change", () => {
    state.deSeasonaliseGrouping =
      seasonalityGroupingSelect.value as SeasonalityGrouping;
    updateSeasonalFactors();
  });

  seasonalityPeriodSelect?.addEventListener("change", () => {
    state.deSeasonalisePeriod =
      seasonalityPeriodSelect.value as SeasonalityPeriod;
    updateSeasonalFactors();
    updateDeseasonalisedData();
  });
  seasonalFactorsButton?.addEventListener("click", () => {
    if (!state.isShowingDeseasonalisedData) {
      resetDeseasonalisedDialogTable()
    }

    // disable certain options based on the determined periodicity
    const interval = determinePeriodicity(state.xdata);
    if (interval === "quarter") {
      seasonalityPeriodOption("quarter").disabled = true;
      seasonalityPeriodOption("month").disabled = true;
      seasonalityPeriodOption("week").disabled = true;
    } else if (interval === "month") {
      seasonalityPeriodOption("quarter").disabled = false;
      seasonalityPeriodOption("month").disabled = true;
      seasonalityPeriodOption("week").disabled = true;
    } else if (interval === "week") {
      seasonalityPeriodOption("quarter").disabled = false;
      seasonalityPeriodOption("month").disabled = false;
      seasonalityPeriodOption("week").disabled = true;
    } else {
      seasonalityPeriodOption("quarter").disabled = false;
      seasonalityPeriodOption("month").disabled = false;
      seasonalityPeriodOption("week").disabled = false;
    }
    showModal("deseason-dialog", "deseason-backdrop");
  });
  seasonalFactorsDialogCloseButton?.addEventListener("click", () => {
    hideModal("deseason-dialog", "deseason-backdrop");
  });
  seasonalFactorsDialogApplyButton?.addEventListener("click", () => {
    state.isShowingDeseasonalisedData = true;
    redraw("applyDeseasonalisation");
    seasonalFactorsRemoveButtons.forEach(showElement);
    hideModal("deseason-dialog", "deseason-backdrop");
    // disable trends
    trendDialogOpenButton.disabled = true;
  });
  seasonalFactorsRemoveButtons?.forEach((b) =>
    b.addEventListener("click", () => {
      resetDeseasonalisedDialogTable();
      state.isShowingDeseasonalisedData = false;
      redraw("removeDeseasonalisation");
      seasonalFactorsRemoveButtons.forEach(hideElement);
      hideModal("deseason-dialog", "deseason-backdrop");
      // enable trends
      trendDialogOpenButton.disabled = false;
    })
  );
  seasonalFactorsDialogResetFactors?.addEventListener("click", () => {
    updateSeasonalFactors();
    if (state.isShowingDeseasonalisedData) {
      redraw("resetSeasonalFactors");
    }
  });

  const periodSelect = document.querySelector("#period-select") as HTMLSelectElement;
  if (periodSelect) {
    periodSelect.addEventListener("change", (e) => {
      const select = e.target as HTMLSelectElement;
      const periodText = select.options[select.selectedIndex].text.toLowerCase();
      document.querySelectorAll(".selected-period")?.forEach(selectedPeriodSpan => {
        selectedPeriodSpan.textContent = periodText === "annual" ? "year" : periodText;
      });
    });
  }


  // Trends
  calculateTrendStats();

  const trendDialog = document.querySelector(
    "#trend-dialog"
  ) as HTMLDialogElement;
  const trendDialogCloseButton = document.querySelector(
    "#trend-close-btn"
  ) as HTMLButtonElement;
  const trendDialogApplyButton = document.querySelector(
    "#trend-apply-btn"
  ) as HTMLButtonElement;
  const trendDialogResetButton = document.querySelector(
    "#trend-reset-data"
  ) as HTMLButtonElement;
  const trendSlopeInput = document.querySelector(
    "#trend-slope-input"
  ) as HTMLInputElement;
  const trendInterceptInput = document.querySelector(
    "#trend-intercept-input"
  ) as HTMLInputElement;

  trendDialogOpenButton?.addEventListener("click", () => {
    if (!state.isShowingTrendLines) {
      resetTrendDialogTable();
    }
    showModal("trend-dialog", "trend-backdrop")
  });
  trendDialogCloseButton?.addEventListener("click", () => hideModal("trend-dialog", "trend-backdrop"));
  trendDialogApplyButton?.addEventListener("click", () => {
    state.trendLines = createTrendlines(state.regressionStats, state.xdata);
    state.isShowingTrendLines = true;
    // disable seasonality
    seasonalFactorsButton.disabled = true;
    // disable locked limits
    lockLimitButton.disabled = true;
    // re-render charts
    redraw("applyTrends");

    trendRemoveButtons.forEach(showElement);
    hideModal("trend-dialog", "trend-backdrop")
  });
  trendRemoveButtons?.forEach((b) =>
    b.addEventListener("click", () => {
      state.isShowingTrendLines = false;

      trendRemoveButtons.forEach(hideElement);

      hideModal("trend-dialog", "trend-backdrop");

      // enable seasonality
      seasonalFactorsButton.disabled = false;
      // enable locked limits
      lockLimitButton.disabled = false;

      resetTrendDialogTable();
      redraw("removeTrends");
    })
  );
  trendDialogResetButton?.addEventListener("click", () => {
    resetTrendDialogTable();
    if (state.isShowingTrendLines) {
      redraw("resetTrends");
    }
  });

  trendSlopeInput?.addEventListener("change", () => {
    console.assert(!isNaN(parseFloat(trendSlopeInput.value)));

    if (trendSlopeInput.value !== "") {
      state.regressionStats.m = parseFloat(trendSlopeInput.value);
    }
  });
  trendInterceptInput?.addEventListener("change", () => {
    console.assert(!isNaN(parseFloat(trendInterceptInput.value)));

    if (trendInterceptInput.value !== "") {
      state.regressionStats.c = parseFloat(trendInterceptInput.value);
    }
  });

  // Lock Limits
  const lockLimitWarningLabel = document.querySelector(
    "#lock-limit-warning"
  ) as HTMLParagraphElement;

  const lockLimitDialog = document.querySelector(
    "#lock-limit-dialog"
  ) as HTMLDialogElement;
  const lockLimitDialogCloseButton = document.querySelector(
    "#lock-limit-close"
  ) as HTMLButtonElement;
  const lockLimitDialogAddButton = document.querySelector(
    "#lock-limit-add"
  ) as HTMLButtonElement;
  const lockLimitDialogResetButton = document.querySelector(
    "#lock-limit-reset-data"
  ) as HTMLButtonElement;

  lockLimitDialogResetButton?.addEventListener(
    "click",
    () => {
      resetLockedLimitDialogTable();
      if (isLockedLimitsActive()) {
        lockLimits();
      }
    }
  );

  // If the initial state has locked limits, we should show the buttons and warnings
  if (isLockedLimitsActive()) {
    document.querySelectorAll(".lock-limit-remove").forEach(showElement);
    showElement(lockLimitWarningLabel);
    // enable locked limits and disable trend
    lockLimitButton.disabled = false;
    trendDialogOpenButton.disabled = true;
  }

  lockLimitButton?.addEventListener("click", (e) => {
    if (!isLockedLimitsActive()) {
      resetLockedLimitDialogTable();
    }
    setLockedLimitInputs(!isLockedLimitsActive());
    showModal("lock-limit-dialog", "lock-limit-backdrop");
  });

  lockLimitDialogCloseButton?.addEventListener("click", (e) => {
    hideModal("lock-limit-dialog", "lock-limit-backdrop");
  });

  // Set up all lock-limit-remove buttons
  document.querySelectorAll(".lock-limit-remove")?.forEach((d) =>
    d.addEventListener("click", () => {
      hideElement(lockLimitWarningLabel);
      document.querySelectorAll(".lock-limit-remove").forEach(hideElement);
      hideModal("lock-limit-dialog", "lock-limit-backdrop");

      // enable trend limits
      trendDialogOpenButton.disabled = false;

      state.lockedLimitStatus &= ~LockedLimitStatus.LOCKED; // set to unlocked
      redraw("removeLockLimit");
    })
  );

  lockLimitDialogAddButton?.addEventListener("click", () => {
    lockLimits();
    hideModal("lock-limit-dialog", "lock-limit-backdrop");
    // show lock-limit-remove button
    document.querySelectorAll(".lock-limit-remove").forEach(showElement);
    showElement(lockLimitWarningLabel);
    lockLimitWarningLabel.classList.remove("hidden");
    // disable trend limits
    trendDialogOpenButton.disabled = true;
  });

  // Data table
  initialiseHandsOnTable();

  // CSV data management
  const csvFile = document.getElementById("csv-file") as HTMLInputElement;
  csvFile?.addEventListener("change", (e) => {
    // check if there is a file input
    if (!csvFile.files?.length) {
      console.log("No file input");
      return;
    }
    const input = csvFile!.files[0];
    const reader = new FileReader();
    reader.addEventListener("loadend", () => {
      // parse into array
      const text = reader.result as string;
      // if not passed test, display error (inside function run) and return
      let { passed, multiplier, xLabel, yLabel, xdata } =
        csvTestingParser(text);
      if (!passed) {
        return;
      }
      console.log(passed, multiplier, xLabel, yLabel, xdata);
      // else handle multiplier (manipulate data and labels)
      const superscript = "⁰¹²³⁴⁵⁶⁷⁸⁹";
      function formatPower(d: number) {
        return d
          .toString()
          .split("")
          .map(function (c) {
            return superscript[Number(c)];
          })
          .join("");
      }
      if (multiplier > 0) {
        yLabel += ` (x10${formatPower(multiplier)})`;
        for (let i = 0; i < xdata.length; i++) {
          xdata[i].value /= 10 ** multiplier;
        }
      }

      // hide error message
      const errorMsg = document.getElementById("file-error") as HTMLDivElement;
      errorMsg.style.display = "none";
      errorMsg.innerText = "";
      // UPDATE STATE
      // sortX(xdata)
      state.xLabel = xLabel;
      state.yLabel = yLabel;
      state.tableData = xdata;
      state.lockedLimitBaseData = deepClone(state.tableData);
      hot.updateSettings({
        data: xdata,
        colHeaders: [xLabel, yLabel],
      });
      // Refreshes divider state when we upload new data
      state.dividerLines = [];

      renderCharts("loadCsv");
    });

    // Reads the csv file as string, after which it emits the loadend event
    reader.readAsText(input);
  });

  // Other button listeners
  const downloadDataButton = document.querySelector(
    "#download-data"
  ) as HTMLButtonElement;
  const refreshChartsButton = document.querySelector(
    "#refresh-charts"
  ) as HTMLButtonElement;
  const shareLinkButton = document.querySelector(
    "#share-link"
  ) as HTMLButtonElement;

  document.querySelector("#export-xplot")?.addEventListener("click", () => {
    exportCanvasWithBackground('xplot');
  });
  document.querySelector("#export-mrplot")?.addEventListener("click", () => {
    exportCanvasWithBackground('mrplot');
  });

  downloadDataButton?.addEventListener("click", () => {
    downloadCSV(state, isLockedLimitsActive());
  });

  refreshChartsButton?.addEventListener("click", () => {
    renderCharts("refreshCharts");
  });
  shareLinkButton?.addEventListener("click", () => {
    let link = generateShareLink(state);
    // https://stackoverflow.com/questions/417142/what-is-the-maximum-length-of-a-url-in-different-browsers
    if (link.length > MAX_LINK_LENGTH) {
      alert("Link too long! Consider removing a few datapoints.");
      return;
    }

    // copy to clipboard
    navigator.clipboard.writeText(link).then(
      function () {
        console.log("Async: Copying to clipboard was successful!");
      },
      function (err) {
        console.error("Async: Could not copy text: ", err);
      }
    );

    // toggle message on share button click
    const dataCopiedMessageLabel = document.getElementById(
      "data-copied-msg"
    ) as HTMLDivElement;
    hideElement(shareLinkButton);
    showElement(dataCopiedMessageLabel);
    setTimeout(() => {
      hideElement(dataCopiedMessageLabel);
      showElement(shareLinkButton);
    }, 1500);
  });

  // Resize charts
  const x = document.querySelector("#xplot");
  const mr = document.querySelector("#mrplot");

  const createResizeObserverFor = (plot: EChartsType) =>
    new ResizeObserver((entries) => {
      if (entries[0]) {
        const { width, height } = entries[0].contentRect;
        plot.resize({
          width,
          height,
          silent: true,
          animation: {
            duration: 500,
          },
        });
        window.redrawEcharts();
      }
    });

  createResizeObserverFor(xplot).observe(x);
  createResizeObserverFor(mrplot).observe(mr);
});

function calculateLimits(xdata: DataValue[]): Partial<LineValueType> {
  const movements = getMovements(xdata);
  // since avgX and avgMovement is used for further calculation, we only round it after calculating unpl, lnpl, url
  const avgX = USE_MEDIAN_AVG
    ? calculateMedian(xdata.map((x) => x.value))
    : xdata.reduce((a, b) => a + b.value, 0) / xdata.length;
  // filteredMovements might be empty
  const avgMovement = USE_MEDIAN_MR
    ? calculateMedian(movements.map((x) => x.value))
    : movements.reduce((a, b) => a + b.value, 0) /
    Math.max(movements.length, 1);
  const delta =
    (USE_MEDIAN_MR
      ? MEDIAN_NPL_SCALING
      : NPL_SCALING) * avgMovement;
  const UNPL = avgX + delta;
  const LNPL = avgX - delta;
  const URL =
    (USE_MEDIAN_MR
      ? MEDIAN_URL_SCALING
      : URL_SCALING) * avgMovement;
  const lowerQuartile = (LNPL + avgX) / 2;
  const upperQuartile = (UNPL + avgX) / 2;
  return {
    avgX: round(avgX),
    avgMovement: round(avgMovement),
    UNPL: round(UNPL),
    LNPL: round(LNPL),
    URL: round(URL),
    lowerQuartile: round(lowerQuartile),
    upperQuartile: round(upperQuartile),
  };
}

function calculateMedian(arr: number[]): number {
  // Sort the array in ascending order
  arr.sort((a, b) => a - b);

  const n = arr.length;
  if (n === 0) return 0;
  if (n % 2 !== 0) {
    // If odd, return the middle element
    const middleIndex = Math.floor(n / 2);
    return arr[middleIndex];
  } else {
    // If even, return the average of the two middle elements
    const middleIndex1 = n / 2 - 1;
    const middleIndex2 = n / 2;
    return (arr[middleIndex1] + arr[middleIndex2]) / 2;
  }
}

// This function is a 'no-divider' version of the calculation of the limits (specifically for the locked limits)
function calculateLockedLimits() {
  let xdata = state.lockedLimitBaseData.filter(
    (dv) => dv.x && (dv.value || dv.value == 0)
  );
  return calculateLimits(xdata);
}

/**
 * Set the value, placeholder and style of lockedlimit inputs
 * @param updateInputValue whether to update the value of the inputs
 */
function setLockedLimitInputs(updateInputValue: boolean) {
  let lv = calculateLockedLimits(); // calculate locked limits from the table
  document
    .querySelectorAll(".lock-limit-input")
    .forEach((el: HTMLInputElement) => {
      el.value = updateInputValue
        ? lv[el.dataset.limit]
        : state.lockedLimits[el.dataset.limit];
    });
  // if we update the values of the inputs, we reset all modified status
  if (updateInputValue) {
    state.lockedLimitStatus &= ~LockedLimitStatus.AVGX_MODIFIED;
    state.lockedLimitStatus &= ~LockedLimitStatus.LNPL_MODIFIED;
    state.lockedLimitStatus &= ~LockedLimitStatus.UNPL_MODIFIED;
  }

  // set color and placeholders
  (
    document.querySelector(
      '.lock-limit-input[data-limit="avgX"]'
    ) as HTMLElement
  ).style["color"] = isAvgXModified(state.lockedLimitStatus)
      ? "rgb(220 38 38)"
      : "black";
  (
    document.querySelector(
      '.lock-limit-input[data-limit="avgX"]'
    ) as HTMLInputElement
  ).placeholder = `${lv.avgX}`;
  (
    document.querySelector(
      '.lock-limit-input[data-limit="UNPL"]'
    ) as HTMLElement
  ).style["color"] = isUnplModified(state.lockedLimitStatus)
      ? "rgb(220 38 38)"
      : "black";
  (
    document.querySelector(
      '.lock-limit-input[data-limit="UNPL"]'
    ) as HTMLInputElement
  ).placeholder = `${lv.UNPL}`;
  (
    document.querySelector(
      '.lock-limit-input[data-limit="LNPL"]'
    ) as HTMLElement
  ).style["color"] = isLnplModified(state.lockedLimitStatus)
      ? "rgb(220 38 38)"
      : "black";
  (
    document.querySelector(
      '.lock-limit-input[data-limit="LNPL"]'
    ) as HTMLInputElement
  ).placeholder = `${lv.LNPL}`;
  (
    document.querySelector(
      '.lock-limit-input[data-limit="avgMovement"]'
    ) as HTMLInputElement
  ).placeholder = `${lv.avgMovement}`;
  (
    document.querySelector(
      '.lock-limit-input[data-limit="URL"]'
    ) as HTMLInputElement
  ).placeholder = `${lv.URL}`;
}

// An implementation of excel drag-to-extend-series for handsontable
function autofillTable(selectionData, sourceRange, targetRange, direction) {
  if (sourceRange.from.col == 1) {
    // use default behaviour if its the data column
    // if targetRange is larger than source range, the selected data will be repeated.
    return selectionData;
  }

  // most likely, the user is trying to extend the current pattern.
  // if there is a pattern, we will try to extend the pattern
  // otherwise we will just use default behaviour.
  let dateArray = selectionData.map((x) => new Date(fromDateStr(x[0])));
  let difference = 86400000; // one day
  if (dateArray.length >= 2) {
    difference = dateArray[1].valueOf() - dateArray[0].valueOf();
    for (let i = 2; i < dateArray.length; i++) {
      if (dateArray[i].valueOf() - dateArray[i - 1].valueOf() != difference) {
        // pattern is broken
        return selectionData;
      }
    }
  }

  let result = [];
  // If only one row is selected, we increment by one day only
  if (dateArray.length < 2) {
    if (direction == "down") {
      result = dateArray.map(
        (x) => new Date(new Date(x.valueOf()).setDate(x.getDate() + 1))
      );
      for (
        let i = result.length;
        i <= targetRange.to.row - targetRange.from.row;
        i++
      ) {
        let cd = result[result.length - 1];
        result.push(new Date(new Date(cd.valueOf()).setDate(cd.getDate() + 1)));
      }
    } else if (direction == "up") {
      result = dateArray.map(
        (x) => new Date(new Date(x.valueOf()).setDate(x.getDate() - 1))
      );
      for (
        let i = result.length;
        i <= targetRange.to.row - targetRange.from.row;
        i++
      ) {
        let cd = result[0];
        // much more expensive than downwards extension, but user should rarely hit this...
        result = [
          new Date(new Date(cd.valueOf()).setDate(cd.getDate() - 1)),
          ...result,
        ];
      }
    }
    return result.map((x) => [toDateStr(x)]);
  }

  // If more than one row was selected we increment by computed difference
  // Note that this doesn't handle daylight savings time, so two days out of
  // every year this will give wrong results.
  if (direction == "down") {
    result = dateArray.map(
      (x) => new Date(x.valueOf() + dateArray.length * difference)
    );
    for (
      let i = result.length;
      i <= targetRange.to.row - targetRange.from.row;
      i++
    ) {
      result.push(new Date(result[result.length - 1].valueOf() + difference));
    }
  } else if (direction == "up") {
    result = dateArray.map(
      (x) => new Date(x.valueOf() - dateArray.length * difference)
    );
    for (
      let i = result.length;
      i <= targetRange.to.row - targetRange.from.row;
      i++
    ) {
      // much more expensive than downwards extension, but user should rarely hit this...
      result = [new Date(result[0].valueOf() - difference), ...result];
    }
  }
  return result.map((x) => [toDateStr(x)]);
}

// Modifying paste behaviour
function beforePasteTable(data, coords) {
  if (coords.length > 1) return;
  let coord = coords[0];
  // if selected area has more rows than the data, we add empty '' to the end of data to simulate clearing this rows
  // This change handles the common edge case where user copy, say from other sources, and paste a smaller set of data
  // than the one currently exists in the table.
  // The default behaviour of the CopyPaste plugin is to repeat the data until the end of the selected area,
  // which causes wrong calculation (and crossing lines in the chart) due to repeated data at the end.
  for (let i = coord.startRow + data.length; i <= coord.endRow; i++) {
    // endRow is inclusive.
    data.push(["", ""]);
  }
  return;
}

// Check duplicate in the x value (first column) of the table and set background color accordingly
function checkDuplicatesInTable(arr: string[], tableInstance) {
  let seen = {};
  let isDuplicateDetected = false;
  arr.forEach((el, idx) => {
    if (!el) return; // null or ""
    if (el in seen) {
      tableInstance.setCellMeta(idx, 0, "className", "duplicate");
      isDuplicateDetected = true;
    } else {
      tableInstance.setCellMeta(idx, 0, "className", "");
    }
    seen[el] = true;
  });

  if (isDuplicateDetected) {
    document
      .getElementById("duplicate-data-warning")
      .classList.remove("hidden");
  } else {
    document.getElementById("duplicate-data-warning").classList.add("hidden");
  }
}

/**
 * Returns true if the upper limit lines and lower limit lines are symmetric w.r.t. average line
 * Otherwise, we check which side's limit has been changed and disable the quartile line for that side
 */
function shouldUseQuartile() {
  function isSymmetric(avg: number, unpl: number, lnpl: number) {
    return Math.abs(unpl + lnpl - 2 * avg) < 0.001;
  }

  if ((state.lockedLimitStatus & ~1) == 0) {
    // not modified at all
    return { useUpperQuartile: true, useLowerQuartile: true };
  }

  if (
    (isLnplModified(state.lockedLimitStatus) &&
      isUnplModified(state.lockedLimitStatus)) ||
    isAvgXModified(state.lockedLimitStatus)
  ) {
    return isSymmetric(
      state.lockedLimits.avgX,
      state.lockedLimits.UNPL,
      state.lockedLimits.LNPL
    )
      ? { useUpperQuartile: true, useLowerQuartile: true }
      : { useUpperQuartile: false, useLowerQuartile: false };
  }

  if (isUnplModified(state.lockedLimitStatus)) {
    return { useUpperQuartile: false, useLowerQuartile: true };
  }
  if (isLnplModified(state.lockedLimitStatus)) {
    return { useUpperQuartile: true, useLowerQuartile: false };
  }

  console.assert(false); // should not reach here
  return { useUpperQuartile: true, useLowerQuartile: true };
}

function isLockedLimitsActive(): boolean {
  return (state.lockedLimitStatus & LockedLimitStatus.LOCKED) == 1;
}

function renderCharts(cause: string = "init") {
  // setup initial data
  updateInPlace(
    state.xdata,
    state.tableData.filter((dv) => dv.x && (dv.value || dv.value == 0))
  );
  updateInPlace(state.movements, getMovements(state.xdata));

  state.dividerLines = state.dividerLines
    .filter((dl) => !isShadowDividerLine(dl)) // filter out border "shadow" dividers
    .concat([
      { id: "divider-start", x: 0 },
      {
        id: "divider-end",
        x: Infinity,
      },
    ]);
  state.trendLines = createTrendlines(state.regressionStats, state.xdata);
  redraw(cause);
}

/**
 * Find the extremes on the x-axis
 * @param arr
 * @returns
 */
function findExtremesX(arr: DataValue[]) {
  let min = Infinity;
  let max = -Infinity;
  arr.forEach((el) => {
    let d = fromDateStr(el.x);
    min = Math.min(min, d);
    max = Math.max(max, d);
  });
  return { min, max };
}

function doEChartsThings(stats: _Stats) {
  initialiseECharts(true);
  renderStats(stats);
  renderLimitLines(stats);
  adjustChartAxis(stats);
  renderChartGraphics(stats);
}

/**
 * Sets up base chart options
 * @param shouldReplaceState: if set to true, the previous state of the charts will be discarded.
 */
function initialiseECharts(shouldReplaceState: boolean = false) {
  xplot.setOption({ ...chartBaseOptions }, shouldReplaceState);
  mrplot.setOption({ ...chartBaseOptions }, shouldReplaceState);
  xplot.setOption({
    title: {
      text:
        (state.isShowingDeseasonalisedData ? "Deseasonalised " : " ") +
        "X Plot" +
        (state.yLabel.toLowerCase() !== "value" ? `: ${state.yLabel}` : ""),
    },
    xAxis: {
      name: state.xLabel,
    },
    yAxis: {
      name: state.yLabel,
    },
  });
  mrplot.setOption({
    title: {
      text:
        (state.isShowingDeseasonalisedData ? "Deseasonalised " : " ") +
        "MR Plot" +
        (state.yLabel.toLowerCase() !== "value" ? `: ${state.yLabel}` : ""),
    },
    xAxis: {
      name: state.xLabel,
    },
    yAxis: {
      name: state.yLabel,
    },
  });
}

/**
 * Draws the main data points into charts
 * @param stats
 */
function renderStats(stats: _Stats) {
  // x series
  let xSeries = stats.xdataPerRange.map(mapDataValuesToChartSeries);

  if (state.isShowingTrendLines) {
    const firstDividerX = Math.min(
      ...state.dividerLines.slice(1).map((l) => l.x)
    );
    const transformTrendLine = (dv: DataValue[]) => {
      if (state.dividerLines.length <= 2) return dv;
      const fil = dv.filter((d) => fromDateStr(d.x) < firstDividerX);
      return fil;
    };
    xSeries = xSeries.concat(
      createTrendValueSeries({
        id: "center",
        color: MEAN_SHAPE_COLOR,
        data: transformTrendLine(state.trendLines.centreLine),
      })
    );
    xSeries = xSeries.concat(
      createTrendValueSeries({
        id: "unpl",
        color: LIMIT_SHAPE_COLOR,
        data: transformTrendLine(state.trendLines.unpl),
      })
    );
    xSeries = xSeries.concat(
      createTrendValueSeries({
        id: "lnpl",
        color: LIMIT_SHAPE_COLOR,
        data: transformTrendLine(state.trendLines.lnpl),
      })
    );
    xSeries = xSeries.concat(
      createTrendValueSeries({
        id: "lowerQtl",
        color: "grey",
        data: transformTrendLine(state.trendLines.lowerQtl),
        lineType: "dashed",
      })
    );
    xSeries = xSeries.concat(
      createTrendValueSeries({
        id: "upperQtl",
        color: "grey",
        data: transformTrendLine(state.trendLines.upperQtl),
        lineType: "dashed",
      })
    );
  }

  xplot.setOption({
    series: xSeries,
  });

  mrplot.setOption({
    series: stats.movementsPerRange.map(mapDataValuesToChartSeries),
  });
}

function getXChartYAxisMinMax(stats: _Stats) {
  let min = stats.xchartMin;
  let max = stats.xchartMax;
  // Handle edge case where min equals max
  if (min === max) {
    const value = min;
    if (value === 0) return [-0.5, 0.5];
    // For non-zero values, extend 50% in both directions
    const padding = Math.abs(value * 0.5);
    return [value - padding, value + padding];
  }

  // Calculate the range
  const range = max - min;
  const absMax = Math.max(Math.abs(min), Math.abs(max));

  // Calculate the order of magnitude based on the range and absolute maximum
  const orderOfMagnitude = Math.floor(Math.log10(Math.max(range, absMax)));
  const scale = Math.pow(10, orderOfMagnitude);

  // Calculate padding based on range characteristics
  let padding;
  if (range / scale < 0.25) {
    // For very small ranges relative to their scale, use more padding
    padding = range * 0.5;
  } else if (range / scale < 1) {
    // For small ranges, use moderate padding
    padding = range * 0.25;
  } else {
    // For larger ranges, use proportional padding
    padding = range * 0.1;
  }

  // Ensure minimum padding based on scale
  padding = Math.max(padding, scale / 100);

  let lowerBound = min - padding;
  let upperBound = max + padding;

  // Round to nice numbers based on scale
  const roundingScale = scale / 100;
  lowerBound = Math.floor(lowerBound / roundingScale) * roundingScale;
  upperBound = Math.ceil(upperBound / roundingScale) * roundingScale;

  // Special handling for numbers close to zero
  if (Math.abs(lowerBound) < scale / 1000) lowerBound = 0;
  if (Math.abs(upperBound) < scale / 1000) upperBound = 0;

  return [lowerBound, upperBound];
}

/**
 * Adjusts the charts min and max values to ensure all points visible
 * @param stats
 */
function adjustChartAxis(stats: _Stats) {
  const xMin = Math.min(
    ...stats.xdataPerRange.map((range) =>
      Math.min(...range.map((d) => dayjs(d.x).valueOf()))
    )
  );

  let xMax = Math.max(
    ...stats.movementsPerRange.map((range) =>
      Math.max(...range.map((d) => dayjs(d.x).valueOf()))
    ),
    ...stats.xdataPerRange.map((range) =>
      Math.max(...range.map((d) => dayjs(d.x).valueOf()))
    )
  );
  xMax = dayjs(xMax).add(2, "day").valueOf();

  const [xplotYMin, xplotYMax] = getXChartYAxisMinMax(stats);

  xplot.setOption({
    yAxis: {
      min: xplotYMin,
      max: xplotYMax,
    },
    xAxis: { min: xMin, max: xMax },
  });
  mrplot.setOption({
    xAxis: { min: xMin, max: xMax },
    yAxis: { min: 0, max: stats.mrchartMax },
  });
}

function renderChartGraphics(stats: _Stats) {
  state.dividerLines.forEach((dl) => renderDividerLine(dl, stats));
}

function promptNewColumnName() {
  let newColName = prompt("Insert a new column name", state.yLabel);
  if (newColName) {
    let colHeaders = hot.getColHeader();
    colHeaders[1] = newColName;
    state.yLabel = newColName;
    hot.updateSettings({
      colHeaders,
    });
    redraw("editChartTitle");
  }
}

function updateSeasonalFactors() {
  if (state.tableData.length === 0) {
    return; // data has not been loaded yet
  }
  const [sf, hasMissingPeriods] = calculateSeasonalFactors(
    removeAllNulls(state.tableData), // use tableData as a basis
    removeAllNulls(state.seasonalFactorData),
    state.deSeasonalisePeriod,
    state.deSeasonaliseGrouping
  );
  state.seasonalFactorTableData = sf.map((d) => [d]);
  state.hasMissingPeriods = hasMissingPeriods;

  if (state.hasMissingPeriods) {
    showElement(document.querySelector("#deseason-warn-3"));
  } else {
    hideElement(document.querySelector("#deseason-warn-3"));
  }

  if (seasonalFactorsHot != null) {
    seasonalFactorsHot.updateData(state.seasonalFactorTableData);
    updateDTableSeasonalFactorMapping();
  }
}

function updateDeseasonalisedData() {
  if (state.xdata.length === 0) {
    return;
  }
  state.deSeasonalisedData = applySeasonalFactors(
    state.xdata,
    state.seasonalFactorTableData.map((r) => r[0]),
    state.deSeasonaliseGrouping,
    state.deSeasonalisePeriod
  );

  updateDTableSeasonalFactorMapping();
}

function resetDeseasonalisedDialogTable() {
  state.seasonalFactorData = deepClone(state.tableData);
  deseasonaliseHot.updateSettings({
    data: state.seasonalFactorData,
  });
  updateDTableSeasonalFactorMapping();
  updateSeasonalFactors();
}

function resetLockedLimitDialogTable() {
  // let the locked limit data reflect the latest x data if locked limits are not currently active
  updateInPlace(state.lockedLimitBaseData, state.xdata);
  lockedLimitHot.updateSettings({
    data: state.lockedLimitBaseData,
    colHeaders: [state.xLabel, state.yLabel],
  });
}

function resetTrendDialogTable() {
  state.trendData = deepClone(state.tableData);
  trendHot.updateData(state.trendData);
  state.trendLines = createTrendlines(state.regressionStats, state.trendData);
}

function calculateTrendStats() {
  const trendSlopeInput = document.querySelector(
    "#trend-slope-input"
  ) as HTMLInputElement;
  const trendInterceptInput = document.querySelector(
    "#trend-intercept-input"
  ) as HTMLInputElement;

  // compute trendlines based on data
  const rStats = calculateRegressionStats(removeAllNulls(state.trendData));
  state.regressionStats = rStats;
  if (trendSlopeInput) {
    trendSlopeInput.value = round(state.regressionStats.m, 6).toString();
    trendSlopeInput.placeholder = round(state.regressionStats.m, 6).toString();
  }
  if (trendInterceptInput) {
    trendInterceptInput.value = round(state.regressionStats.c, 6).toString();
    trendInterceptInput.placeholder = round(state.regressionStats.c, 6).toString();
  }
}

/* Modals */
function showModal(dialogId: string, backdropId: string) {
  const dialog = document.getElementById(dialogId);
  const backdrop = document.getElementById(backdropId);

  if (!dialog || !backdrop) return;

  dialog.classList.add('active');
  backdrop.classList.add('active');

  // Prevent body scroll when modal is open
  document.body.style.overflow = 'hidden';
}

function hideModal(dialogId: string, backdropId: string) {
  const dialog = document.getElementById(dialogId);
  const backdrop = document.getElementById(backdropId);

  if (!dialog || !backdrop) return;

  dialog.classList.remove('active');
  backdrop.classList.remove('active');

  // Restore body scroll
  document.body.style.overflow = '';
}

// Initialize modal functionality
function initializeModal(dialogId: string, backdropId: string, closeButtonId: string) {
  const dialog = document.getElementById(dialogId);
  const backdrop = document.getElementById(backdropId);
  const closeButton = document.getElementById(closeButtonId);

  if (!dialog || !backdrop || !closeButton) return;

  // Close on backdrop click
  backdrop.addEventListener('click', () => {
    hideModal(dialogId, backdropId);
  });

  // Close on close button click
  closeButton.addEventListener('click', () => {
    hideModal(dialogId, backdropId);
  });

  // Close on escape key
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && dialog.classList.contains('active')) {
      hideModal(dialogId, backdropId);
    }
  });
}

/* Echarts */
const ECHARTS_DATE_FORMAT = "{d} {MMM}";
const defaultValueFormatter = (n: number) => numberStringSpaced(round(n));
const backgroundImage = new Image();
backgroundImage.src = "../xmrit-bg.png";

const chartBaseOptions = {
  backgroundColor: "rgba(255, 255, 255, 0.5)",
  xAxis: {
    type: "time",
    axisLabel: {
      formatter: ECHARTS_DATE_FORMAT,
      hideOverlap: true,
      color: "#000",
    },
    axisLine: {
      lineStyle: {
        color: "#000",
      },
      onZero: false,
    },
    position: "bottom",
    nameLocation: "center",
    nameGap: 25,
  },
  yAxis: {
    splitLine: {
      show: false,
    },
    axisLabel: {
      fontSize: 11,
      color: "#000",
      hideOverlap: true,
      formatter: defaultValueFormatter,
      padding: [0, 10, 0, 0],
    },
    splitNumber: 6,
    nameLocation: "middle",
    nameRotate: 90,
    nameGap: 45,
    nameTextStyle: {
      color: "#000",
    },
  },
  title: {
    left: "center",
    triggerEvent: true,
  },
  tooltip: {
    show: true,
  },
};

const mapDataValueToChartDataPoint =
  ({ showLabel }: { showLabel: boolean }) =>
    (dv: DataValue) => ({
      value: [fromDateStr(dv.x), dv.value],
      itemStyle: {
        color: dataStatusColor(dv.status),
      },
      label: {
        show: showLabel,
        color: dataLabelsStatusColor(dv.status),
        fontWeight: "bold",
        formatter: (params) => {
          return defaultValueFormatter(params.data.value[1]);
        },
      },
      tooltip: {
        formatter: `${dayjs(dv.x).format("ddd, D MMM YYYY")}:<br/> ${dv.value}`,
      },
    });

const mapDataValuesToChartSeries = (subD, i) => ({
  name: `${i}-data`,
  z: 100,
  type: "line",
  symbol: "circle",
  symbolSize: 7,
  lineStyle: {
    color: "#000",
  },
  labelLayout: {
    hideOverlap: true,
  },
  data: subD.map(mapDataValueToChartDataPoint({ showLabel: true })),
});

const createTrendValueSeries = ({
  id,
  color,
  data,
  lineType = "solid",
}: {
  id: string;
  color: string;
  data: DataValue[];
  lineType?: "solid" | "dashed";
}) => ({
  name: `trend-${id}`,
  z: 50,
  type: "line",
  symbol: "none",
  symbolSize: 0,
  lineStyle: {
    color,
    width: lineType === "solid" ? 3 : 1,
    type: lineType,
  },
  labelLayout: {
    hideOverlap: true,
  },
  data: data.map(mapDataValueToChartDataPoint({ showLabel: false })),
});

/* Sharelink */

/**
 * Sharelink v0:
 * - v: 0
 * - d: lz77.compress(xLabel,yLabel,date-cols...).base64(float32array of value-cols).
 * - l: float32array of [avgX, avgMovement, LNPL, UNPL, URL, lockedLimitStatus]
 * - s: float32array of dividers x values (unix timestamp in milliseconds)
 * @returns
 */
function generateShareLink(state: AppState): string {
  let currentUrlParams = extractDataFromUrl();
  const paramsObj = {
    v: currentUrlParams.get("v") ?? "0", // get version from url or default to 0
  };

  if (paramsObj["v"] == "1") {
    // in sharelink version 1, we assume users don't intend to change their src data so we just copy the user's url here
    // Otherwise, they should change the result returned from the remote url.
    paramsObj["d"] = currentUrlParams.get("d");
  } else {
    let validXdata = state.tableData.filter((dv) => dv.x || dv.value);
    // basically first 2 are labels, followed by date-column, followed by value-column.
    // date-column are compressed using lz77
    // value-column are encoded as bytearray and converted into base64 string
    const dateText =
      `${state.xLabel.replace(",", ";")},${state.yLabel.replace(",", ";")},` +
      validXdata
        .map((d) => {
          if (d.x) {
            // hack[1]: if dates contain , e.g. "Jan 1, 2020", we replace the comma using ; instead
            return d.x.replace(",", ";");
          } else {
            return "";
          }
        })
        .join(",");
    const valueText = validXdata.map((d) => round(d.value));
    paramsObj["d"] =
      btoaUrlSafe(lz77.compress(dateText)) +
      "." +
      encodeNumberArrayString(valueText);
  }

  if (state.dividerLines) {
    const dividers = encodeNumberArrayString(
      state.dividerLines
        .filter((dl) => !isShadowDividerLine(dl))
        .map((dl) => dl.x)
    );
    if (dividers.length > 0) {
      paramsObj["s"] = dividers;
    }
  }
  if (
    state.lockedLimits &&
    (state.lockedLimitStatus & LockedLimitStatus.LOCKED) == 1
  ) {
    // IMPORTANT: If you change the number array below in any way except appending to it, please bump up the version number ('v') as it will break all existing links
    // and update the decoding logic below
    paramsObj["l"] = encodeNumberArrayString([
      state.lockedLimits.avgX,
      state.lockedLimits.avgMovement,
      state.lockedLimits.LNPL,
      state.lockedLimits.UNPL,
      state.lockedLimits.URL,
      state.lockedLimitStatus,
    ]);
  }
  if (state.isShowingDeseasonalisedData) {
    // IMPORTANT: If you change the number array below in any way except appending to it, please bump up the version number ('v') as it will break all existing links
    // and update the decoding logic below
    paramsObj["p"] = encodeNumberArrayString(
      state.seasonalFactorTableData.map((a) => a[0])
    ); // seasonal factors to apply
    paramsObj["p0"] = state.deSeasonalisePeriod; // deseason period
  }
  if (state.isShowingTrendLines) {
    // IMPORTANT: If you change the number array below in any way except appending to it, please bump up the version number ('v') as it will break all existing links
    // and update the decoding logic below
    paramsObj["t"] = encodeNumberArrayString([
      state.regressionStats.m,
      state.regressionStats.c,
      state.regressionStats.avgMR,
    ]); // encode trend stats in order
  }
  const pageParams = new URLSearchParams(paramsObj);
  const fullPath = `${window.location.origin}${window.location.pathname
    }#${pageParams.toString()}`;
  return fullPath;
}

async function decodeShareLink(
  version: string,
  d: string,
  dividers: string,
  limits: string,
  sf: string,
  trStats: string
) {
  let data = [];
  let labels = [];
  if (version == "1") {
    // version 1. d is a remote url. We fetch labels and data from a remote url.
    let res = await fetch(d);
    let body = await res.json();
    labels = [body.xLabel ?? "Date", body.yLabel ?? "Value"];
    data = body.xdata.map((v, idx) => {
      return {
        order: idx,
        x: v["x"], // Date in YYYY-MM-DD format
        value: round(+v["value"]),
        status: DataStatus.NORMAL,
      };
    });
  } else {
    // version 0, the oldest. d is an encoded table
    d = decodeURIComponent(d);
    dividers = decodeURIComponent(dividers);
    let [datePart, valuePart] = d.split(".", 2);
    let dateText = lz77.decompress(atobUrlSafe(datePart));
    let values = decodeNumberArrayString(valuePart);
    // hack[1]: reverse the , -> ; encoding
    let dates = dateText.split(",").map((d) => d.replace(";", ","));
    labels = dates.splice(0, 2);
    // sort dates based on string
    // dates.sort((a, b) => Date.parse(a) < Date.parse(b) ? -1 : 1);
    data = dates.map((d, idx) => {
      return {
        order: idx,
        x: d, // Date in YYYY-MM-DD format
        value: round(values[idx]),
        status: DataStatus.NORMAL,
      };
    });
  }

  let dividerLines = [] as DividerType[];
  if (dividers?.length > 0) {
    dividerLines = dividerLines.concat(
      decodeNumberArrayString(dividers).map((x, i) => {
        return {
          id: `divider-${i + 1}`,
          x: x,
        };
      })
    );
  }
  let lockedLimits = structuredClone(INACTIVE_LOCKED_LIMITS);
  let lockedLimitStatus = LockedLimitStatus.UNLOCKED;
  if (limits?.length > 0) {
    const [avgX, avgMovement, LNPL, UNPL, URL, ...rest] =
      decodeNumberArrayString(limits).map(round);
    if (rest.length > 0) {
      lockedLimitStatus = rest[0];
    }
    lockedLimits = {
      avgX,
      avgMovement,
      LNPL,
      UNPL,
      URL,
      lowerQuartile: (avgX + LNPL) / 2,
      upperQuartile: (avgX + UNPL) / 2,
    } as LineValueType;
  }
  let seasonalFactors;
  if (sf.length > 0) {
    seasonalFactors = decodeNumberArrayString(sf);
  }

  let regressionStats: RegressionStats;
  if (trStats.length > 0) {
    const [m, c, avgMR] = decodeNumberArrayString(trStats);
    regressionStats = { m, c, avgMR };
  }

  return {
    xLabel: labels[0].replace(";", ","),
    yLabel: labels[1].replace(";", ","),
    data,
    dividerLines,
    lockedLimits,
    lockedLimitStatus,
    seasonalFactors,
    regressionStats,
  };
}

async function exportCanvasWithBackground(id: string) {  
  // Get the source canvas
  const container = document.getElementById(id);
  if (!container) return;
  const sourceCanvas = container.querySelector('canvas');
  if (!sourceCanvas) return;

  // Create a new canvas
  const finalCanvas = document.createElement('canvas');
  finalCanvas.width = sourceCanvas.width;
  finalCanvas.height = sourceCanvas.height;
  console.log(finalCanvas.width, finalCanvas.height)
  const ctx = finalCanvas.getContext('2d');
  if (!ctx) return;

  // Load and draw background image
  const bgImage = new Image();
  bgImage.src = '/xmrit-bg.png';
  
  try {
    await new globalThis.Promise((resolve, reject) => {
      bgImage.onload = resolve;
      bgImage.onerror = reject;
    });

    // Draw background (scaled to fit)
    ctx.drawImage(bgImage, 0, 0, finalCanvas.width, finalCanvas.height);
    
    // Draw the original canvas content
    ctx.drawImage(sourceCanvas, 0, 0);
    // Convert to PNG and download
    const link = document.createElement('a');
    link.download = `${id}.png`;
    link.href = finalCanvas.toDataURL('image/png');
    link.click();
  } catch (error) {
    console.error('Failed to export canvas:', error);
  }
}

/**
 * Utility Functions
 */

function fromDateStr(ds: string): number {
  return dayjs(ds).valueOf();
}

function numberStringSpaced(num: number): string {
  // Split the number into integer and decimal parts
  let [integerPart, decimalPart] = num.toString().split('.');

  // Format the integer part with spaces
  integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, " ");

  // If there is a decimal part, format it similarly with groups of 3 digits
  if (decimalPart) {
    // Split the decimal part into groups of 3
    decimalPart = decimalPart.replace(/(\d{3})(?=\d)/g, '$1 '); // Add space after every 3 digits
    return `${integerPart}.${decimalPart}`;
  }

  return integerPart;
}

function toDateStr(d: Date): string {
  const offset = d.getTimezoneOffset();
  d = new Date(d.getTime() - offset * 60 * 1000);
  return d.toISOString().slice(0, 10);
}

function extractDataFromUrl(): URLSearchParams {
  const urlParams = new URLSearchParams(window.location.search);
  if (urlParams.has("d")) {
    return urlParams;
  }
  const hashParams = new URLSearchParams(window.location.hash.slice(1)); // Remove '#' character
  return hashParams;
}

function round(n: number, decimal_point: number = DECIMAL_POINT): number {
  let pow = 10 ** decimal_point;
  return Math.round(n * pow) / pow;
}

function atobUrlSafe(s: string) {
  return atob(s.replace(/-/g, "+").replace(/_/g, "/"));
}

function btoaUrlSafe(s: string) {
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=/g, "");
}

function encodeNumberArrayString(input: number[]) {
  const buffer = new ArrayBuffer(input.length * 4);
  const view = new DataView(buffer);
  input.forEach((i, idx) => {
    view.setFloat32(idx * 4, i);
  });
  return btoaUrlSafe(
    new Uint8Array(buffer).reduce(
      (data, byte) => data + String.fromCharCode(byte),
      ""
    )
  );
}

function decodeNumberArrayString(s: string) {
  const bs = atobUrlSafe(s);
  var bytes = new Uint8Array(bs.length);
  for (let i = 0; i < bs.length; i++) {
    bytes[i] = bs.charCodeAt(i);
  }
  const view = new DataView(bytes.buffer);
  const result = [] as number[];
  for (let i = 0; i < bytes.length / 4; i++) {
    result.push(view.getFloat32(i * 4));
  }
  return result;
}

function isShadowDividerLine(dl: { id: string }): boolean {
  return dl.id != null && (dl.id == "divider-start" || dl.id == "divider-end");
}

function forceFloat(s: string) {
  if (!s) return s;
  return s.replace(/[^0-9.\-]/g, "");
}

function deepClone(src: DataValue[]): DataValue[] {
  return src.map((el) => {
    return { x: el.x, value: el.value, order: el.order, status: el.status };
  });
}

function updateInPlace(dest: DataValue[], src: DataValue[]) {
  dest.splice(0, dest.length, ...deepClone(src));
}

function debounce(func: TimerHandler, wait: number = 100) {
  let timer;
  return function (event) {
    if (timer) clearTimeout(timer);
    timer = setTimeout(func, wait, event);
  };
}

function hideElement(e: Element) {
  e?.classList.add("hidden");
}

function showElement(e: Element) {
  e?.classList.remove("hidden");
}

// sorts an array of DataValues in place
function sortDataValues(dv: DataValue[]) {
  return dv.sort((a, b) => fromDateStr(a.x) - fromDateStr(b.x));
}

const dataValueIsValid = (d: DataValue) =>
  d.x && (d.value || d.value == 0);

function removeAllNulls(dv: DataValue[]) {
  return dv.filter(dataValueIsValid);
}

function debug(v: any, msg?: string) {
  if (msg) {
    console.debug(msg);
  }
  console.debug(v);
  return v;
}