import dayjs from "dayjs";
import lz77 from "./lz77";
import { init as initEChart } from "echarts";
import type { EChartsType } from "echarts";
import { chartBaseOptions, mapDataValuesToChartSeries } from "./chart";
import { DataStatus, DataValue } from "./util";
import { initializeSigmaPlugin, mountSigmaProvider } from "./sigma-integration";

/**
 * Typescript type and interface definitions
 */

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

/**
 * Global constants and variables
 */
const MAX_LINK_LENGTH = 2000; // 2000 characters in url
const NPL_SCALING = 2.66;
const URL_SCALING = 3.268;
// PADDING_FROM_EXTREMES is the percentage from the data limits (max/min) that we use as chart limits.
const PADDING_FROM_EXTREMES = 0.1;
const DECIMAL_POINT = 2;
const LINE_STROKE_WIDTH = 2;
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

// This is the state of the app.
const state = {
  // Data from Sigma Computing
  tableData: [] as DataValue[],  // Contains the data from Sigma
  xLabel: "Date",
  yLabel: "Value",

  // xdata, movements are bound to the charts.
  // So, all update must be done in place (e.g. use updateInPlace) to maintain reactiveness.
  xdata: [] as DataValue[], // data passed to plot(...). Will not contain null values.
  movements: [] as DataValue[],
  dividerLines: [] as DividerType[],
  // this is the state of the locked limit lines, either calculated from tableData or overwritten by user.
  // Only written when the user press "Lock limits" in the popup.
  // Locked limits should be considered active if some values are non-zero.
  lockedLimits: structuredClone(INACTIVE_LOCKED_LIMITS),
  lockedLimitStatus: LockedLimitStatus.UNLOCKED,
};

// Initialize charts - these will be properly set on DOMContentLoaded
let xChart: EChartsType;
let mrChart: EChartsType;

/**
 * Detection Checks
 */
function checkRunOfEight(data: DataValue[], avg: number) {
  let isApplied = false;
  if (data.length < 8) {
    return;
  }
  let aboveOrBelow = 0; // We use an 8-bit integer. Bit is set to 0 if below avg, or 1 otherwise
  for (let i = 0; i < 7; i++) {
    if (data[i].value > avg) {
      aboveOrBelow |= 1 << i % 8;
    }
  }
  for (let i = 7; i < data.length; i++) {
    if (data[i].value > avg) {
      // set bit to 1
      aboveOrBelow |= 1 << i % 8;
    } else {
      // set bit to 0
      aboveOrBelow &= ~(1 << i % 8);
    }
    if (aboveOrBelow == 0 || aboveOrBelow == 255) {
      for (let j = i - 7; j <= i; j++) {
        data[j].status = DataStatus.RUN_OF_EIGHT_EXCEPTION;
        isApplied = true;
      }
    }
  }
  if (isApplied) {
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
  lowerQuartile: number,
  upperQuartile: number
) {
  let isApplied = false;
  if (data.length < 4) {
    return;
  }

  let belowQuartile = 0;
  let aboveQuartile = 0;
  // setup sliding window
  for (let i = 0; i < 3; i++) {
    if (data[i].value < lowerQuartile) {
      belowQuartile += 1;
    } else if (data[i].value > upperQuartile) {
      aboveQuartile += 1;
    }
  }

  for (let i = 3; i < data.length; i++) {
    // set value for the current window
    if (data[i].value < lowerQuartile) {
      belowQuartile += 1;
    } else if (data[i].value > upperQuartile) {
      aboveQuartile += 1;
    }

    if (belowQuartile >= 3 || aboveQuartile >= 3) {
      for (let j = i - 3; j <= i; j++) {
        data[j].status = DataStatus.FOUR_NEAR_LIMIT_EXCEPTION;
        isApplied = true;
      }
    }

    // reset value to prepare for next window
    if (data[i - 3].value < lowerQuartile) {
      belowQuartile -= 1;
    } else if (data[i - 3].value > upperQuartile) {
      aboveQuartile -= 1;
    }
  }
  if (isApplied) {
  }
}

function checkOutsideLimit(
  data: DataValue[],
  lowerLimit: number,
  upperLimit: number
) {
  let isApplied = false;
  data.forEach((dv) => {
    if (dv.value < lowerLimit) {
      dv.status = DataStatus.NPL_EXCEPTION;
      isApplied = true;
    } else if (dv.value > upperLimit) {
      dv.status = DataStatus.NPL_EXCEPTION;
      isApplied = true;
    }
  });
  if (isApplied) {
  }
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
  let [xPosition, _] = xChart.convertFromPixel("grid", [
    (xChart.getWidth() * (dividerCount + 1)) / 4,
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

  redraw();
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
  let x = document.querySelector("#xplot");
  let mr = document.querySelector("#mrplot");
  if (state.xdata.length > 31) {
    div.classList.remove("lg:flex-nowrap");
    x.classList.add("w-full");
    mr.classList.add("w-full");
  } else {
    div.classList.add("lg:flex-nowrap");
    x.classList.remove("w-full");
    mr.classList.remove("w-full");
  }
}

function redraw(immediately: boolean = true): _Stats {
  console.log("redraw - Starting with data:", { 
    xdataLength: state.xdata.length,
    movementsLength: state.movements.length,
    dividerLinesLength: state.dividerLines.length
  });
  
  let stats = wrangleData();
  console.log("redraw - wrangleData completed with stats:", {
    xchartMin: stats.xchartMin,
    xchartMax: stats.xchartMax,
    mrchartMax: stats.mrchartMax,
    lineValuesCount: stats.lineValues.length,
    xdataPerRangeCount: stats.xdataPerRange.length,
    movementsPerRangeCount: stats.movementsPerRange.length
  });
  
  redrawDividerButtons();
  reflowCharts();
  
  console.log("redraw - Calling doEChartsThings...");
  doEChartsThings(stats);
  console.log("redraw - doEChartsThings completed");
  
  return stats;
}

function removeDividerLine() {
  // remove the last added annotation
  let id = `divider-${state.dividerLines.length - 2}`;
  state.dividerLines = state.dividerLines.filter((d) => d.id != id);
  redraw();
}

// HELPER FUNCTIONS

// Does 2 things:
// (1) Filters data based on divider lines and calculates SPC statistics per range
// By default,`dividerLines` contains two 'invisible' dividerlines at both ends
// with no Line associated with it (so it doesn't get rendered in the chart)
// (2) Checks against the 3 XMR rules for each range and color data points accordingly
function wrangleData(): _Stats {
  console.log("wrangleData - Starting with xdata:", { 
    xdataLength: state.xdata.length,
    firstFewItems: state.xdata.slice(0, 3)
  });

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
  
  // Handle empty data gracefully
  if (tableData.length === 0) {
    console.warn("wrangleData - No valid data found in tableData");
    return {
      xchartMin: 0,
      xchartMax: 100,
      mrchartMax: 100,
      lineValues: [],
      xdataPerRange: [],
      movementsPerRange: []
    };
  }
  
  // Log sample data for debugging
  console.log("wrangleData - Sample of filtered data:", tableData.slice(0, 3));
  
  console.log("wrangleData - Filtered tableData:", { 
    tableDataLength: tableData.length,
    firstFewItems: tableData.slice(0, 3)
  });
  
  tableData.sort((a, b) => fromDateStr(a.x) - fromDateStr(b.x));
  updateInPlace(state.xdata, tableData);
  
  // Since a user might paste in data that falls beyond either limits of the previous x-axis range
  // we need to update our "shadow" divider lines so that the filteredXdata will always get all data
  let { min: xdataXmin, max: xdataXmax } = findExtremesX(state.xdata);
  
  if (isFinite(xdataXmin) && isFinite(xdataXmax)) {
    dividerLines[0].x = xdataXmin;
    dividerLines[dividerLines.length - 1].x = xdataXmax;
  } else {
    console.warn("wrangleData - Invalid date range detected:", { xdataXmin, xdataXmax });
    // Use default values in case we have invalid dates
    dividerLines[0].x = 0;
    dividerLines[dividerLines.length - 1].x = Infinity;
  }

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
  
  // If we have no data, return early with empty stats
  if (state.xdata.length === 0) {
    console.warn("wrangleData - No data in xdata, returning empty stats");
    stats.xchartMin = 0;
    stats.xchartMax = 100;
    stats.mrchartMax = 100;
    return stats;
  }

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
    let lv = {
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

  // We have calculated the max of the limit lines before, now we compare with the value of each data point
  // since some of them might be outside the limit and we want it to be in view
  state.xdata.forEach((dv) => {
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
  redraw();
});

screen.orientation.addEventListener("change", (e) => {
  redraw();
});

// No longer using Handsontable with Sigma Computing integration

function extractDataFromUrl(): URLSearchParams {
  const urlParams = new URLSearchParams(window.location.search);
  if (urlParams.has("d")) {
    return urlParams;
  }
  const hashParams = new URLSearchParams(window.location.hash.slice(1)); // Remove '#' character
  return hashParams;
}

function setDummyData() {
  // Set dummy data
  const values = [
    5045, 4350, 4350, 3975, 4290, 4430, 4485, 4285, 3980, 3925, 3645, 3760,
    3300, 3685, 3463, 5200,
  ];
  const sampleData: DataValue[] = values.map(function (el, i) {
    // const parsedDate = d3.timeParse("%Y-%m-%d")(`2020-01-${i + 1}`);
    // const x: Date = parsedDate || new Date();
    return {
      order: i,
      x: `2020-01-${(i < 9 ? "0" : "") + (i + 1)}`,
      value: el,
      status: DataStatus.NORMAL,
    };
  });

  state.tableData = sampleData;
}

function lockLimits() {
  const lockedLimits = calculateLockedLimits(); // calculate locked limits from the data
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

  redraw();
}

function initialiseHandsOnTable() {
  const lockedLimitDataTable = document.querySelector(
    "#lock-limit-dataTable"
  ) as HTMLDivElement;
  const table = document.querySelector("#dataTable") as HTMLDivElement;

  lockedLimitHot = new Handsontable(lockedLimitDataTable, {
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
  });
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
        redraw();
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
      redraw();
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
  console.log("DOMContentLoaded - Initializing application");
  
  const pageParams = extractDataFromUrl();

  // Initialize with empty state - data will come from Sigma
  state.xdata = [];
  state.movements = [];
  state.dividerLines = [
    { id: "divider-start", x: 0 },
    { id: "divider-end", x: Infinity }
  ];
  
  // Initialize ECharts instances
  console.log("DOMContentLoaded - Initializing chart elements");
  const xplotElement = document.getElementById("xplot");
  const mrplotElement = document.getElementById("mrplot");
  
  if (!xplotElement || !mrplotElement) {
    console.error("DOMContentLoaded - Chart elements not found in DOM:", {
      xplotElement: !!xplotElement,
      mrplotElement: !!mrplotElement
    });
    return;
  }
  
  xChart = initEChart(xplotElement);
  mrChart = initEChart(mrplotElement);
  console.log("DOMContentLoaded - Chart instances created:", {
    xChartInitialized: !!xChart,
    mrChartInitialized: !!mrChart
  });
  
  // Set up chart event handlers
  setupChartEventHandlers();
  
  // Initialize Sigma plugin
  console.log("DOMContentLoaded - Initializing Sigma plugin");
  initializeSigmaPlugin();
  
  // Initialize charts
  console.log("DOMContentLoaded - Rendering initial charts");
  renderCharts();
  
  // Set up divider buttons
  const addDividerButton = document.querySelector(
    "#add-divider"
  ) as HTMLButtonElement;
  const removeDividerButton = document.querySelector(
    "#remove-divider"
  ) as HTMLButtonElement;

  addDividerButton.addEventListener("click", addDividerLine);
  removeDividerButton.addEventListener("click", removeDividerLine);

  // Lock Limits
  const lockLimitButton = document.querySelector(
    "#lock-limit-btn"
  ) as HTMLButtonElement;
  const lockLimitWaringLabel = document.querySelector(
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

  // If the initial state has locked limits, we should show the buttons and warnings
  if (isLockedLimitsActive()) {
    document
      .querySelectorAll(".lock-limit-remove")
      .forEach((d) => d.classList.remove("hidden"));

    lockLimitWaringLabel.classList.remove("hidden");
  }

  lockLimitButton.addEventListener("click", (e) => {
    setLockedLimitInputs(!isLockedLimitsActive());
    lockLimitDialog.showModal();
  });

  lockLimitDialogCloseButton.addEventListener("click", (e) => {
    lockLimitDialog.close();
  });

  document.querySelectorAll(".lock-limit-remove").forEach((d) =>
    d.addEventListener("click", () => {
      d.classList.add("hidden"); // hide buttons
      lockLimitWaringLabel.classList.add("hidden"); // hide label
      state.lockedLimitStatus &= ~LockedLimitStatus.LOCKED; // set to unlocked
      lockLimitDialog.close();
      redraw();
    })
  );

  lockLimitDialogAddButton.addEventListener("click", () => {
    lockLimits();
    lockLimitDialog.close();
    // show lock-limit-remove button
    document
      .querySelectorAll(".lock-limit-remove")
      .forEach((d) => d.classList.remove("hidden"));
    lockLimitWaringLabel.classList.remove("hidden");
  });

  // Initialize Sigma data provider
  mountSigmaProvider((data, xLabel, yLabel) => {
    if (typeof window.addDebugLog === 'function') {
      window.addDebugLog(`Sigma data callback received: ${data.length} items, xLabel=${xLabel}, yLabel=${yLabel}`);
      
      if (data.length > 0) {
        window.addDebugLog(`First data item: ${JSON.stringify(data[0])}`);
      }
    }
    
    // Update application state with Sigma data
    state.xLabel = xLabel;
    state.yLabel = yLabel;
    state.tableData = data;
    
    if (typeof window.addDebugLog === 'function') {
      window.addDebugLog(`State updated with data, tableData length: ${state.tableData.length}`);
    }
    
    // Reset divider lines when data changes
    state.dividerLines = [
      { id: "divider-start", x: 0 },
      { id: "divider-end", x: Infinity }
    ];
    
    // Update our data
    updateInPlace(state.xdata, data);
    
    const movements = getMovements(data);
    updateInPlace(state.movements, movements);
    
    if (typeof window.addDebugLog === 'function') {
      window.addDebugLog(`Data arrays updated - xdata: ${state.xdata.length}, movements: ${state.movements.length}`);
      window.addDebugLog("Calling redraw with data...");
    }
    
    // Redraw charts with new data
    redraw();
    
    if (typeof window.addDebugLog === 'function') {
      window.addDebugLog("Redraw completed");
    }
  });

  // Share link button
  const shareLinkButton = document.querySelector(
    "#share-link"
  ) as HTMLButtonElement;

  shareLinkButton.addEventListener("click", () => {
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
    dataCopiedMessageLabel.classList.remove("hidden");
    setTimeout(() => {
      dataCopiedMessageLabel.classList.add("hidden");
    }, 2000);
  });
});

function calculateLimits(xdata: DataValue[]): Partial<LineValueType> {
  const movements = getMovements(xdata);
  // since avgX and avgMovement is used for further calculation, we only round it after calculating unpl, lnpl, url
  const avgX = xdata.reduce((a, b) => a + b.value, 0) / xdata.length;
  // filteredMovements might be empty
  const avgMovement =
    movements.reduce((a, b) => a + b.value, 0) / Math.max(movements.length, 1);
  const UNPL = avgX + NPL_SCALING * avgMovement;
  const LNPL = avgX - NPL_SCALING * avgMovement;
  const URL = URL_SCALING * avgMovement;
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

// This function is a 'no-divider' version of the calculation of the limits (specifically for the locked limits)
function calculateLockedLimits() {
  let xdata = state.tableData.filter(
    (dv) => dv.x && (dv.value || dv.value == 0)
  );
  return calculateLimits(xdata);
}

/**
 * Set the value, placeholder and style of lockedlimit inputs
 * @param updateInputValue whether to update the value of the inputs
 */
function setLockedLimitInputs(updateInputValue: boolean) {
  let lv = calculateLockedLimits(); // calculate locked limits from Sigma data
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

function forceFloat(s: string) {
  if (!s) return s;
  return s.replace(/[^0-9.\-]/g, "");
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

function toDateStr(d: Date): string {
  const offset = d.getTimezoneOffset();
  d = new Date(d.getTime() - offset * 60 * 1000);
  return d.toISOString().slice(0, 10);
}

function fromDateStr(ds: string): number {
  try {
    // Try to parse as-is with dayjs
    const parsed = dayjs(ds);
    if (parsed.isValid()) {
      return parsed.valueOf();
    }
    
    // Try JavaScript Date parsing
    const jsDate = new Date(ds);
    if (!isNaN(jsDate.getTime())) {
      return jsDate.getTime();
    }
    
    // Log error if we can't parse the date
    console.error(`fromDateStr - Failed to parse date: ${ds}`);
    // Return a default value
    return 0;
  } catch (err) {
    console.error(`fromDateStr - Error parsing date ${ds}:`, err);
    return 0;
  }
}

function updateInPlace(dest: DataValue[], src: DataValue[]) {
  console.log(`updateInPlace - Updating array from ${dest.length} to ${src.length} items`);
  dest.splice(0, dest.length, ...deepClone(src));
  console.log(`updateInPlace - After update: ${dest.length} items`);
}

function deepClone(src: DataValue[]): DataValue[] {
  return src.map((el) => {
    return { x: el.x, value: el.value, order: el.order, status: el.status };
  });
}

function isLockedLimitsActive(): boolean {
  return (state.lockedLimitStatus & LockedLimitStatus.LOCKED) == 1;
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

function btoaUrlSafe(s: string) {
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=/g, "");
}

function atobUrlSafe(s: string) {
  return atob(s.replace(/-/g, "+").replace(/_/g, "/"));
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

/**
 * Sharelink v0:
 * - v: 0
 * - d: lz77.compress(xLabel,yLabel,date-cols...).base64(float32array of value-cols).
 * - l: float32array of [avgX, avgMovement, LNPL, UNPL, URL, lockedLimitStatus]
 * - s: float32array of dividers x values (unix timestamp in milliseconds)
 * @returns
 */
function generateShareLink(s: {
  xLabel: string;
  yLabel: string;
  xdata: {
    x: string; // date in yyyy-mm-dd
    value: number;
  }[];
  dividerLines?: {
    id: string;
    x: number; // date in unix milliseconds
  }[];
  lockedLimits?: {
    avgX: number;
    avgMovement?: number;
    UNPL?: number; // upper natural process limit
    LNPL?: number; // lower natural process limit
    URL?: number; // upper range limit
  };
  lockedLimitStatus?: number;
}): string {
  let currentUrlParams = extractDataFromUrl();
  const paramsObj = {
    v: currentUrlParams.get("v") ?? "0", // get version from url or default to 0
  };

  if (paramsObj["v"] == "1") {
    // in sharelink version 1, we assume users don't intend to change their src data so we just copy the user's url here
    // Otherwise, they should change the result returned from the remote url.
    paramsObj["d"] = currentUrlParams.get("d");
  } else {
    let validXdata = s.xdata.filter((dv) => dv.x || dv.value);
    // basically first 2 are labels, followed by date-column, followed by value-column.
    // date-column are compressed using lz77
    // value-column are encoded as bytearray and converted into base64 string
    const dateText =
      `${s.xLabel.replace(",", ";")},${s.yLabel.replace(",", ";")},` +
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

  if (s.dividerLines) {
    const dividers = encodeNumberArrayString(
      s.dividerLines.filter((dl) => !isShadowDividerLine(dl)).map((dl) => dl.x)
    );
    if (dividers.length > 0) {
      paramsObj["s"] = dividers;
    }
  }
  if (s.lockedLimits && (s.lockedLimitStatus & LockedLimitStatus.LOCKED) == 1) {
    // IMPORTANT: If you change the number array below in any way except appending to it, please bump up the version number ('v') as it will break all existing links
    // and update the decoding logic below
    paramsObj["l"] = encodeNumberArrayString([
      s.lockedLimits.avgX,
      s.lockedLimits.avgMovement,
      s.lockedLimits.LNPL,
      s.lockedLimits.UNPL,
      s.lockedLimits.URL,
      s.lockedLimitStatus,
    ]);
  }
  const pageParams = new URLSearchParams(paramsObj);
  const fullPath = `${window.location.origin}${
    window.location.pathname
  }#${pageParams.toString()}`;
  return fullPath;
}

async function decodeShareLink(
  version: string,
  d: string,
  dividers: string,
  limits: string
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
  return {
    xLabel: labels[0].replace(";", ","),
    yLabel: labels[1].replace(";", ","),
    data,
    dividerLines,
    lockedLimits,
    lockedLimitStatus,
  };
}

function renderCharts() {
  // setup initial data from Sigma if available
  if (state.tableData.length > 0) {
    updateInPlace(
      state.xdata,
      state.tableData.filter((dv) => dv.x && (dv.value || dv.value == 0))
    );
    updateInPlace(state.movements, getMovements(state.xdata));
  }

  // Initialize or reset divider lines
  state.dividerLines = state.dividerLines
    .filter((dl) => !isShadowDividerLine(dl)) // filter out border "shadow" dividers
    .concat([
      { id: "divider-start", x: 0 },
      {
        id: "divider-end",
        x: Infinity,
      },
    ]);

  redraw();
}

/**
 * Find the extremes on the x-axis
 * @param arr
 * @returns
 */
function findExtremesX(arr: DataValue[]) {
  console.log("findExtremesX - Processing array:", { 
    length: arr.length, 
    sample: arr.slice(0, 3)
  });

  if (!arr || arr.length === 0) {
    console.warn("findExtremesX - Empty array provided");
    return { min: 0, max: Infinity }; // Default values if array is empty
  }
  
  let min = Infinity;
  let max = -Infinity;
  let validDates = 0;
  
  arr.forEach((el) => {
    try {
      if (!el.x) {
        console.warn("findExtremesX - Missing date value in entry:", el);
        return;
      }
      
      let d = fromDateStr(el.x);
      
      if (isNaN(d) || !isFinite(d)) {
        console.warn("findExtremesX - Invalid date value:", { x: el.x, parsed: d });
        return;
      }
      
      min = Math.min(min, d);
      max = Math.max(max, d);
      validDates++;
    } catch (err) {
      console.error("findExtremesX - Error parsing date:", { x: el.x, error: err });
    }
  });
  
  // If no valid dates were found, use default values
  if (validDates === 0 || !isFinite(min) || !isFinite(max)) {
    console.warn("findExtremesX - No valid dates found");
    return { min: 0, max: Infinity };
  }
  
  console.log("findExtremesX - Found range:", { min, max });
  return { min, max };
}

function round(n: number): number {
  let pow = 10 ** DECIMAL_POINT;
  return Math.round(n * pow) / pow;
}

function isShadowDividerLine(dl: { id: string }): boolean {
  return dl.id != null && (dl.id == "divider-start" || dl.id == "divider-end");
}

// echarts port
function doEChartsThings(stats: _Stats) {
  console.log("doEChartsThings - Starting chart rendering");
  
  console.log("doEChartsThings - Initializing ECharts");
  initialiseECharts(true);
  
  console.log("doEChartsThings - Rendering data series with stats:", {
    xdataPerRangeLength: stats.xdataPerRange.length,
    movementsPerRangeLength: stats.movementsPerRange.length
  });
  renderStats(stats);
  
  console.log("doEChartsThings - Adjusting chart axis");
  adjustChartAxis(stats);
  
  console.log("doEChartsThings - Rendering limit lines");
  renderLimitLines(stats);
  
  console.log("doEChartsThings - Rendering divider lines");
  renderEChartDividerLines(stats);
  
  console.log("doEChartsThings - Chart rendering complete");
}

function initialiseECharts(shouldReplaceState: boolean = false) {
  xChart.setOption({ ...chartBaseOptions }, shouldReplaceState);
  mrChart.setOption({ ...chartBaseOptions }, shouldReplaceState);
  xChart.setOption({
    title: {
      text: state.yLabel,
    },
    xAxis: {
      name: state.xLabel,
    },
    yAxis: {
      name: state.yLabel,
    },
  });
  mrChart.setOption({
    title: {
      text: `MR: ${state.yLabel}`,
    },
    xAxis: {
      name: state.xLabel,
    },
    yAxis: {
      name: state.yLabel,
    },
  });
}

function renderStats(stats: _Stats) {
  console.log("renderStats - Rendering with data:", {
    xdataPerRangeLength: stats.xdataPerRange.length,
    movementsPerRangeLength: stats.movementsPerRange.length
  });
  
  if (!stats.xdataPerRange || stats.xdataPerRange.length === 0) {
    console.warn("renderStats - No X data to render");
    // Render empty charts
    xChart.setOption({
      series: []
    });
  } else {
    xChart.setOption({
      series: stats.xdataPerRange.map(mapDataValuesToChartSeries),
    });
  }
  
  if (!stats.movementsPerRange || stats.movementsPerRange.length === 0) {
    console.warn("renderStats - No movement data to render");
    // Render empty charts
    mrChart.setOption({
      series: []
    });
  } else {
    mrChart.setOption({
      series: stats.movementsPerRange.map(mapDataValuesToChartSeries),
    });
  }
}

function renderLimitLines(stats: _Stats) {
  // Create a bunch of series for each split
  let xSeries = [];
  let mrSeries = [];

  for (let i = 0; i < stats.lineValues.length; i++) {
    let lv = stats.lineValues[i];

    let strokeWidth = LINE_STROKE_WIDTH;
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
          show: false,
        },
        label: {
          formatter: showLabel && `${Number(statisticY).toFixed(2)}`,
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
        showLabel: true,
      }),
      createHorizontalLimitLineSeries({
        name: `${i}-URL`,
        lineStyle: {
          color: LIMIT_SHAPE_COLOR,
          width: strokeWidth,
          type: lineType,
        },
        statisticY: lv.URL ?? 0,
        showLabel: true,
      }),
    ]);

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
        showLabel: true,
      }),
      createHorizontalLimitLineSeries({
        name: `${i}-unpl`,
        lineStyle: {
          color: LIMIT_SHAPE_COLOR,
          width: strokeWidth,
          type: lineType,
        },
        statisticY: lv.UNPL ?? 0,
        showLabel: true,
      }),
      createHorizontalLimitLineSeries({
        name: `${i}-lnpl`,
        lineStyle: {
          color: LIMIT_SHAPE_COLOR,
          width: strokeWidth,
          type: lineType,
        },
        statisticY: lv.LNPL ?? 0,
        showLabel: true,
      }),
    ]);
  }

  xChart.setOption({
    series: xSeries,
  });
  mrChart.setOption({
    series: mrSeries,
  });
}

function getScalingFactor(range: number): number {
  // Dynamically determine interval based on data range
  if (range >= 1000) return 500;
  if (range >= 500) return 100;
  if (range >= 100) return 50;
  if (range >= 50) return 10;
  if (range >= 10) return 5;
  return 1;
}

function ceilToNearestInterval(val: number, interval: number) {
  return Math.ceil(val / interval) * interval;
}

function floorToNearestInterval(val: number, interval: number) {
  return Math.floor(val / interval) * interval;
}

function getXChartYAxisMinMax(stats: _Stats) {
  // Calculate data range to determine appropriate scaling
  const dataRange = stats.xchartMax - stats.xchartMin;
  const interval = getScalingFactor(dataRange);
  
  return [
    floorToNearestInterval(
      stats.xchartMin -
        dataRange * PADDING_FROM_EXTREMES,
      interval
    ),
    ceilToNearestInterval(
      stats.xchartMax +
        dataRange * PADDING_FROM_EXTREMES,
      interval
    ),
  ];
}

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

  // Calculate appropriate Y-axis bounds and interval
  const [xChartYMin, xChartYMax] = getXChartYAxisMinMax(stats);
  const dataRange = stats.xchartMax - stats.xchartMin;
  const interval = getScalingFactor(dataRange);
  
  // Calculate appropriate movement chart scaling
  const mrDataRange = stats.mrchartMax;
  const mrInterval = getScalingFactor(mrDataRange);
  const mrChartMax = ceilToNearestInterval((1 + PADDING_FROM_EXTREMES) * stats.mrchartMax, mrInterval);

  xChart.setOption({
    yAxis: {
      min: xChartYMin,
      max: xChartYMax,
      interval: interval, // Set dynamic interval
      axisLabel: {
        formatter: (value) => Number(value).toFixed(2)
      }
    },
    xAxis: { min: xMin, max: xMax },
  });
  mrChart.setOption({
    yAxis: {
      max: mrChartMax,
      interval: mrInterval, // Set dynamic interval
      axisLabel: {
        formatter: (value) => Number(value).toFixed(2)
      }
    },
    xAxis: { min: xMin, max: xMax },
  });
}

function renderEChartDividerLines(stats: _Stats) {
  state.dividerLines.forEach((dl) => renderDividerLine2(dl, stats));
}

function renderDividerLine2(dividerLine: DividerType, stats: _Stats) {
  if (isShadowDividerLine(dividerLine)) {
    // empty or undefined id means it is the shadow divider lines, so we don't render it
    return;
  }

  // Convert domain from data to pixel dimension
  const [xChartYMin, xChartYMax] = getXChartYAxisMinMax(stats);
  const p1 = xChart.convertToPixel("grid", [dividerLine.x, xChartYMin]);
  const p2 = xChart.convertToPixel("grid", [dividerLine.x, xChartYMax]);

  xChart.setOption({
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
        z: 100, // ensure the divider renders above all other lines
        style: {
          lineWidth: DIVIDER_LINE_WIDTH,
          lineDash: "solid",
          stroke: "purple",
        },
        draggable: "horizontal",
        ondragend: (dragEvent) => {
          for (let d of state.dividerLines) {
            if (d.id == dragEvent.target.id) {
              const translatedPt = xChart.convertFromPixel("grid", [
                dragEvent.offsetX,
                0,
              ]);
              d.x = translatedPt[0];
              break;
            }
          }
          redraw();
        },
      },
    ],
  });
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
    redraw();
  }
}

// Add event handlers after chart instances are initialized
function setupChartEventHandlers() {
  if (xChart) {
    xChart.on("click", (params) => {
      if (params.componentType === "title") {
        promptNewColumnName();
      }
    });
  }

  if (mrChart) {
    mrChart.on("click", (params) => {
      if (params.componentType === "title") {
        promptNewColumnName();
      }
    });
  }
}
