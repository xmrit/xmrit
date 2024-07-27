import Handsontable from "handsontable";
import lz77 from "./lz77";
import * as Highcharts from "highcharts";
import * as Exporting from "highcharts/modules/exporting";
import * as OfflineExporting from "highcharts/modules/offline-exporting";
import * as Annotation from "highcharts/modules/annotations";
import dayjs from "dayjs";

// initialize Highcharts
Exporting.default(Highcharts);
OfflineExporting.default(Highcharts);
Annotation.default(Highcharts);

/**
 * Typescript type and interface definitions
 */

type DataValue = {
	order: number;
	x: string; // Date in YYYY-MM-DD format
	value: number;
	status: DataStatus;
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

// LineType is a user selectable section in d3.
type LineType = Highcharts.Annotation;

// Divider Type is the backing type for the divider line.
interface DividerType {
	id: string;
	x: number;
	line: LineType | null;
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
const LIMIT_LINE_WIDTH = 2;
const DIVIDER_LINE_WIDTH = 4;
const TEXT_STONE_600 = "rgb(87 83 78)";
const INACTIVE_LOCKED_LIMITS = {
	avgX: 0,
	LNPL: Infinity,
	UNPL: -Infinity,
	avgMovement: 0,
	URL: -Infinity,
} as LineValueType;

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
};

// global variables for the charts
let xplot: Highcharts.Chart;
let mrplot: Highcharts.Chart;

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
			aboveOrBelow |= 1 << (i % 8);
		}
	}
	for (let i = 7; i < data.length; i++) {
		if (data[i].value > avg) {
			// set bit to 1
			aboveOrBelow |= 1 << (i % 8);
		} else {
			// set bit to 0
			aboveOrBelow &= ~(1 << (i % 8));
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
	upperQuartile: number,
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
	upperLimit: number,
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

function toLineSeriesObject(range): Highcharts.SeriesOptionsType {
	return {
		type: "line",
		color: TEXT_STONE_600,
		name: "",
		data: range,
		dataLabels: {
			enabled: true,
			style: {
				color: TEXT_STONE_600,
			},
		},
		marker: {
			symbol: "circle",
		},
	};
}

/**
 * Sets the xplot and mrplot global variable. Will only set once.
 */
function createPlots(xChartSelector: string, mrChartSelector: string) {
	function createPlot(selector: string, data: DataValue[][]): Highcharts.Chart {
		let options: Highcharts.Options = {
			chart: {
				type: "line",
				scrollablePlotArea: {
					minWidth: 600,
				},
				plotBackgroundImage: "../xmrit-bg.png",
			},
			time: {
				useUTC: false,
			},

			title: {
				text: selector === xChartSelector ? xplotTitle() : mrplotTitle(),
				useHTML: true,
			},

			plotOptions: {
				series: {
					stickyTracking: false,
				},
			},

			series: translateToSeriesData(data).map(toLineSeriesObject),

			yAxis: {
				title: {
					text: `<div class="axis-label-y">${state.yLabel}</div>`,
					useHTML: true,
				},
				gridLineWidth: 0, // disable horizontal grid lines
			},
			xAxis: {
				type: "datetime",
				title: {
					text: state.xLabel,
				},
				maxPadding: 0.2, // space on the right end of xAxis to give space for labels
				events: {
					setExtremes: function (e) {
						// https://api.highcharts.com/highcharts/xAxis.events.setExtremes
						if (e.trigger === "zoom" && !e.min && !e.max) {
							// When user click "reset zoom", we re-render the whole chart to make sure the limits are on display
							// here we use a setTimeout to very quickly override the initial zoom behaviour
							setTimeout(redraw, 1);
						}
					},
				},
			},
			legend: {
				enabled: false,
			},
			credits: {
				enabled: false,
			},
			exporting: {
				buttons: {
					contextButton: {
						enabled: false,
						menuItems: ["downloadPNG", "downloadJPEG", "downloadSVG"],
					},
				},
				fallbackToExportServer: false,
				allowHTML: true,
			},
		};

		return Highcharts.chart(selector, options);
	}

	// setup initial data
	updateInPlace(
		state.xdata,
		state.tableData.filter((dv) => dv.x && (dv.value || dv.value == 0)),
	);
	updateInPlace(state.movements, getMovements(state.xdata));

	// highcharts handle freeing up memory via destroy() when we write to the same container
	xplot = createPlot(xChartSelector, [state.xdata]);
	mrplot = createPlot(mrChartSelector, [state.movements]);

	state.dividerLines = state.dividerLines
		.filter((dl) => !isShadowDividerLine(dl)) // filter out border "shadow" dividers
		.concat([
			{ id: "divider-start", x: xplot.xAxis[0].min, line: null },
			// as mrplot might have 1 less point than xplot, when mrplot.xAxis[0].max is undefined, we substitute it with Infinity
			{
				id: "divider-end",
				x: Math.min(xplot.xAxis[0].max, mrplot.xAxis[0].max || Infinity),
				line: null,
			},
		]);
	// Since our annotations extend up to the last divider line, we need to make sure that that divider lines are still in view for both xplot and mrplot.
	// so we take the minimax value. Somehow, the xplot and mrplot xAxis[0].max are not the same.
	// I'm concerned that this might cause some "shadow" divider to be rendered,
	// but i think as long as renderDividerLine does not render divider lines with empty string ID, we should be fine.
	state.dividerLines.forEach(renderDividerLine);

	redraw();
}

// TODO: renderSeries should not draw a path crossing the limit lines.
function renderSeries(stats: _Stats, immediately: boolean = true) {
	xplot.update(
		{
			series: translateToSeriesData(stats.xdataPerRange).map(
				toLineSeriesObject,
			),
		},
		immediately,
		true,
	);
	mrplot.update(
		{
			series: translateToSeriesData(stats.movementsPerRange).map(
				toLineSeriesObject,
			),
		},
		immediately,
		true,
	);
}

function translateToSeriesData(d: DataValue[][]) {
	// d is the data from the table.
	// We need to convert it to the format that Highcharts expects.
	// We also need to convert the dates to milliseconds since epoch.
	return d.map((subD) =>
		subD.map((dv) => {
			return {
				x: fromDateStr(dv.x),
				y: dv.value,
				color: dataStatusColor(dv.status),
				dataLabels: {
					style: {
						color: dataLabelsStatusColor(dv.status),
						textOutline: 0,
					},
				},
			};
		}),
	);
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
	let xPosition = xplot.xAxis[0].toValue(
		(xplot.plotWidth * (dividerCount + 1)) / 4,
	);
	// trick: dividerLine might coincides with a data point, so we move it slightly to the right
	if (xPosition % 10) {
		xPosition += 1;
	}

	let dividerLine = {
		id: `divider-${dividerCount + 1}`,
		x: xPosition,
		line: null,
	};
	state.dividerLines.push(dividerLine);

	renderDividerLine(dividerLine);
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

function redraw(immediately: boolean = true): _Stats {
	let stats = wrangleData();
	// it is important that we render the series first before limit lines
	renderSeries(stats, immediately);
	renderLimitLines(stats, immediately);
	redrawDividerButtons();
	// let url = generateShareLink()
	// if (url.length <= MAX_LINK_LENGTH) {
	//   // store state in url if data is not too big
	//   // see corresponding simulation on "popstate"
	//   console.log("before pushstate", state.dividerLines)
	//   window.history.pushState({
	//     state: {
	//       tableData: state.tableData,
	//       xLabel: state.xLabel,
	//       yLabel: state.yLabel,

	//       xdata: state.xdata,
	//       movements: state.movements,
	//       dividerLines: state.dividerLines.map(line => {
	//         return {
	//           id: line.id,
	//           x: line.x,
	//         }
	//       })
	//     },
	//   }, "", url);
	// }
	return stats;
}

function renderDividerLine(dividerLine: DividerType) {
	if (isShadowDividerLine(dividerLine)) {
		// empty or undefined id means it is the shadow divider lines, so we don't render it
		return;
	}
	redraw(false);

	// Add the divider line
	xplot.removeAnnotation(dividerLine.id);
	dividerLine.line = xplot.addAnnotation(
		{
			id: dividerLine.id,
			animation: { defer: 0 },
			events: {
				afterUpdate: function (e) {
					for (let d of state.dividerLines) {
						if (d.line == e.target) {
							d.x = e.target.options.shapes[0].points[0].x;
							break;
						}
					}
					redraw();
				},
			},
			shapes: [
				{
					points: [
						// The y value is ~2^53, which is the largest number we can represent as a number (float64)
						// It should be large enough for most purposes.
						// This is definitely a hack :) to make this annotation looks like a plotLine
						//
						// An alternative is to implement the dividerLine using Highcharts.plotLines,
						// but that would require us to implement the drag-and-drop logic for the plotline
						// Another alternative is to keep extending the yAxis max and min based on the data,
						// but that would require us to carefully update it on every `afterUpdate` (avoiding infinite loop)
						{ x: dividerLine.x, y: -9e15, xAxis: 0, yAxis: 0 },
						{ x: dividerLine.x, y: 9e15, xAxis: 0, yAxis: 0 },
					],
					type: "path",
					stroke: "purple",
					strokeWidth: DIVIDER_LINE_WIDTH,
				},
			],
			draggable: "x",
			zIndex: 2,
		},
		false,
	);

	xplot.redraw(true);
	mrplot.redraw(true);
}

/**
 * renderLimitLines have a side effect of setting the chart yAxis extremes
 * renderLimitLines depends on the xAxis extremes being set correctly. Hence it must be called after renderSeries
 * @param stats
 * @param redraw whether to redraw. defaults to true
 */
function renderLimitLines(stats: _Stats, redraw: boolean = true): void {
	function labelFromShape(
		shape: Highcharts.AnnotationsShapesOptions,
	): Highcharts.AnnotationsLabelsOptions {
		return {
			point: shape.points[1],
			shape: "connector",
			align: "left",
		};
	}

	let shapesForXplot = [];
	let shapesForMrPlot = [];
	for (let i = 0; i < stats.lineValues.length; i++) {
		let lv = stats.lineValues[i];
		if (i == stats.lineValues.length - 1) {
			// try to extend the limit lines a bit after the last data point. We add minimum 1-day worth of x
			// this change is purely for rendering purposes, not logical change.
			// There is no reason for the math or for the number 5 aside from it looks good on a few data I tried on (daily, weekly, monthly)
			lv.xRight =
				lv.xRight +
				Math.max(
					86400 * 1000,
					(Math.min(xplot.xAxis[0].max, mrplot.xAxis[0].max) - lv.xRight) / 5,
				);
		}
		let strokeWidth = LIMIT_LINE_WIDTH;
		let meanShapeColor = "red";
		let limitShapeColor = "steelblue";
		let dashStyle = "ShortDash";
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
			dashStyle = "Solid";
		}
		if (options.useLowerQuartile) {
			let lowerQuartileShape = {
				points: [
					{ x: lv.xLeft, y: lv.lowerQuartile, xAxis: 0, yAxis: 0 },
					{ x: lv.xRight, y: lv.lowerQuartile, xAxis: 0, yAxis: 0 },
				],
				type: "path",
				dashStyle: "Dot",
				strokeWidth: 1,
			};
			shapesForXplot.push(lowerQuartileShape);
		}
		if (options.useUpperQuartile) {
			let upperQuartileShape = {
				points: [
					{ x: lv.xLeft, y: lv.upperQuartile, xAxis: 0, yAxis: 0 },
					{ x: lv.xRight, y: lv.upperQuartile, xAxis: 0, yAxis: 0 },
				],
				type: "path",
				dashStyle: "Dot",
				strokeWidth: 1,
			};
			shapesForXplot.push(upperQuartileShape);
		}
		let avgXShape = {
			points: [
				{ x: lv.xLeft, y: lv.avgX, xAxis: 0, yAxis: 0 },
				{ x: lv.xRight, y: lv.avgX, xAxis: 0, yAxis: 0 },
			],
			type: "path",
			dashStyle,
			stroke: meanShapeColor,
			strokeWidth,
		};
		let lnplShape = {
			points: [
				{ x: lv.xLeft, y: lv.LNPL, xAxis: 0, yAxis: 0 },
				{ x: lv.xRight, y: lv.LNPL, xAxis: 0, yAxis: 0 },
			],
			type: "path",
			dashStyle,
			stroke: limitShapeColor,
			strokeWidth,
		};
		let unplShape = {
			points: [
				{ x: lv.xLeft, y: lv.UNPL, xAxis: 0, yAxis: 0 },
				{ x: lv.xRight, y: lv.UNPL, xAxis: 0, yAxis: 0 },
			],
			type: "path",
			dashStyle,
			stroke: limitShapeColor,
			strokeWidth,
		};
		shapesForXplot.push(avgXShape, lnplShape, unplShape);

		let avgMovementShape = {
			points: [
				{ x: lv.xLeft, y: lv.avgMovement, xAxis: 0, yAxis: 0 },
				{ x: lv.xRight, y: lv.avgMovement, xAxis: 0, yAxis: 0 },
			],
			type: "path",
			dashStyle,
			stroke: meanShapeColor,
			strokeWidth,
		};
		let urlShape = {
			points: [
				{ x: lv.xLeft, y: lv.URL, xAxis: 0, yAxis: 0 },
				{ x: lv.xRight, y: lv.URL, xAxis: 0, yAxis: 0 },
			],
			type: "path",
			dashStyle,
			stroke: limitShapeColor,
			strokeWidth,
		};
		shapesForMrPlot.push(avgMovementShape, urlShape);
	}

	// Adjust the range of yaxis in both charts to keep all limit lines in view
	xplot.yAxis[0].setExtremes(
		stats.xchartMin -
			(stats.xchartMax - stats.xchartMin) * PADDING_FROM_EXTREMES,
		stats.xchartMax +
			(stats.xchartMax - stats.xchartMin) * PADDING_FROM_EXTREMES,
		false,
	);
	mrplot.yAxis[0].setExtremes(
		0,
		(1 + PADDING_FROM_EXTREMES) * stats.mrchartMax,
		false,
	);

	xplot.removeAnnotation("limit-lines");
	xplot.addAnnotation(
		{
			id: "limit-lines",
			animation: { defer: 0 },
			draggable: "",
			shapes: shapesForXplot,
			labels: shapesForXplot.map((shape) => labelFromShape(shape)).slice(-3),
			zIndex: 1,
		},
		false,
	);
	mrplot.removeAnnotation("limit-lines");
	mrplot.addAnnotation(
		{
			id: "limit-lines",
			animation: { defer: 0 },
			draggable: "",
			shapes: shapesForMrPlot,
			labels: shapesForMrPlot.map((shape) => labelFromShape(shape)).slice(-2),
			zIndex: 1,
		},
		false,
	);

	xplot.redraw(redraw);
	mrplot.redraw(redraw);
}

function removeDividerLine() {
	// remove the last added annotation
	let id = `divider-${state.dividerLines.length - 2}`;
	xplot.removeAnnotation(id);
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
	let dividerLines = state.dividerLines;
	// need to make sure dividerLines are sorted
	dividerLines.sort((a, b) => a.x - b.x);

	console.assert(
		dividerLines.length >= 2,
		"dividerLines should contain at least two divider lines",
	);

	// make sure state.xdata only contains valid data (i.e. have both x and value columns set)
	updateInPlace(
		state.xdata,
		state.tableData.filter((dv) => dv.x && (dv.value || dv.value == 0)),
	);
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
				opts.useUpperQuartile ? state.lockedLimits.upperQuartile : Infinity,
			);
			checkOutsideLimit(
				filteredXdata,
				state.lockedLimits.LNPL,
				state.lockedLimits.UNPL,
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
					"file-error",
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

let hot;
let lockedLimitHot;

/**
 * Initializes event listener for title / yaxis label changes.
 */
function registerYAxisTitleChangeListener() {
	const listener = (e: PointerEvent) => {
		let newColName = prompt("Insert a new column name", state.yLabel);
		if (newColName) {
			let colHeaders = hot.getColHeader();
			colHeaders[1] = newColName;
			state.yLabel = newColName;
			hot.updateSettings({
				colHeaders,
			});
			xplot.update({
				yAxis: { title: { text: yAxisTitle(newColName), useHTML: true } },
				title: { text: xplotTitle(), useHTML: true },
			});
			mrplot.update({
				yAxis: { title: { text: yAxisTitle(newColName), useHTML: true } },
				title: { text: mrplotTitle(), useHTML: true },
			});
			// re-register event listener
			registerYAxisTitleChangeListener();
		}
	};
	document
		.querySelectorAll(".axis-label-y")
		.forEach((el) => el.addEventListener("click", listener));
	document
		.querySelectorAll(".plot-title")
		.forEach((el) => el.addEventListener("click", listener));
}

function extractDataFromUrl(): URLSearchParams {
	const urlParams = new URLSearchParams(window.location.search);
	if (urlParams.has("d")) {
		return urlParams;
	}
	const hashParams = new URLSearchParams(window.location.hash.slice(1)); // Remove '#' character
	return hashParams;
}

// LOGIC ON PAGE LOAD
document.addEventListener("DOMContentLoaded", async function (_e) {
	const pageParams = extractDataFromUrl();
	if (pageParams.has("d")) {
		let version = pageParams.get("v") || "0";
		let separator = pageParams.get("s") || "";
		let ll = pageParams.get("l") || "";
		let {
			xLabel,
			yLabel,
			data,
			dividerLines,
			lockedLimits,
			lockedLimitStatus,
		} = await decodeShareLink(version, pageParams.get("d")!, separator, ll);
		state.xLabel = xLabel;
		state.yLabel = yLabel;
		state.tableData = data;
		state.dividerLines = dividerLines;
		state.lockedLimits = lockedLimits;
		state.lockedLimitStatus = lockedLimitStatus;
	} else {
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
	// deep clone
	state.lockedLimitBaseData = deepClone(state.tableData);
	renderCharts();
	document
		.querySelector("#add-divider")
		.addEventListener("click", addDividerLine);
	document
		.querySelector("#remove-divider")
		.addEventListener("click", removeDividerLine);
	// export chart logic
	document.querySelector("#export-xplot").addEventListener("click", () => {
		// To fix an edgecase where old limit lines are still around when user exports the chart eventhough it doesn't show up in the displayed chart.
		// This edgecase is easier to reproduce if you modify the data after switching the tab to background for a while.
		renderCharts();
		xplot.exportChartLocal({ filename: "xplot" });
	});
	document.querySelector("#export-mrplot").addEventListener("click", () => {
		// To fix an edgecase where old limit lines are still around when user exports the chart eventhough it doesn't show up in the displayed chart.
		// This edgecase is easier to reproduce if you modify the data after switching the tab to background for a while.
		renderCharts();
		mrplot.exportChartLocal({ filename: "mrplot" });
	});
	registerYAxisTitleChangeListener();

	// limit-lines
	if (isLockedLimitsActive()) {
		// if the initial state has locked limits, we should show the buttons and warnings
		document
			.querySelectorAll(".lock-limit-remove")
			.forEach((d) => d.classList.remove("hidden"));
		document.querySelector("#lock-limit-warning").classList.remove("hidden");
	}
	document.querySelector("#lock-limit-btn").addEventListener("click", (e) => {
		if (!isLockedLimitsActive()) {
			// let the locked limit data reflect the latest table data if locked limits are not currently active
			updateInPlace(state.lockedLimitBaseData, state.tableData);
			lockedLimitHot.updateSettings({
				data: state.lockedLimitBaseData,
				colHeaders: [state.xLabel, state.yLabel],
			});
		}
		setLockedLimitInputs(!isLockedLimitsActive());
		let dialog = document.querySelector(
			"#lock-limit-dialog",
		) as HTMLDialogElement;
		dialog.showModal();
	});
	document.querySelector("#lock-limit-close").addEventListener("click", (e) => {
		let dialog = document.querySelector(
			"#lock-limit-dialog",
		) as HTMLDialogElement;
		dialog.close();
	});
	document.querySelectorAll(".lock-limit-remove").forEach((d) =>
		d.addEventListener("click", (e) => {
			xplot.removeAnnotation("locked-limits");
			mrplot.removeAnnotation("locked-limits");
			d.classList.add("hidden"); // hide buttons
			document.querySelector("#lock-limit-warning").classList.add("hidden");

			state.lockedLimitStatus &= ~LockedLimitStatus.LOCKED; // set to unlocked
			redraw();
			let dialog = document.querySelector(
				"#lock-limit-dialog",
			) as HTMLDialogElement;
			dialog.close();
		}),
	);
	document.querySelector("#lock-limit-add").addEventListener("click", (e) => {
		let lv = calculateLockedLimits(); // calculate locked limits from the table
		let obj = structuredClone(INACTIVE_LOCKED_LIMITS);
		document
			.querySelectorAll(".lock-limit-input")
			.forEach((el: HTMLInputElement) => {
				obj[el.dataset.limit] =
					el.value !== "" ? Number(el.value) : lv[el.dataset.limit];
			});
		obj.lowerQuartile = round((obj.avgX + obj.LNPL) / 2);
		obj.upperQuartile = round((obj.avgX + obj.UNPL) / 2);

		// validate user input
		if (
			obj.avgX < obj.LNPL ||
			obj.avgX > obj.UNPL ||
			obj.avgMovement > obj.URL
		) {
			alert(
				"Please ensure that the following limits are satisfied:\n" +
					"1. Average X is between Lower Natural Process Limit (LNPL) and Upper Natural Process Limit (UNPL)\n" +
					"2. Average Movement is less than or equal to Upper Range Limit (URL)",
			);
			return;
		}
		if (lv.avgX != obj.avgX) {
			state.lockedLimitStatus |= LockedLimitStatus.AVGX_MODIFIED;
		}
		if (lv.LNPL != obj.LNPL) {
			state.lockedLimitStatus |= LockedLimitStatus.LNPL_MODIFIED;
		}
		if (lv.UNPL != obj.UNPL) {
			state.lockedLimitStatus |= LockedLimitStatus.UNPL_MODIFIED;
		}
		state.lockedLimits = obj; // set state
		state.lockedLimitStatus |= LockedLimitStatus.LOCKED; // set to locked
		redraw();

		let dialog = document.querySelector(
			"#lock-limit-dialog",
		) as HTMLDialogElement;
		dialog.close();
		// show lock-limit-remove button
		document
			.querySelectorAll(".lock-limit-remove")
			.forEach((d) => d.classList.remove("hidden"));
		document.querySelector("#lock-limit-warning").classList.remove("hidden");
	});
	lockedLimitHot = new Handsontable(
		document.querySelector("#lock-limit-dataTable"),
		{
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
				return autofillTable(
					selectionData,
					sourceRange,
					targetRange,
					direction,
				);
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
		},
	);

	// CSV input logic
	const csvFile = document.getElementById("csv-file") as HTMLInputElement;
	csvFile.addEventListener("change", (_event) => {
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

			renderCharts();
		});
		// Reads the csv file as string, after which it emits the loadend event
		reader.readAsText(input);
	});

	// Set event listener to download dummy data as csv file
	document.querySelector("#download-data").addEventListener("click", () => {
		// only download lines with at least one non-null value.
		let csvContent =
			`${state.xLabel},${state.yLabel}\r\n` +
			state.xdata
				.filter((dv) => dv.x || dv.value)
				.map((d) => `${d.x || ""},${round(d.value) || ""}`)
				.join("\r\n");
		if (isLockedLimitsActive()) {
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
	});

	document.querySelector("#refresh-charts").addEventListener("click", () => {
		renderCharts();
	});

	document.querySelector("#share-link").addEventListener("click", () => {
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
			},
		);

		// toggle message on share button click
		document.getElementById("data-copied-msg").classList.remove("hidden");
		setTimeout(() => {
			document.getElementById("data-copied-msg").classList.add("hidden");
		}, 2000);
	});

	let table = document.querySelector("#dataTable");
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
				this.getColHeader()[coords.col],
			);
			if (newColName) {
				let colHeaders = this.getColHeader();
				colHeaders[coords.col] = newColName;
				this.updateSettings({
					colHeaders,
				});
				if (coords.col == 0) {
					state.xLabel = newColName;
					xplot.update({
						xAxis: { title: { text: yAxisTitle(newColName), useHTML: true } },
					});
					mrplot.update({
						xAxis: { title: { text: yAxisTitle(newColName), useHTML: true } },
					});
				} else {
					state.yLabel = newColName;
					xplot.update({
						yAxis: { title: { text: yAxisTitle(newColName), useHTML: true } },
						title: { text: xplotTitle(), useHTML: true },
					});
					mrplot.update({
						yAxis: { title: { text: yAxisTitle(newColName), useHTML: true } },
						title: { text: mrplotTitle(), useHTML: true },
					});
				}
				// re-register event listener
				registerYAxisTitleChangeListener();
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
			changes.forEach(([row, prop, oldVal, newVal]) => {
				if (prop != "x") return;
				xOnly[row] = newVal;
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
});

function calculateLimits(xdata: DataValue[]): LineValueType {
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
	} as LineValueType;
}

// This function is a 'no-divider' version of the calculation of the limits (specifically for the locked limits)
function calculateLockedLimits() {
	let xdata = state.lockedLimitBaseData.filter(
		(dv) => dv.x && (dv.value || dv.value == 0),
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
			'.lock-limit-input[data-limit="avgX"]',
		) as HTMLElement
	).style["color"] = isAvgXModified(state.lockedLimitStatus)
		? "rgb(220 38 38)"
		: "black";
	(
		document.querySelector(
			'.lock-limit-input[data-limit="avgX"]',
		) as HTMLInputElement
	).placeholder = `${lv.avgX}`;
	(
		document.querySelector(
			'.lock-limit-input[data-limit="UNPL"]',
		) as HTMLElement
	).style["color"] = isUnplModified(state.lockedLimitStatus)
		? "rgb(220 38 38)"
		: "black";
	(
		document.querySelector(
			'.lock-limit-input[data-limit="UNPL"]',
		) as HTMLInputElement
	).placeholder = `${lv.UNPL}`;
	(
		document.querySelector(
			'.lock-limit-input[data-limit="LNPL"]',
		) as HTMLElement
	).style["color"] = isLnplModified(state.lockedLimitStatus)
		? "rgb(220 38 38)"
		: "black";
	(
		document.querySelector(
			'.lock-limit-input[data-limit="LNPL"]',
		) as HTMLInputElement
	).placeholder = `${lv.LNPL}`;
	(
		document.querySelector(
			'.lock-limit-input[data-limit="avgMovement"]',
		) as HTMLInputElement
	).placeholder = `${lv.avgMovement}`;
	(
		document.querySelector(
			'.lock-limit-input[data-limit="URL"]',
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
				(x) => new Date(new Date(x.valueOf()).setDate(x.getDate() + 1)),
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
				(x) => new Date(new Date(x.valueOf()).setDate(x.getDate() - 1)),
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
			(x) => new Date(x.valueOf() + dateArray.length * difference),
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
			(x) => new Date(x.valueOf() - dateArray.length * difference),
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

// window.addEventListener("popstate", function (e) {
//   console.log("popstate", e)
//   if (e.state) {
//     console.log("state", e.state)

//     updateInPlace(state.tableData, e.state.state.tableData)
//     state.xLabel = e.state.state.xLabel
//     state.yLabel = e.state.state.yLabel

//     hot.updateSettings({
//       colHeaders: [state.xLabel, state.yLabel]
//     })
//     updateInPlace(state.xdata, e.state.state.xdata)
//     updateInPlace(state.movements, e.state.state.movements)
//     // state.dividerLines.forEach(dl => !isShadowDividerLine(dl) && xplot.removeAnnotation(dl.id))
//     state.dividerLines = e.state.state.dividerLines

//     renderLimitLines(wrangleData())
//     renderSeries()
//     state.dividerLines.forEach(renderDividerLine)
//   }
// });

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
			state.lockedLimits.LNPL,
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
	return dayjs(ds).valueOf();
}

function getSign(n: number) {
	return n > 0 ? 1 : n < 0 ? -1 : 0;
}

function updateInPlace(dest: DataValue[], src: DataValue[]) {
	dest.splice(0, dest.length, ...deepClone(src));
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
			"",
		),
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
			s.dividerLines.filter((dl) => !isShadowDividerLine(dl)).map((dl) => dl.x),
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
	const fullPath = `${window.location.origin}${window.location.pathname}#${pageParams.toString()}`;
	return fullPath;
}

async function decodeShareLink(
	version: string,
	d: string,
	dividers: string,
	limits: string,
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
					line: null,
				};
			}),
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
	createPlots("xplot", "mrplot");
}

const xplotTitle = () =>
	`<div class="plot-title">X Plot<span>${state.yLabel.toLowerCase() !== "value" ? ": " + state.yLabel : ""}</span></div>`;
const mrplotTitle = () =>
	`<div class="plot-title">MR Plot<span>${state.yLabel.toLowerCase() !== "value" ? ": " + state.yLabel : ""}</span></div>`;
const yAxisTitle = (newColName) =>
	`<div class="axis-label-y">${newColName}</div>`;

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

function round(n: number): number {
	let pow = 10 ** DECIMAL_POINT;
	return Math.round(n * pow) / pow;
}

function isShadowDividerLine(dl: { id: string }): boolean {
	return dl.id && (dl.id == "divider-start" || dl.id == "divider-end");
}
