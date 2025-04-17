import dayjs from "dayjs";
import {
  dataLabelsStatusColor,
  dataStatusColor,
  DataValue,
  fromDateStr,
} from "./util";

const ECHARTS_DATE_FORMAT = "{d} {MMM}";

const toolbox = {
  feature: {
    saveAsImage: {
      type: "png",
      pixelRatio: 2,
    },
  },
};

const backgroundColor = "#ffffff";

const xAxis = {
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
  },
};

const yAxis = {
  splitLine: {
    show: false,
  },
  axisLabel: {
    fontSize: 11,
    color: "#000",
    hideOverlap: true,
    formatter: (value) => Number(value).toFixed(2),
  },
  interval: 500,
  nameLocation: "middle",
  nameRotate: 90,
  nameGap: 40,
  nameTextStyle: {
    color: "#000",
  },
};

export const chartBaseOptions = {
  toolbox,
  backgroundColor,
  xAxis,
  yAxis,
  title: {
    left: "center",
    triggerEvent: true,
  },
  tooltip: {
    show: true,
  },
};

const mapDataValueToChartDataPoint = (dv: DataValue) => ({
  value: [fromDateStr(dv.x), dv.value],
  itemStyle: {
    color: dataStatusColor(dv.status),
  },
  label: {
    show: true,
    color: dataLabelsStatusColor(dv.status),
    fontWeight: "bold",
    fontSize: 10,
    formatter: Number(dv.value).toFixed(2),
  },
  tooltip: {
    formatter: `${dayjs(dv.x).format("ddd, D MMM YYYY")}:<br/> ${Number(dv.value).toFixed(2)}`,
  },
});

export const mapDataValuesToChartSeries = (subD: DataValue[], i: number) => ({
  name: `${i}-data`,
  z: 98,
  type: "line",
  symbol: "circle",
  symbolSize: 7,
  lineStyle: {
    color: "#000",
  },
  data: subD.map(mapDataValueToChartDataPoint),
});
