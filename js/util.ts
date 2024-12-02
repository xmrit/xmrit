import dayjs from "dayjs";

export type DataValue = {
  order: number;
  x: string; // Date in YYYY-MM-DD format
  value: number;
  status: DataStatus;
};

export enum DataStatus {
  NORMAL = 0,
  RUN_OF_EIGHT_EXCEPTION = 1, // 8 on one side
  FOUR_NEAR_LIMIT_EXCEPTION = 2, // 3 out of 4 on the extreme quarters
  NPL_EXCEPTION = 3, // out of the NPL limit lines.
}

export function fromDateStr(ds: string): number {
  return dayjs(ds).valueOf();
}

export function dataStatusColor(d: DataStatus): string {
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

export function dataLabelsStatusColor(d: DataStatus): string {
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
