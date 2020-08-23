import { ITimesheetListItem } from './ITimesheetListItem';

export interface ITimesheetState {
  status: string;
  showTable: string;
  TimesheetListItems : ITimesheetListItem[];
  TimesheetListItem : ITimesheetListItem;
}
