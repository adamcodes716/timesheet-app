import { ITimesheetListItem } from './ITimesheetListItem';

export interface ITimesheetState {
  status: string;
  TimesheetListItems : ITimesheetListItem[];
  TimesheetListItem : ITimesheetListItem;
}
