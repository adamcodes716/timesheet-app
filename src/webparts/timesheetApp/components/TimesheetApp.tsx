import * as React from 'react';
import styles from './TimesheetApp.module.scss';
import { ITimesheetAppProps } from './ITimesheetAppProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { ITimesheetListItem } from './ITimesheetListItem';
import { ITimesheetState } from './ITimesheetState';

import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import {
  TextField,
  autobind,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode,
  Dropdown,
  IDropdown,
  IDropdownOption,
  ITextFieldStyles,
  IDropdownStyles,
  DetailsRowCheck,
  Selection
} from 'office-ui-fabric-react';

export default class TimesheetApp extends React.Component<ITimesheetAppProps, ITimesheetState> {

  constructor(props: ITimesheetAppProps, state: ITimesheetState) {
    super(props);

    this.state = {
      status: "New",
  TimesheetListItems : [],
  TimesheetListItem : {
      ID: "",
      FullName: "",
      WeekEnding: "01/01/2020",
      TotalHours: 0,
      }
    };

  }

  public render(): React.ReactElement<ITimesheetAppProps> {
    return (
      <div className={ styles.timesheetApp }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Timesheet Application</span>
              <p className={ styles.subTitle }>Submit your hours here</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
