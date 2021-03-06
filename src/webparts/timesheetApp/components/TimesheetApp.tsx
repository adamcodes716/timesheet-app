import * as React from 'react';
import styles from './TimesheetApp.module.scss';
import { ITimesheetAppProps } from './ITimesheetAppProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { v4 as uuidv4 } from 'uuid';

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

// Configure the columns for the DetailsList component
let _timesheetListColumns = [

  {
    key: 'TimesheetId',
    name: 'Timesheet ID',
    fieldName: 'TimesheetId',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'FullName',
    name: 'Name',
    fieldName: 'FullName',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'WeekEnding',
    name: 'Week Ending',
    fieldName: 'WeekEnding',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'TotalHours',
    name: 'Total Hours',
    fieldName: 'TotalHours',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'Status',
    name: 'Status',
    fieldName: 'Status',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  }
];
  const listName = "Timesheets";
  const textFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 300 } };
  const narrowTextFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 100 } };
  const mediumTextFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 150 } };
  const narrowDropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };

export default class TimesheetApp extends React.Component<ITimesheetAppProps, ITimesheetState> {

  private _selection: Selection;


  private _onItemsSelectionChanged = () => {
    console.log ("Setting new timesheet list",this._selection.getSelection()[0] );
    this.setState({
      TimesheetListItem: (this._selection.getSelection()[0] as ITimesheetListItem), showTable: "false"
    });
  }

  constructor(props: ITimesheetAppProps, state: ITimesheetState) {
    super(props);

    this.state = {
      status: "",
      showTable: "true",
      TimesheetListItems : [],
      TimesheetListItem : {
          TimesheetId: "WT-" + uuidv4().toUpperCase().slice(-6),
          Id: "5",
          FullName: this.props.displayName,
          WeekEnding: null,
          TotalHours: "0",
          Manager: this.props.managerName,
          Status: ""
          }
    };

  // bind an event handler
  this._selection = new Selection({
    onSelectionChanged: this._onItemsSelectionChanged,
  });
}

  private _getListItems(): Promise<ITimesheetListItem[]> {
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('" + listName + "')/items";
    console.log ("url", url);
    return this.props.context.spHttpClient.get(url,SPHttpClient.configurations.v1)
    .then(response => {
    return response.json();
    })
    .then(json => {
    return json.value;
    }) as Promise<ITimesheetListItem[]>;
    }

    public bindDetailsList(message: string) : void {

      this._getListItems().then(listItems => {
        this.setState({ TimesheetListItems: listItems, status: message, showTable: "true"});
      });
    }

    public componentDidMount(): void {
       this.bindDetailsList("All timesheets have been loaded Successfully");
    }

    @autobind
    public btnCreate_click(): void {
      this.setState({ showTable: "false", status: null});
    }

    @autobind
    public btnCancel_click(): void {
      this.setState({ showTable: "true"});
      this.bindDetailsList("Timesheets have been loaded");
    }

    @autobind
    public btnAdd_click(): void {

      const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('" + listName + "')/items";
      this.state.TimesheetListItem.Status = "Submitted";
      this.state.TimesheetListItem.Id = Math.random().toString().slice(-6);

         const spHttpClientOptions: ISPHttpClientOptions = {
        "body": JSON.stringify(this.state.TimesheetListItem)
      };

      this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {

        if (response.status === 201) {
          this.setState ({showTable : "true"});
          this.bindDetailsList("Your timesheet has been submitted");



        } else {
          let errormessage: string = "An error has occured i.e.  " + response.status + " - " + response.statusText;
          this.setState({status: errormessage});
        }
      });
    }

    @autobind
  public btnUpdate_click(): void {

    let id: string = this.state.TimesheetListItem.Id;
    console.log ("number", id);
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('" + listName + "')/items(" + id + ")";

    const headers: any = {
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*",
    };

       const spHttpClientOptions: ISPHttpClientOptions = {
        "headers": headers,
        "body": JSON.stringify(this.state.TimesheetListItem)
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {

      if (response.status === 204) {
        this.setState ({showTable : "true"});
        this.bindDetailsList("Your timesheet has been saved");

      } else {
        let errormessage: string = "An error has occured i.e.  " + response.status + " - " + response.statusText;
        this.setState({status: errormessage});
      }
    });
  }


  @autobind
  public btnDelete_click(): void {
    let id: string = this.state.TimesheetListItem.Id;
    //const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('" + listName + "')/items(" + id + ")";
    const url: string = this.props.siteUrl + "/_api/web/lists/getbytitle('" + listName + "')/getItemById(" + id + ")";

    const headers: any = { "X-HTTP-Method": "DELETE", "IF-MATCH": "*" };
    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {
      if (response.status === 204) {
        this.setState ({showTable : "true"});
        this.bindDetailsList("Your timesheet has been deleted");

      } else {
        let errormessage: string = "An error has occured i.e.  " + response.status + " - " + response.statusText;
        this.setState({status: errormessage});
      }
    });
  }

  public render(): React.ReactElement<ITimesheetAppProps> {
    return (
      <div className={ styles.timesheetApp }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Weekly Timesheet Application</span>
           </div>
          </div>
        </div>

        <div id="timesheetContainger" className={ styles.container } style={{display: !(this.state.showTable === "true") ? 'block' : 'none' }}>
          <div className={ styles.row }>
          <div className={ styles.smallColumn }>
                <TextField
                  label="Timesheet ID"
                  required={ true }
                  value={ (this.state.TimesheetListItem.TimesheetId).toString()}
                  styles={mediumTextFieldStyles}
                  readOnly
                  onChanged={e => {this.state.TimesheetListItem.TimesheetId=e;}}
                />
            </div>
            <div className={ styles.smallColumn }>
                <TextField
                  label="Name"
                  required={ true }
                  value={ (this.state.TimesheetListItem.FullName)}
                  styles={mediumTextFieldStyles}
                  readOnly
                  onChanged={e => {this.state.TimesheetListItem.FullName=e;}}
                />
            </div>
            <div className={ styles.smallColumn }>
            <TextField
                  label="Manager"
                  required={ true }
                  value={ (this.state.TimesheetListItem.Manager)}
                  styles={mediumTextFieldStyles}
                  readOnly
                  onChanged={e => {this.state.TimesheetListItem.Manager=e;}}
                />
            </div>

           </div>
           <div className={ styles.row }>
              <div className={ styles.smallColumn }>
                  <TextField
                      label="Week Ending"
                      required={ true }
                      value={ (this.state.TimesheetListItem.WeekEnding)}
                      styles={narrowTextFieldStyles}
                      onChanged={e => {this.state.TimesheetListItem.WeekEnding=e;}}
                  />
              </div>
              <div className={ styles.smallColumn }>
                   <TextField
                      label="Hours"
                      required={ true }
                      value={ (this.state.TimesheetListItem.TotalHours)}
                      styles={narrowTextFieldStyles}
                      onChanged={e => {this.state.TimesheetListItem.TotalHours=e;}}
                    />
              </div>
              <div className={ styles.smallColumn }>

              </div>
            </div>
          </div>

          <div className={ styles.container } style={{display: this.state.showTable === "false" ? 'block' : 'none' }}>
                <p className={styles.title}>
                 <PrimaryButton
                    text='Submit'
                    title='Submit'
                    onClick={this.btnAdd_click}
                  />
                 &nbsp;
                  <PrimaryButton
                    text='Delete'
                    onClick={this.btnDelete_click}
                  />
                  &nbsp;
                  <PrimaryButton
                    text='Cancel'
                    onClick={this.btnCancel_click}
                  />
                </p>
              </div>




          <div className={ styles.container } style={{display: this.state.showTable === "true" ? 'block' : 'none' }}>
                <p className={styles.title}>
                <PrimaryButton
                    text='Create'
                    title='Create'
                    onClick={this.btnCreate_click}
                  />
                 </p>
              </div>

              <div  className={ styles.container } >
                 <div style={{display: this.state.showTable === "true" ? 'block' : 'none' }}>
                  <DetailsList
                        items={ this.state.TimesheetListItems}
                        columns={ _timesheetListColumns }
                        setKey='Id'
                        checkboxVisibility={ CheckboxVisibility.onHover}
                        selectionMode={ SelectionMode.single}
                        layoutMode={ DetailsListLayoutMode.fixedColumns }
                        compact={ true }
                        selection={this._selection}
                    />
                   </div>
                 <div id="divStatus">
                  {this.state.status}
                </div>
                </div>
      </div>
    );
  }
}
