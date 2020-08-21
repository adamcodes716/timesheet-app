import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TimesheetAppWebPartStrings';
import TimesheetApp from './components/TimesheetApp';
import { ITimesheetAppProps } from './components/ITimesheetAppProps';

export interface ITimesheetAppWebPartProps {
  description: string;
}

export default class TimesheetAppWebPart extends BaseClientSideWebPart <ITimesheetAppWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITimesheetAppProps> = React.createElement(
      TimesheetApp,
      {
        description: this.properties.description,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        displayName: this.context.pageContext.user.displayName,
        managerName: this.context.pageContext.user.displayName

      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
