import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import SafetyIncident from './components/SafetyIncident';
import { ISafetyIncidentProps } from './components/ISafetyIncidentProps';
import { ISafetyIncidentWebPartProps } from './ISafetyIncidentWebPartProps';

export default class SafetyIncidentWebPart extends BaseClientSideWebPart<ISafetyIncidentWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISafetyIncidentProps> = React.createElement(SafetyIncident, {
      listName: this.properties.listName,
      siteName: this.properties.siteName
    });
    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: 'Sharepoint List Name',
                  value: 'DefaultListName'
                }),
                PropertyPaneTextField('siteName', {
                  label: 'Sharepoint Site Name',
                  value: 'Safety'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
