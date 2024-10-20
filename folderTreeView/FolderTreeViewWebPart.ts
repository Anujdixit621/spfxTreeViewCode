import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FolderTreeViewWebPartStrings';
import FolderTreeView from './components/FolderTreeView';
import { IFolderTreeViewProps } from './components/IFolderTreeViewProps';

export interface IFolderTreeViewWebPartProps {
  libraryUrl: string;
  description: string;
  siteName: string;
}

export default class FolderTreeViewWebPart extends BaseClientSideWebPart<IFolderTreeViewWebPartProps> {

  // private _isDarkTheme: boolean = false;
  // private _environmentMessage: string = '';
  public render(): void {
    const element: React.ReactElement<IFolderTreeViewProps> = React.createElement(
      FolderTreeView,
      {
        libraryUrl: this.properties.libraryUrl , // Default library URL
        siteName : this.properties.siteName
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
                }),
                PropertyPaneTextField('siteName', {
                  label: "Site Name",
                  placeholder: "Enter the site name",
                  value:"",

                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
