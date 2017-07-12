import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'reactpnpasyncStrings';
import Reactpnpasync from './components/Reactpnpasync';
import { IReactpnpasyncProps } from './components/IReactpnpasyncProps';
import { IReactpnpasyncWebPartProps } from './IReactpnpasyncWebPartProps';

export default class ReactpnpasyncWebPart extends BaseClientSideWebPart<IReactpnpasyncWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactpnpasyncProps > = React.createElement(
      Reactpnpasync,
      {
        description: this.properties.description,
        pageContext: this.context.pageContext
      }
    );

    ReactDom.render(element, this.domElement);
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
