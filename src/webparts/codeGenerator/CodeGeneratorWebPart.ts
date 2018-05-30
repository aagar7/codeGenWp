import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'CodeGeneratorWebPartStrings';
import CodeGenerator from './components/CodeGenerator';
import { ICodeGeneratorProps } from './components/ICodeGeneratorProps';

export interface ICodeGeneratorWebPartProps {
  description: string;
}

export default class CodeGeneratorWebPart extends BaseClientSideWebPart<ICodeGeneratorWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICodeGeneratorProps > = React.createElement(
      CodeGenerator,
      {
        description: this.properties.description
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
