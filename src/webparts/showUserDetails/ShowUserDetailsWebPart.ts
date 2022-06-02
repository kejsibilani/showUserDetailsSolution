import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ShowUserDetailsWebPartStrings';
import ShowUserDetails from './components/ShowUserDetails';
import { IShowUserDetailsProps } from './components/IShowUserDetailsProps';
import * as pnp from 'sp-pnp-js';

export interface IShowUserDetailsWebPartProps {
  level: string;
  listName: string;
}

export default class ShowUserDetailsWebPart extends BaseClientSideWebPart<IShowUserDetailsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
    pnp.setup({

      spfxContext: this.context

    });



    // optional, we are setting up the sp-pnp-js logging for debugging    
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IShowUserDetailsProps> = React.createElement(
      ShowUserDetails,
      {
        context: this.context,
        level: this.properties.level,
        listName: this.properties.listName,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
          header: {
            description: "Property pane to specify the users level"
          },
          groups: [
            {
              groupName: "Show Users By Level",
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: "Insert List Name"
                }),
                PropertyPaneTextField('level', {
                  label: "Insert Level"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
