import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneLabel,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ResourceDocumentSubPagesWebPartStrings';
import ResourceDocumentSubPages from './components/ResourceDocumentSubPages';
import { IResourceDocumentSubPagesProps } from './components/IResourceDocumentSubPagesProps';

export interface IResourceDocumentSubPagesWebPartProps {
  description: string;
  documentLibraryName: string;
}

export default class ResourceDocumentSubPagesWebPart extends BaseClientSideWebPart<IResourceDocumentSubPagesWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IResourceDocumentSubPagesProps> = React.createElement(
      ResourceDocumentSubPages,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        userDisplayName: this.context.pageContext.user.displayName,
        // This authenticated context is used by the service to call SharePoint REST.
        context: this.context,
        // Property-pane changes flow to React and trigger a content refresh.
        documentLibraryName: this.properties.documentLibraryName || 'LearningandDevelopmentVideo'
      }
    );

    ReactDom.render(element, this.domElement);
    this._removeContainerMaxWidths();
  }

  /** Removes the width constraints added by modern pages and both workbenches. */
  private _removeContainerMaxWidths(): void {
    const containerSelector: string = [
      '.CanvasZone',
      '.CanvasZoneSection-container',
      '.CanvasSection',
      '.CanvasSection-col',
      '.ControlZone',
      '.ControlZone-container',
      '.SPCanvas-canvas',
      '.SPCanvasContent',
      '#workbenchPageContent',
      '.workbenchPageContent',
      '[class*="CanvasZone"]',
      '[class*="CanvasSection"]',
      '[class*="ControlZone"]',
      '[class*="SPCanvas"]',
      '[class*="workbench"]',
      '[data-automation-id="CanvasZone"]'
    ].join(', ');

    let container: HTMLElement | null = this.domElement;
    while (container) {
      if (container.matches(containerSelector)) {
        container.style.setProperty('width', '100%', 'important');
        container.style.setProperty('max-width', 'none', 'important');
        container.style.setProperty('margin-left', '0', 'important');
        container.style.setProperty('margin-right', '0', 'important');
        container.style.setProperty('padding-left', '0', 'important');
        container.style.setProperty('padding-right', '0', 'important');
      }

      container = container.parentElement;
    }
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
              groupName: 'Training library settings',
              groupFields: [
                PropertyPaneLabel('libraryHelp', {
                  text: 'Enter the SharePoint document library display name or internal/root folder name. Files inside it and all nested folders are loaded automatically, then shown as videos or documents.'
                }),
                PropertyPaneTextField('documentLibraryName', {
                  label: 'Document library name',
                  value: 'Learning and Development Video',
                  description: 'Example: Learning and Development Video Library or LearningandDevelopmentVideoLibrary. Leave blank to hide dynamic training content.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
