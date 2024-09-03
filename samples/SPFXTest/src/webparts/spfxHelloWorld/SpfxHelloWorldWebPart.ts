/* eslint-disable no-void */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SpfxHelloWorldWebPartStrings';
import SpfxHelloWorld from './components/SpfxHelloWorld';
import { ISpfxHelloWorldProps } from './components/ISpfxHelloWorldProps';
import { spfi, SPFI, SPFx as spSPFx } from '@pnp/sp';
import { graphfi, GraphFI, SPFx as graphSPFx } from '@pnp/graph';

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/graph/users";

export interface ISpfxHelloWorldWebPartProps {
  description: string;
}

export default class SpfxHelloWorldWebPart extends BaseClientSideWebPart<ISpfxHelloWorldWebPartProps> {
  private LOG_SOURCE = "SpfxHelloWorldWebPart";
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _sp: SPFI;
  private _graph: GraphFI;
  private _results: string[] = [];

  public async onInit(): Promise<void> {
    this._sp = spfi().using(spSPFx(this.context));
    this._graph = graphfi().using(graphSPFx(this.context));
    const message = await this._getEnvironmentMessage();
    this._environmentMessage = message;

    this._results = [];
    // Call SPGet
    const testSPGetResult = await this._testSPGet();
    this._results.push(`_testSPGet: ${testSPGetResult}`);

    // Call SPPost
    const testSPPostResult = await this._testSPPost();
    this._results.push(`_testSPPost: ${testSPPostResult}`);

    // Call Graph
    const testGraphResult = await this._testGraph();
    this._results.push(`_testGraph: ${testGraphResult}`);
  }

  private _testSPGet = async (): Promise<boolean> =>  {
    let retVal = false;
    try{
      const web = await this._sp.web();
      retVal = true;
      console.log(this.LOG_SOURCE, "(_testSPGet)", web);
    }catch(err){
      console.error(this.LOG_SOURCE, "(_testSPGet)", err);
    }
    return retVal;
  }

  private _testSPPost = async (): Promise<boolean> =>  {
    let retVal = false;
    try{
      const web = await this._sp.web.lists.getByTitle("Test").items.add({Title: `Testing ${this.LOG_SOURCE} - ${this.context.pageContext.user.loginName}`});
      retVal = true;
      console.log(this.LOG_SOURCE, "(_testSPPost)", web);
    }catch(err){
      console.error(this.LOG_SOURCE, "(_testSPPost)", err);
    }
    return retVal;
  }

  private _testGraph = async (): Promise<boolean> =>  {
    let retVal = false;
    try{
      const me = await this._graph.me();
      retVal = true;
      console.log(this.LOG_SOURCE, "(_testGraph)", me);
    }catch(err){
      console.error(this.LOG_SOURCE, "(_testGraph)", err);
    }
    return retVal;
  }

  public render(): void {

    const element: React.ReactElement<ISpfxHelloWorldProps> = React.createElement(
      SpfxHelloWorld,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        loginName: this.context.pageContext.user.loginName,
        results: this._results
      }
    );

    ReactDom.render(element, this.domElement);
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
