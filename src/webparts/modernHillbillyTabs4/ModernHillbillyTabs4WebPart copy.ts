/* eslint-disable guard-for-in */
/* eslint-disable @typescript-eslint/no-for-in-array */
// import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import styles from './components/ModernHillbillyTabs4.module.scss';

import * as strings from 'ModernHillbillyTabs4WebPartStrings';
// import ModernHillbillyTabs4 from './components/ModernHillbillyTabs4';
// import { IModernHillbillyTabs4Props } from './components/IModernHillbillyTabs4Props';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import * as $ from 'jquery';

export interface IModernHillbillyTabs4WebPartProps {
  description: string;
  sectionClass: string;
  webpartClass: string;
  tabData: any[];
}

export default class ModernHillbillyTabs4WebPart extends BaseClientSideWebPart<IModernHillbillyTabs4WebPartProps> {

  // private _isDarkTheme: boolean = false;
  // private _environmentMessage: string = '';

  public render(): void {

    // Ensure jQuery is available globally for AddTabs.js
    // (window as any).jQuery = $;
    // require('./AddTabs.js');
    // require('./AddTabs.css');

    // const element: React.ReactElement<IModernHillbillyTabs4Props> = React.createElement(
    //   ModernHillbillyTabs4,
    //   {
    //     description: this.properties.description,
    //     isDarkTheme: this._isDarkTheme,
    //     environmentMessage: this._environmentMessage,
    //     hasTeamsContext: !!this.context.sdks.microsoftTeams,
    //     userDisplayName: this.context.pageContext.user.displayName,

       

    //   }
    // );


    
    if (this.displayMode === DisplayMode.Read)
    {
      let tabWebPartID = "";
      let zoneDIV = null;
      
      tabWebPartID = $(this.domElement).closest("div." + this.properties.webpartClass).attr("id") || "";       
      zoneDIV = $(this.domElement).closest("div." + this.properties.sectionClass);
      
      console.log(zoneDIV);
      
      const tabsDiv = tabWebPartID + "tabs";
      const contentsDiv = tabWebPartID + "Contents";
      
      this.domElement.innerHTML = "<div data-addui='tabs'><div role='tabs' id='"+tabsDiv+"'></div><div role='contents' id='"+contentsDiv+"'></div></div>";

      const thisTabData = this.properties.tabData;
      for(const x in thisTabData)
      {
        $("#"+tabsDiv).append("<div>"+thisTabData[x].TabLabel+"</div>");
        $("#"+contentsDiv).append($("#"+thisTabData[x].WebPartID));
      }

      
      // renderTabs();
      } else {
        this.domElement.innerHTML = `
        <div class="${ styles.modernHillbillyTabs }">
          <div class="${ styles.container }">
            <div class="${ styles.row }">
              <div class="${ styles.column }">
                <p class="${ styles.subTitle }">Place Web Parts into Tabs.</p>
                <p class="${ styles.description }">To use Modern Hillbilly Tabs: 
                  <ul>
                    <li>Place this web part in the same section of the page as the web parts you would like to put into tabs.</li> 
                    <li>Add the web parts to the section and then edit the properties of this web part.</li>
                    <li>Click on the button to 'Manage Tab Labels' and then specify the labels for each web part using the property control.</li>
                  </ul> 
              </div>
            </div>
          </div>
        </div>`;
      }
  }

  protected onInit(): Promise<void> {
    return super.onInit();
    // return this._getEnvironmentMessage().then(message => {
    //   this._environmentMessage = message;
    // });
  }

  // private _getEnvironmentMessage(): Promise<string> {
  //   if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
  //     return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
  //       .then(context => {
  //         let environmentMessage: string = '';
  //         switch (context.app.host.name) {
  //           case 'Office': // running in Office
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
  //             break;
  //           case 'Outlook': // running in Outlook
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
  //             break;
  //           case 'Teams': // running in Teams
  //           case 'TeamsModern':
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
  //             break;
  //           default:
  //             environmentMessage = strings.UnknownEnvironment;
  //         }

  //         return environmentMessage;
  //       });
  //   }

  //   return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  // }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // this._isDarkTheme = !!currentTheme.isInverted;
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

  private getZones(): Array<[string,string]> {
    const zones = new Array<[string,string]>();

    const tabWebPartID = $(this.domElement).closest("div." + this.properties.webpartClass).attr("id");       
    const zoneDIV = $(this.domElement).closest("div." + this.properties.sectionClass);
    let count = 1;
    $(zoneDIV).find("."+this.properties.webpartClass).each(function(){
      const thisWPID = $(this).attr("id");
      if (thisWPID !== tabWebPartID)
      {
        const zoneId = $(this).attr("id");
        const zoneName:string = "Web Part " + count;
        count++;
        zones.push([zoneId || "", zoneName]);
      }
    });

    console.log("zones", zones);

    return zones;
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
                PropertyPaneTextField('sectionClass', {
                  label: strings.SectionClass,
                  description: "Class identifier for Page Section, don't touch this if you don't know what it means."
                }),
                PropertyPaneTextField('webpartClass', {
                  label: strings.WebPartClass,
                  description: "Class identifier for Web Part, don't touch this if you don't know what it means."
                }),
                PropertyFieldCollectionData("tabData", {
                  key: "tabData",
                  label: strings.TabLabels,
                  panelHeader: "Specify Labels for Tabs",
                  manageBtnLabel: "Manage Tab Labels",
                  value: this.properties.tabData,
                  fields: [
                    {
                      id: "WebPartID",
                      title: "Web Part",
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: this.getZones().map((zone:[string,string]) => {
                        return {
                          key: zone["0"],
                          text: zone["1"],
                        };
                      })
                    },
                    {
                      id: "TabLabel",
                      title: "Tab Label",
                      type: CustomCollectionFieldType.string
                    }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
