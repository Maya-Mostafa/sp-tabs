/* eslint-disable guard-for-in */
/* eslint-disable @typescript-eslint/no-for-in-array */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ModernHillbillyTabs4WebPartStrings';
import ModernHillbillyTabs4 from './components/ModernHillbillyTabs4';
import { IModernHillbillyTabs4Props } from './components/IModernHillbillyTabs4Props';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import * as $ from 'jquery';
import { ColorPicker, DefaultButton, IColor, Icon, Label } from '@fluentui/react';
import { IconPicker } from '@pnp/spfx-controls-react/lib/IconPicker';

export interface IModernHillbillyTabs4WebPartProps {
  description: string;
  sectionClass: string;
  webpartClass: string;
  displayMode: any;
  tabData: any[];
  pgWebparts: any[];
  tabStyle: string;
  tabAlign: string;
}

export default class ModernHillbillyTabs4WebPart extends BaseClientSideWebPart<IModernHillbillyTabs4WebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IModernHillbillyTabs4Props> = React.createElement(
      ModernHillbillyTabs4,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        
        sectionClass: this.properties.sectionClass,
        webpartClass: this.properties.webpartClass,
        tabData: this.properties.tabData,
        
        displayMode: this.displayMode,
        pgWebparts: this.properties.pgWebparts,
        tabStyle: this.properties.tabStyle,
        tabAlign: this.properties.tabAlign
      }
    );

    ReactDom.render(element, this.domElement);
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

  private getWebparts() : Array<[string,string]>{
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

    this.properties.pgWebparts = zones;

    return zones;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('tabStyle', {
                  label: 'Tabs Style',
                  options: [
                    {key: 'tabStyle1', text: 'Ordinary tabs'},
                    {key: 'tabStyle2', text: 'Bottom border with shade'},
                    {key: 'tabStyle3', text: 'Top border with shade'},
                    {key: 'tabStyle4', text: 'Bottom border'},
                    {key: 'tabStyle5', text: 'Circles'},
                    {key: 'tabStyle6', text: 'Colored'},
                  ]
                }),
                PropertyPaneDropdown('tabAlign',{
                  label: 'Tabs Alignment',
                  options: [
                    {key: 'alignLeft', text: 'Left'},
                    {key: 'alignCenter', text: 'Center'},
                    {key: 'alignRight', text: 'Right'},
                  ]
                }),
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
                      // options: this.getWebparts(),
                      options: this.getWebparts().map((zone:[string,string]) => {
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
                    },
                    {
                      id: "TabBgColor",
                      title: "Tab Bg Color",
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                        return(
                          React.createElement('div', {className: 'pickColorDiv'},                                                                                    
                            React.createElement(ColorPicker, {
                              color: value,
                              styles: {
                                colorSquare :{},
                                panel: { padding: 0 },
                                colorRectangle: { minWidth: 170, minHeight: 120 },
                              },
                              showPreview: true,
                              onChange: (ev: React.SyntheticEvent<HTMLElement>, color: IColor) => {
                                onUpdate(field.id, color)
                              },
                              key: 'customColorFieldId1'
                            })
                          )
                        )
                      }
                    },
                    {
                      id: "TabForColor",
                      title: "Tab Text Color",
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                        return(
                          React.createElement('div', {className: 'pickColorDiv'},                                                                                    
                            React.createElement(ColorPicker, {
                              color: value,
                              styles: {
                                colorSquare :{},
                                panel: { padding: 0 },
                                colorRectangle: { minWidth: 170, minHeight: 120 },
                              },
                              showPreview: true,
                              onChange: (ev: React.SyntheticEvent<HTMLElement>, color: IColor) => {
                                onUpdate(field.id, color)
                              },
                              key: 'customColorFieldId2'
                            })
                          )
                        )
                      }
                    },
                    {id: "TabIcon", title: 'Icon', type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                        return (
                          React.createElement('div', {className: 'customIconDiv'},
                            item.TabIcon ?
                            React.createElement(Icon, {iconName: item.TabIcon})
                            :
                            React.createElement(Label, {
                              className: 'fileTextbox', 
                            }, 'No Icon selected'
                            ),
                            React.createElement(IconPicker, {
                              key: itemId,
                              currentIcon: value,
                              buttonLabel: 'Select Icon',
                              onChange: (iconName: string) => {
                                onUpdate(field.id, iconName);                                
                              },
                              onSave: (iconName: string) => {
                                onUpdate(field.id, iconName);
                              },
                              onCancel: () => {
                                onUpdate(field.id, ''); 
                              }
                            }),
                            React.createElement(DefaultButton, {
                              className: 'resetIconBtn', 
                              primary: false, 
                              onClick: () => {
                                onUpdate(field.id, ''); 
                              }
                            }, 'Reset Icon')
                          )
                        );
                      }
                    },
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
