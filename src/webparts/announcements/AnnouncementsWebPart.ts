import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  IPropertyPaneDropdownProps,
  PropertyPaneToggle,
  PropertyPaneSlider,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

import * as strings from 'AnnouncementsWebPartStrings';
import Announcements from './components/Announcements';
import { IAnnouncementsProps } from './components/IAnnouncementsProps';
import pnp, { sp, Item, ItemAddResult, ItemUpdateResult } from "sp-pnp-js";

export interface IAnnouncementsWebPartProps {
  description: string;
  currentContext:WebPartContext;
  listName:string;
  displayStyle:boolean;
  headerFont:number;
  contentFont:number;
  lines:number;
  color:string;
  headerColor:string;
}
export interface IPropertyControlsTestWebPartProps {
  color: string;
}

var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();

export default class AnnouncementsWebPart extends BaseClientSideWebPart <IAnnouncementsWebPartProps> {
  private lists: IPropertyPaneDropdownOption[];
  private listDropDownDisablede: boolean = true;
  private loadLists(): Promise<IPropertyPaneDropdownOption[]>{
    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) =>{
      sp.web.lists.get()  
        .then((items): void => {  
          if (items.length === 0) {  
            resolve([{
              key:"No List Found",
              text:"No List Found"
            }]);  
          }  
          else {  
            items.map(item =>{
              options.push({
                key:item.Title,
                text:item.Title
              })
              resolve(options);
            }) 
          }  
        }, (error: any): void => {  
          reject(error);  
        });
    });
  } 
  // protected onInit():Promise<void>{
  //   return new Promise<void>((resolve,_reject)=>{
  //     this.properties.description="Announcements";
  //     resolve(undefined);
  //   });
  // }
  public render(): void {
    const element: React.ReactElement<IAnnouncementsProps> = React.createElement(
      Announcements,
      {
        description: this.properties.description,
        currentContext: this.context,
        listName:this.properties.listName,
        displayStyle: this.properties.displayStyle,
        headerFont:this.properties.headerFont,
        contentFont:this.properties.contentFont,
        lines:this.properties.lines,
        color:this.properties.color,
        headerColor:this.properties.headerColor
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
  protected onPropertyPaneConfigurationStart(): void {
    this.listDropDownDisablede = !this.lists;

    if (this.lists) {
      return;
    }

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    this.loadLists()
      .then((listOptions: IPropertyPaneDropdownOption[]): void => {
        this.lists = options;
        this.listDropDownDisablede = false;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });
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
                PropertyPaneDropdown('listName', {
                  label: 'Select List',
                  options: this.lists,
                  disabled: this.listDropDownDisablede
                }),
                PropertyPaneToggle('displayStyle',{
                  key:"displayStyle",
                  label:"Show Heading?",
                  onText:"Yes",
                  offText:"No",
                }),
                PropertyPaneSlider('headerFont',{
                  label:"Select Header font size:",
                  min:1,
                  max:30,
                  step:1,
                  value:14
                }),
                PropertyPaneSlider('contentFont',{
                  label:"Select Content font size:",
                  min:12,
                  max:35,
                  step:1,
                  value:14
                }),
                PropertyPaneSlider('lines',{
                  label:"Enter the Length for Content",
                  min:1,
                  max:20,
                  step:1,
                  value:3
                }),
                PropertyFieldColorPicker('color', {
                  label: 'Color',
                  selectedColor: this.properties.color,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                }),
                PropertyFieldColorPicker('headerColor', {
                  label: 'Header Color',
                  selectedColor: this.properties.headerColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
