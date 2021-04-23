import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";

import styles from './CopyTodayPnPWebPart.module.scss';
import * as strings from 'CopyTodayPnPWebPartStrings';

export interface ICopyTodayPnPWebPartProps {
  description: string;
}

export default class CopyTodayPnPWebPart extends BaseClientSideWebPart<ICopyTodayPnPWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = '<div id="siteContent"></div>';

    // Get the Title and Miles from the existing item, copy it to a new item in the Running list with today's date
    const existingItemId: number = parseInt(this.getQueryVariable('ItemID'));

    // Get the item to be copied by ID
    sp.web.lists.getByTitle("Running").items.getById(existingItemId).get().then((existingItem: any) => {
      // Create the new item based on the metadata from the existing item
      sp.web.lists.getByTitle("Running").items.add({
        Title: existingItem.Title,
        Date: new Date(), // Use the current date instead
        Miles: existingItem.Miles,
      }).then((iar: IItemAddResult) => {
        let content:string = '<div><h1>Item successfully copied to today!</h1></div>';
        this.domElement.querySelector("#siteContent").innerHTML = content;
        console.log(iar);
      });
    });
  }

  public getQueryVariable(variable) : string {
    var query = window.location.search.substring(1);
    var vars = query.split('&');
    for (var i = 0; i < vars.length; i++) {
        var pair = vars[i].split('=');
        if (decodeURIComponent(pair[0]) == variable) {
            return decodeURIComponent(pair[1]);
        }
    }
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
