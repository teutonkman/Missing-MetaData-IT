import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import { sp, Web } from "@pnp/sp";

import styles from './MissingRequiredInfoWebPart.module.scss';
import * as strings from 'MissingRequiredInfoWebPartStrings';

export interface IMissingRequiredInfoWebPartProps {
  description: string;
}

export default class MissingRequiredInfoWebPart extends BaseClientSideWebPart<IMissingRequiredInfoWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.missingRequiredInfo }">
        <div class="">Libraries with missing required information</div>
        <div class="${ styles.row }">
          <div id="spListContainer"></div>
        </div>
      </div>`;
      this._renderListDataAsync();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }


  public _renderListDataAsync():void {
    let html:string=``;
    let web = new Web("https://jaggedpeak.sharepoint.com/sites/InformationTechnology");

    web.lists.getByTitle("Content Types").items
    .select('Libraries')
    .orderBy('Libraries', true)
    .get().then((lib) => {  
      
      for(var x=0; x < lib.length; x++){  
        let library = lib[x].Libraries;
        let status = 'good';
        

        web.lists.getByTitle(library).items
        //.expand(CreatedBy, Modified_x0020_By)
        .select('FileLeafRef, Record_x0020_Type, Created_x0020_By/UserName, Modified_x0020_By/UserName')
        .get().then((data) => {  
          html+=`<section><ul>${library}`;
          for(var i=0; i < data.length; i++){  
            if (!data[i].Record_x0020_Type){ 
              let createdBy = data[i].Created_x0020_By;
              createdBy = createdBy.replace("i:0#.f|membership|", "");
              createdBy = createdBy.replace("@jaggedpeakenergy.com", "");
              let modifiedBy = data[i].Modified_x0020_By;
              modifiedBy = modifiedBy.replace("i:0#.f|membership|", "");
              modifiedBy = modifiedBy.replace("@jaggedpeakenergy.com", "");
            html+=`<li><span class="${ styles.bad } ${ styles.spacing } styles.spacing }">${data[i].FileLeafRef}</span><span class="${ styles.bad } ${ styles.spacing }">CreatedBy: ${createdBy}</span><span class="${ styles.bad } ${ styles.spacing }">ModifiedBy: ${modifiedBy}</span></li>`;
            status = 'bad';
            }   
          }
          if (status === 'good'){
            html+=`<li><span class="${ styles.good }">good</span></li>`;
          }
          html+=`</ul></section>`;
          document.getElementById("spListContainer").innerHTML = html;
        

        }).catch((data) => {  
        console.log(data);  
        });
        
      }
      

    }).catch((data) => {  
    console.log(data);  
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
