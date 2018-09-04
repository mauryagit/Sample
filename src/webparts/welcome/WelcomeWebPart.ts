import { 
  Version,Environment,EnvironmentType
   } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel,
  PropertyPaneDropdown,
  PropertyPaneChoiceGroup,
  PropertyPaneCheckbox,
  PropertyPaneLink,  
  IWebPartPropertiesMetadata
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './WelcomeWebPart.module.scss';
import * as strings from 'WelcomeWebPartStrings';
import MockHttpClient from "./MockHttpClient";
import {
  SPHttpClient,
  SPHttpClientResponse,
  SPHttpClientConfiguration
} from "@microsoft/sp-http";
import { 
  version 
} from 'react';

export interface IWelcomeWebPartProps {
  description: string;
  customDescription :string;
  greeting :string;
  getsomechoice:string;
  version :string;
}

export interface ISPLists{
  value : ISPList[];
}
export  interface ISPList{
  Title :string;
  Id:number;
}

export default class WelcomeWebPart extends BaseClientSideWebPart<IWelcomeWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.welcome }">
        <div class="${ styles.container }">
        Version : ${this.dataVersion}
          <div class="${ styles.row }">
         
            <div class="${ styles.column }">
           
              <span class="${ styles.title }">${escape(this.properties.description)} ${this.context.pageContext.user.displayName}</span><br/>
              <span class="${styles.subTitle}">${escape(this.properties.greeting)}</span>   
              <span class="${styles.subTitle}">${escape(this.properties.getsomechoice)}</span>   
              <p class="${styles.paragraph}">${escape(this.properties.customDescription)}</p>               
              <p class="${styles.paragraph}">Loading from ${escape(this.context.pageContext.web.title)}</p>
              
              
            </div>
          </div>
          <div id="spListContainer"/>
        </div>
      </div>`;
      this._renderListAsync();
  }

  protected get propertiesMetadata() : IWebPartPropertiesMetadata{
    return {
        'title': {isSearchablePlainText:true},
        'intro' : {isHtmlString:true},
        'image' : {isImageSource:true},
        'url' :{isLink:true}
    };
  }

  private _getMockListData() : Promise<ISPLists>{
    return MockHttpClient.get()
    .then ((data:ISPList[]) =>{
        var listData  : ISPLists = {value : data};
        return listData;
    }) as Promise<ISPLists>;
  }

  private _renderListAsync():void{
    if(Environment.type == EnvironmentType.Local){
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    }else if (Environment.type == EnvironmentType.SharePoint || Environment.type == EnvironmentType.ClassicSharePoint){
      this._getListData().then((response) => {
        this._renderList(response.value);
      })
    }
  }
  private _renderList(items: ISPList[]):void{
      let html:string ="";
      items.forEach((item:ISPList) => {
        html +=` <ul class="${styles.list}">
        <li class="${styles.listItem}">
        <span class="ms-font-l">${item.Title}</span>
        </li></ul>`;
      });
      const listContainer :Element = this.domElement.querySelector("#spListContainer");
      listContainer.innerHTML=html;

  }
  private _getListData(): Promise<ISPLists>{
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`,SPHttpClient.configurations.v1)
    .then((response : SPHttpClientResponse) => {
      return response.json();
    })
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
                }),
                PropertyPaneTextField('customDescription',{
                  label: strings.CustomFieldLabel, multiline:true
                }),
                PropertyPaneDropdown("greeting",{
                  label : strings.greetings,
                  options:[
                    {key: "Good Morning", text:"Good Morning"},
                    {key: "Good Afternoon", text:"Good Afternoon"},
                    {key: "Good Night", text:"Good Night"}
                  ]
                })
                ,
                PropertyPaneChoiceGroup("getsomechoice",{
                  label:strings.choice,options:[
                    {key:"Red", text:"Red"},
                    {key:"Yellow", text:"Yellow"},
                    {key:"Green", text:"Green"}
                  ]
                })
                ,
                PropertyPaneLabel("version",{text : strings.Version})
              ]
            }
        ,
            {
              groupName: strings.AdvancedGroupName,
              groupFields:[
               PropertyPaneCheckbox("checkone",
              {
                text:"abc"
              }),
              PropertyPaneChoiceGroup("filetype",
            {
              label: "Choice", options:[
                { key: 'Word', text: 'Word',
                imageSrc: 'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/docx_32x1.png',
                imageSize: { width: 32, height: 32 },
                selectedImageSrc: 'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/docx_32x1.png'
              },
              { key: 'Excel', text: 'Excel',
                imageSrc: 'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/xlsx_32x1.png',
                imageSize: { width: 32, height: 32 },
                selectedImageSrc: 'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/xlsx_32x1.png'
              },
              { key: 'PowerPoint', text: 'PowerPoint',
                imageSrc: 'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/pptx_32x1.png',
                imageSize: { width: 32, height: 32 },
                selectedImageSrc: 'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/pptx_32x1.png'
              },
              { key: 'OneNote', text: 'OneNote',
                imageSrc: 'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/one_32x1.png',
                imageSize: { width: 32, height: 32 },
                selectedImageSrc: 'https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/png/one_32x1.png'
              }
              ]
            }),
            PropertyPaneLink('link',{
              href:"http://www.google.com",target:"_blank", text:"Google",
              popupWindowProps:{
                title:"Google", width:500, height:500,positionWindowPosition:2
              }
            })
              ]
            }
          ], displayGroupsAsAccordion: true
        }        
      ]
    };
  }
}
