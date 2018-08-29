import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './EmbedVisioWebPart.module.scss';
import * as jQuery from 'jquery';
import * as strings from 'EmbedVisioWebPartStrings';

export interface IEmbedVisioWebPartProps {
  embedUrl: string;
}
declare var OfficeExtension : any;
export default class EmbedVisioWebPart extends BaseClientSideWebPart<IEmbedVisioWebPartProps> {
  private _session:any;
  public constructor(){
    super();
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div id="iframeHost">
      </div>`;
      if(this.properties.embedUrl){
        this.onPropertyPaneConfigurationComplete();
      }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges():boolean{
    return true;
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
              groupFields: [
                PropertyPaneTextField('embedUrl', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneConfigurationComplete(){
    let url;
    url = this.properties.embedUrl.replace("action=view","action=embedview");
    url = this.properties.embedUrl.replace("action=interactivepreview","action=embedview");
    url = this.properties.embedUrl.replace("action=default","action=embedview");
    url = this.properties.embedUrl.replace("action=edit","action=embedview");
    
    jQuery.getScript("https://appsforoffice.microsoft.com/embedded/1.0/visio-web-embedded.js", function(){

   this._session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });

   return this._session.init().then(function () {
    // Initilization is successful 
    console.log("Initilization is successful");
    });
   
   });
  
  }
}
