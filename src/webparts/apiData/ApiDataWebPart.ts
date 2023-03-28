import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ApiDataWebPart.module.scss';
import * as strings from 'ApiDataWebPartStrings';
export interface IApiDataWebPartProps {
  description: string;
}
import * as $ from "jquery";
const t = $;
var renderThisCtrl: any;
import 'datatables.net';
import 'datatables.net-dt';
SPComponentLoader.loadCss("https://cdn.datatables.net/1.12.1/css/jquery.dataTables.min.css");

export default class ApiDataWebPart extends BaseClientSideWebPart<IApiDataWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  public table:any;
  public render(): void {
    renderThisCtrl = this;    
    this.domElement.innerHTML = `
                                  <table id="tblfirst" style="display:none;">
                                    <thead style="display:none;">
                                      <tr role="row">
                                        <th></th>
                                        <th></th>
                                      </tr>
                                    </thead>
                                    <tbody>
                                      <tr>
                                        <td id="supH"></td>
                                        <td id="supN"></td>
                                      </tr>
                                      <tr>
                                        <td></td>
                                        <td></td>
                                      </tr>
                                      <tr>
                                        <td id="genH"></td>
                                        <td id="genE"></td>
                                      </tr>                        
                                    </tbody>
                                  </table>
                                  <table id="tblContacts" class="display table-responsive-md no-footer dataTable" role="grid">
                                    <thead>
                                      <tr role="row" id="trheader">
                                        
                                      </tr>
                                    </thead>
                                    <tbody id="tblbody">                        
                                    </tbody>
                                </table>`;
        var currentUser = this.context.pageContext.legacyPageContext.userEmail;
        var webUrl = this.context.pageContext.web.absoluteUrl;
        var siteName = this.context.pageContext.web.title;
        var siteId = this.context.pageContext.site.id.toString();
        this.getData(currentUser,webUrl,siteName,siteId);
  }

  private getData(user:string,webUrl:string,siteName:string,siteId:string){
    var bodyContent = "{\r\n \"Email\":\"" + user + "\",\r\n \"SiteId\": \"" + siteId + "\",\r\n \"WebUrl\": \"" + webUrl + "\",\r\n \"SiteName\": \"" + siteName + "\"\r\n}";
    $.ajax({
      url: "https://prod-50.westeurope.logic.azure.com/workflows/1d6ee1ea0b944b6da9b901f71e3cd6fa/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=JsGPZzEu0n59m9Lb4G6C-VhyuuO_ewZROORLE0LVTBE",
      crossDomain: true,
      method: "POST",
      data: bodyContent,
      processData: false,
      headers: {
          "content-type": "application/json",
          "cache-control": "no-cache"
      },
      success: function (data) {
          if(data){
            // var pass = btoa(data.Key + ':' + data.Value);
            if(data.MainHeader && data.MainHeader.length > 0){
              $('#tblfirst').show();
              $('#supH').html('<b>' + data.MainHeader[0].Title || "" + ': </b>');
              $('#genH').html('<b>' + data.MainHeader[1].Title || "" + ': </b>');
              $('#supN').html(data.MainHeader[0].Value || "");
              $('#genE').html(data.MainHeader[1].Value || "");
            }          
                console.log(data);
                if(data.DataRows.length > 0){
                  if (renderThisCtrl.table instanceof (<any>$.fn.dataTable).Api) {
                    $('#tblContacts').DataTable().clear();
                    $('#tblContacts').DataTable().destroy();
                  } 
                  var headerArray = data.DataRows;
                  $('#trheader').html('');
                  $('#tblbody').html('');
                  for (let index = 0; index < headerArray.length; index++) {
                    const element = headerArray[index];
                    if(index == 0)
                      $('#trheader').append('<th class="sorting_asc" tabindex="0" aria-controls="tblAll" rowspan="1" colspan="1" aria-label="File: activate to sort column descending" aria-sort="ascending">' + element.Title || "" + '</th>');
                    else
                    $('#trheader').append('<th class="sorting" tabindex="0" aria-controls="tblAll" rowspan="1" colspan="1" aria-label="File: activate to sort column descending" aria-sort="ascending">' + element.Title || "" + '</th>');
                  }           
                  var element2 = data.Results;     
                  for (let index2 = 0; index2 < element2.length; index2++) {
                    var tr = `<tr role="row" class="odd">`;
                    for (let index = 0; index < data.DataRows.length; index++) {
                      var temp = data.DataRows[index].Name || "";
                      var columnArr:any,dataVal:any;
                      if(temp.indexOf('.') > -1){
                        columnArr = temp.split('.');
                        dataVal = element2[index2][columnArr[0]][columnArr[1]] || "";
                      }
                      else
                        dataVal = element2[index2][temp] || "";
                      var control = renderThisCtrl.getControl(data.DataRows[index].Prop.Type,data.DataRows[index].Prop.Style,data.DataRows[index].Prop.Size,dataVal);
                      if(index == 0){
                        tr += `<td class="sorting_1">` + control + `</td>`;
                      }
                      else{
                        tr += `<td>` + control + `</td>`;                         
                      }                     
                    }
                    
                    tr += `</tr>`;
                    $('#tblbody').append(tr);
                  }
                  // (<any>$.fn.dataTable).ext.order['dom-text'] = function (settings:any, col:any) {
                  //   return this.api()
                  //       .column(col, { order: 'index' })
                  //       .nodes()
                  //       .map(function (td:any, i:any) {
                  //           return $('input', td).val();
                  //       });
                  // };
                  renderThisCtrl.table = $('#tblContacts').DataTable({
                    paging: true,
                    ordering: true,
                    processing: true,
                    "lengthChange": false,
                    "info": false
                  //   columns:[
                  //     { orderDataType: 'dom-text', type: 'string' }
                  // ],
                  });
                }                
          }
      },
      error: function (err) {
        console.log(err);
      }
    });           
  }

  public getControl(type:string,style:string,size:string,data:string){
    var HTMLElement = '';
    switch (type) {
      case "Text":
        HTMLElement = '<label style="' + style +'">' + data + '</label>';
        break;
      case "SPProfilePic":
        var url = renderThisCtrl.context.pageContext.web.absoluteUrl + "/_layouts/15/userphoto.aspx?size=" + size + "&accountname=" + data;
        var url_d = renderThisCtrl.context.pageContext.web.absoluteUrl + "/_layouts/15/userphoto.aspx?size=" + size;
        HTMLElement = '<img src="' + url + '" style="' + style +'" onerror="this.onerror=null;this.src=\'' + url_d + '\'"  />';
        break;
      case "Number":
        HTMLElement = '<label style="' + style +'">' + data + '</label>';
        break;
      default:
        break;
    }
    return HTMLElement;
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
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

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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
