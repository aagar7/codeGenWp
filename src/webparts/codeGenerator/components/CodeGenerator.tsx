import * as React from 'react';
import styles from './CodeGenerator.module.scss';
import { ICodeGeneratorProps } from './ICodeGeneratorProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {List} from 'stryker.common.framework';

class JsonObject{
  public Columns:ColumnsData[];
  public Data:object[][];
}

class ColumnsData{
  public name:string;
  public sortable:boolean;
  public filterable:boolean;
  public type:string;
}

export default class CodeGenerator extends React.Component<ICodeGeneratorProps, {}> {

  componentDidMount(){
    this.getData('https://strykerglobaltechcenter.sharepoint.com', 'Site Provisioning',['Title','u_SiteType','u_SiteUrl']);
  }

  public getData(webUrl: string, listName: string, fieldsToGet: string[]){  
    return List.getListItems(listName, fieldsToGet, undefined, undefined, undefined, undefined, undefined, webUrl, undefined).then(t=>{
      var jsonObject = new JsonObject();
      jsonObject.Data = new Array<Array<object>>();
      jsonObject.Columns = new Array<ColumnsData>();
      if( fieldsToGet!=undefined && fieldsToGet!=null){
        fieldsToGet.forEach(field=>{
          var fieldData = new ColumnsData();
          fieldData.name = field;
          fieldData.sortable = true;
          fieldData.filterable = true;
          fieldData.type = undefined;
          jsonObject.Columns.push(fieldData);
        });
      }
      t.forEach(item => {
        var values = new Array<object>();
        fieldsToGet.forEach(field => {
          values.push(item[field]);
        });
        jsonObject.Data.push(values);
      });
      console.debug(jsonObject);
    }).catch(t=>null);
  }
  
  public render(): React.ReactElement<ICodeGeneratorProps> {    
    return (
      <div className={ styles.codeGenerator }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
