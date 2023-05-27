import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
//import { DynamicForm } from "@pnp/spfx-controls-react/lib/DynamicForm";
//import { DynamicForm} from "../../../../xLib/DynamicForm";
//import { DynamicForm} from "../../../../../sp-dev-fx-controls/lib/DynamicForm";
//import ReactQuill from 'react-quill'; //, { Quill,editor }
//import { DynamicForm } from "../../../controls/DynamicForm/DynamicForm";
import 'react-quill/dist/quill.snow.css';

import styles from './HelloWorld.module.scss';
//import { Button } from 'antd';
import 'antd/dist/reset.css';

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import "@pnp/sp/items/index";
import "@pnp/sp/items/get-all";
import "@pnp/sp/batching";
import "@pnp/sp/fields";

import { SPFI  } from '@pnp/sp'; 
import { getSP } from '../../../PnpjsConfig';

import {  MessageBar, PrimaryButton, Stack, TextField, Label, IStackTokens, } from 'office-ui-fabric-react';
//import {  MessageBar } from 'office-ui-fabric-react';

//https://www.c-sharpcorner.com/article/spfx-form-customizer-extension-to-customize-sharepoint-neweditdisplay-form-of/
export interface IHelloWorldProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

//create state
export interface IHelloWorldState {
  showmessageBar:boolean; //to show/hide message bar on success
  itemObject:any;
 }

{/* <link
  rel="stylesheet"
  href="https://unpkg.com/react-quill@1.3.3/dist/quill.snow.css"
/> */}
const stackTokens: IStackTokens = { childrenGap: 40 };
const LOG_SOURCE: string = 'HelloWorld';
/* const modules = {
  toolbar: {
      container: "#toolbar",
  }
}
const formats = [
'font','size',
'bold','italic','underline','strike',
'color','background',
'script',
'header','blockquote','code-block',
'indent','list',
'direction','align',
'link','image','video','formula',
] */

const ReactQuill = require('react-quill');
//const { Quill } = ReactQuill;

var _sp: SPFI;

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {

  constructor(props: IHelloWorldProps,state:IHelloWorldState) {
    super(props);
    this.state = {showmessageBar:false,itemObject:{}};

    _sp = getSP();
  }


  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: HelloWorld mounted');
    if(this.props.context.itemId)
    {
      this.getCurrentItem();
    }
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: HelloWorld unmounted');
  }

  public render(): React.ReactElement<{}> {
    return <div className={styles.helloWorld}>
       <TextField required onChange={evt => this.updateTitleValue(evt)} value={this.state.itemObject.title} label="Add Title" />
       <TextField required onChange={evt => this.updateDescriptionValue(evt)} value={this.state.itemObject.Desc} label="Add Description" multiline/>
       
        <ReactQuill theme="snow"
          value={this.state.itemObject.MyRichtext}
          onChange={(evt: any) => this.updateMyRichtextValue(evt)}
          //modules={modules}
          //formats={formats}
        />
       {/* <TextField style={{"height":"200px"}} required onChange={evt => this.updateMyRichtextValue(evt)} value={this.state.itemObject.MyRichtext} label="Add My Richtext" multiline/> */}
       <Label>Item id: {this.props.context.itemId}</Label>

      <br/>

      <Stack horizontal tokens={stackTokens}>
      <PrimaryButton text="Create New Item" onClick={()=>this.createNewItem()}  />
      <PrimaryButton text="Reset" onClick={()=>this.resetControls()}  />
    </Stack> 

    {/* <DynamicForm 
          context={this.props.context as any} 
          listId={this.props.context.list.guid.toString()}  
          listItemId={this.props.context.itemId}
          onCancelled={this.props.onClose}
          hiddenFields={["ID", "Description"]}
          onBeforeSubmit={async (listItem) => { return false; }}
          onSubmitError={(listItem, error) => { alert(error.message); }} 
          onSubmitted={this.props.onSave}>
    </DynamicForm>  */}

    <br/>
    {this.state.showmessageBar &&
            <MessageBar   onDismiss={()=>this.setState({showmessageBar:false})}
              dismissButtonAriaLabel="Close">
              "Item saved Sucessfully..."
          </MessageBar>
    }

    </div>;
  }

  private async createNewItem(){
    console.log("Listname: " + this.props.context.list.title);
    const iar: any = await _sp.web.lists.getByTitle(this.props.context.list.title).items.add({
      Title: this.state.itemObject.title + new Date(),
      Description: this.state.itemObject.Desc,
      MyRichtext: this.state.itemObject.MyRichtext
    });
    
    console.log(iar);
    this.setState({showmessageBar:true});
    //this.props.onSave();
  } 

  private async getCurrentItem(){
    console.log("Listname: " + this.props.context.list.title);
    console.log("ID: " + this.props.context.itemId);
    console.log("SP1: " + await _sp.web.lists.getByTitle(this.props.context.list.title));
    console.log("SP2: " + await _sp.web.lists.getByTitle(this.props.context.list.title));
    const iar: any = await _sp.web.lists.getByTitle(this.props.context.list.title).items.getById(this.props.context.itemId)(); 

    console.log(iar);
    console.log(iar.Description);
    console.log(iar.title);
    console.log(iar.MyRichtext);
    
    this.setState({
      itemObject: {title:iar.Title,Desc:iar.Description,MyRichtext:iar.MyRichtext}
    });
    //this.props.onSave();
  }

  private updateTitleValue(evt: any) {
    var item = this.state.itemObject;
    item.title = evt.target.value;
    this.setState({
      itemObject: item
    });
  }

  private updateMyRichtextValue(evt: any) {
    /* const { onChange } = this.props;
    const newState: any = {}; // eslint-disable-line @typescript-eslint/no-explicit-any
    const quill = this.getEditor();
    if (quill) {
      const range = quill.getSelection();
      if (range) {
        const formats = quill.getFormat(range);

        if (!isEqual(formats, this.state.formats)) {
          console.log(`current format: ${formats.list}`);
          newState.formats = formats;
        }
      }

    }
    // do we need to pass this to a handler?
    if (onChange) {
      // yes, get the changed text from the handler
      const newText: string = onChange(value);
      newState.text = newText;
    } else {
      // no, write the text to the state
      newState.text = value;
    }
    this.setState({
      ...newState
    }); */

    console.log("EVT: ");
    console.log(evt);
    console.log("STATE: ");
    console.log(this.state)
    var item = this.state.itemObject;
    item.MyRichtext = evt;
    this.setState({
      itemObject: item
    });
  }

  private updateDescriptionValue(evt: any) {
    var item = this.state.itemObject;
    item.Desc = evt.target.value;
    this.setState({
      itemObject: item
    });
  }

  private async resetControls(){
    var item = this.state.itemObject;
    item.title = "";
    item.Desc = "";
    item.MyRichtext = "";
    this.setState({
      itemObject: item
    });
  } 
}
