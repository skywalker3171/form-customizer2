import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import {
  BaseFormCustomizer
} from '@microsoft/sp-listview-extensibility';

import HelloWorld, { IHelloWorldProps } from './components/HelloWorld';
import { getSP } from '../../PnpjsConfig';

/**
 * If your form customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHelloWorldFormCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = 'HelloWorldFormCustomizer';

export default class HelloWorldFormCustomizer
  extends BaseFormCustomizer<IHelloWorldFormCustomizerProperties> {

  public onInit(): Promise<void> {
    // Add your custom initialization to this method. The framework will wait
    // for the returned promise to resolve before rendering the form.
    Log.info(LOG_SOURCE, 'Activated HelloWorldFormCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    getSP(this.context);
    return Promise.resolve();
  }

  private _listItem = {} as any;

  public render(): void {
    // Use this method to perform your custom rendering.

    const helloWorld: React.ReactElement<{}> =
      React.createElement(HelloWorld, {
        context: this.context,
        listGuid: this.context.list.guid,
        itemID: this.context.itemId,
        listItem: this._listItem,
        EditFormUrl: this._getEditFormLink(),
        AddFormUrl: this._getAddFormLink(),
        displayMode: this.displayMode,
        onSave: this._onSave,
        onClose: this._onClose
       } as IHelloWorldProps);

    ReactDOM.render(helloWorld, this.domElement);
  }

  public onDispose(): void {
    // This method should be used to free any resources that were allocated during rendering.
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  private _onSave = (): void => {

    // You MUST call this.formSaved() after you save the form.
    this.formSaved();
  }

  private _onClose =  (): void => {
    // You MUST call this.formClosed() after you close the form.
    this.formClosed();
  }

  private _getEditFormLink = (): string => {

    const tenantUri = window.location.protocol + "//" + window.location.host;
    const EditFormUrl = `${this.context.pageContext.site.absoluteUrl}/_layouts/15/SPListForm.aspx?PageType=6&List=${this.context.list.guid.toString()}&ID=${this.context.itemId}&Source=${tenantUri + this.context.list.serverRelativeUrl}/AllItems.aspx?as=json&ContentTypeId=${this.context.contentType.id}&RootFolder=${this.context.list.serverRelativeUrl}`

    return EditFormUrl;

  }

  private _getAddFormLink = (): string => {

    const tenantUri = window.location.protocol + "//" + window.location.host;
    const AddFormUrl = `${this.context.pageContext.site.absoluteUrl}/_layouts/15/SPListForm.aspx?PageType=8&List=${this.context.list.guid.toString()}&Source=${tenantUri + this.context.list.serverRelativeUrl}/AllItems.aspx&RootFolder=${this.context.list.serverRelativeUrl}&Web=${this.context.pageContext.web.id}`

    return AddFormUrl;

  }
}
