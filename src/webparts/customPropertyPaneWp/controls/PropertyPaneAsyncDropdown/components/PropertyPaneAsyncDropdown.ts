import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {IPropertyPaneField,
    PropertyPaneFieldType} from '@microsoft/sp-webpart-base';
import {IDropdownOption} from 'office-ui-fabric-react/lib/components/Dropdown';
import { IPropertyPaneAsyncDropdownProps } from './IPropertyPaneAsyncDropdownProps';
import { IPropertyPaneAsyncDropdownInternalProps } from './IPropertyPaneAsyncDropdownInternalProps';
import AsyncDropdown from './AsyncDropdown';
import { IAsyncDropdownProps } from './IAsyncDropdownProps';

export class PropertyPaneAsyncDropdown implements IPropertyPaneField<IPropertyPaneAsyncDropdownProps>{
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public targetProperty :string;
    public properties : IPropertyPaneAsyncDropdownInternalProps;
    private elm : HTMLElement;

    constructor(targetProperty: string, properties : IPropertyPaneAsyncDropdownProps){
        this.targetProperty=targetProperty;
        this.properties={
            key: properties.label,
            label:properties.label,
            loadOptions : properties.loadOptions,
            onPropertyChange : properties.onPropertyChange,
            selectedKey: properties.selectedKey,
            disabled: properties.disabled,
            onRender:this.onRender.bind(this)              
        };
    }
    public render() :void{
        if(!this.elm){
            return;
        }
        this.onRender(this.elm);
    }
    private onRender(elm: HTMLElement):void{
        if(!this.elm){
            this.elm = elm;
        }
        const element : React.ReactElement<IAsyncDropdownProps> =React.createElement(AsyncDropdown,{
            label: this.properties.label,
            loadOptions : this.properties.loadOptions,
            onChanged: this.onChanged.bind(this),
             selectedKey: this.properties.selectedKey,
              disabled: this.properties.disabled,    
              stateKey: new Date().toString()
        });
        ReactDOM.render(element,elm);
    }

    private onChanged(option: IDropdownOption,index? :number):void{
        this.properties.onPropertyChange(this.targetProperty,option.key);
    }
}