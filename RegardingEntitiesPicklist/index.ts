import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { IPickListProps, PickList } from "./PickList";
import {IFilterValues, Service} from './service';
import * as React from "react";
import { IPersonaProps } from "@fluentui/react/lib/Persona";

export class RegardingEntitiesPicklist implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;
    private peopleList: IPersonaProps[];
    private fieldValue: string;
    private debugMode : boolean = false;
    private selectedItem?:IPersonaProps[];
    private Service: Service;
    private itemsChanged:IFilterValues;
    
    constructor() { 
    //@ts-ignore
        if(typeof Xrm === 'undefined'){
            this.debugMode = true;
        }
    }
    //@ts-ignore
    async onChangePeopleList(person: IPersonaProps[]): void  {
        if(person && person.length > 0){
            this.selectedItem = await this.Service.getEntity(person[0].id || '')
            this.fieldValue = person[0].id || '';
        }else{
            this.fieldValue = ''
            this.selectedItem = []
        }
        this.notifyOutputChanged();
    }
    //@ts-ignore
    async onChangeSuggestion(searchText:string):IPersonaProps[]{
        return await this.Service.getEntityList(searchText) ||[];
    }
    
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        this.Service = new Service(context, this.debugMode);
        this.itemsChanged = this.Service.setFilterItems(context);
        this.Service.getFirstElementSelected(context.parameters.value.raw || '').then(
            (response) =>{
                this.selectedItem = response;
                console.log('RETRIVED DATA: ', response)
                context.factory.requestRender()
        });
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
         
        let changed = this.Service.itemsChanged(this.itemsChanged);
        if(changed){
            this.selectedItem = [];
            this.fieldValue = '';
            this.itemsChanged = this.Service.setFilterItems(context);
            this.notifyOutputChanged()
        }
        const props: IPickListProps = 
        { 
            name:context.parameters.value.attributes?.DisplayName,
            peopleList:this.peopleList, 
            selectedItem:this.selectedItem,
            onChangePeopleList:this.onChangePeopleList.bind(this),
            onChangeSuggestion:this.onChangeSuggestion.bind(this),
            disabled: context.mode.isControlDisabled,
            service: this.Service,
            pcfValue:context.parameters.value.raw || '',
        };

        return React.createElement(
            PickList,
            props
        );
    }

    public getOutputs(): IOutputs {
        return {
            value: this.fieldValue
         };
    }

    public destroy(): void {
    }
}
