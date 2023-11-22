/* eslint-disable */
import { IInputs } from "./generated/ManifestTypes";
import { elementContains, IPersonaProps } from "@fluentui/react";
import { resolve } from "path";
import { rejects } from "assert";

export interface IFilterValues {
    filter1?:string,
    filter2?:string,
    filter3?:string,
    filter4?:string,
    filter5?:string,
}


export class Service {
    private context : ComponentFramework.Context<IInputs>;
    private debugMode: boolean;
    private entityName: string;
    private relatedEntityDisplay:string;
    private relatedEntityFilter:string;
    private relatedEntityValue:string;
    private exampleData : IPersonaProps[] = [
        {
            key: 1,
            imageInitials: 'PV',
            text: 'Italia',
            secondaryText: 'Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            title: 'ejemplo',
            id: 'ITA'
        },
        {
            key: 2,
            imageInitials: 'AR',
            text: 'España',
            secondaryText: 'Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            id:'ES'
        },
        {
            key: 3,
            imageInitials: 'AL',
            text: 'Alex Lundberg',
            secondaryText: 'Software Developer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
            id:'MA'
        },
        {
            key: 4,
            imageInitials: 'RK',
            text: 'Roko Kolar',
            secondaryText: 'Financial Analyst',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
        },
        {
            key: 5,
            imageInitials: 'CB',
            text: 'Christian Bergqvist',
            secondaryText: 'Sr. Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
        },
        {
            key: 6,
            imageInitials: 'VL',
            text: 'Valentina Lovric',
            secondaryText: 'Design Developer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
        },
        {
            key: 7,
            imageInitials: 'MS',
            text: 'Maor Sharett',
            secondaryText: 'UX Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
        },
        {
            key: 8,
            imageInitials: 'PV',
            text: 'Anny Lindqvist',
            secondaryText: 'Designer',
            tertiaryText: 'In a meeting',
            optionalText: 'Available at 4:00pm',
        }
    ];
    
    constructor(_context:ComponentFramework.Context<IInputs>, _debugMode: boolean){
        this.context = _context;
        this.debugMode = _debugMode;
        this.entityName = _context.parameters.entityName.raw || "",
        this.relatedEntityDisplay = _context.parameters.relatedEntityDisplay.raw || "",
        this.relatedEntityFilter = _context.parameters.relatedEntityFilter.raw || "",
        this.relatedEntityValue = _context.parameters.relatedEntityValue.raw || ""
    }

    private getFilteredQuery():string{
        let query = this.context.parameters.odatafilter.raw || '';
        for (let i = 1; i < 5; i++) {
            if(query?.includes('{' + i + '}')){
                //@ts-ignore
                var actualFilterValue = this.context.parameters['filter'+i] ? "'" + this.context.parameters['filter'+i].raw + "'": '';
                query = query.replace('{' + i + '}',actualFilterValue);
                
            }
        }
        return query;
   }

   
   private async retriveRecords(entityType:string, query : string){
    console.log("retrived entity " + entityType + " query: " + query);
    return await this.context.webAPI.retrieveMultipleRecords(entityType, query).then(
        function success(result) {
            return result;
        },
        function (error) {
            return null;
        })

   }

   public itemsChanged(itemsChanged:IFilterValues):boolean{
        let changed = false;
        for (let i = 1; i < 5; i++) {
            //@ts-ignore
            let currentValue = itemsChanged['filter'+i] || "";
            //@ts-ignore
            let value = this.context.parameters['filter'+i].raw || "";
            console.log(this.entityName + ": currentValue filter"+i+" : " + currentValue + " value : "+value + " changed: " + (currentValue != value))
            if(!changed && currentValue != value) changed = true;
        }
        return changed;
   }

   public setFilterItems(context: ComponentFramework.Context<IInputs>):IFilterValues{
    return {
        filter1:context.parameters.filter1.raw || '',
        filter2:context.parameters.filter2.raw || '',
        filter3:context.parameters.filter3.raw || '',
        filter4:context.parameters.filter4.raw || '',
        filter5:context.parameters.filter5.raw || '',
    }
   }

   public transformItemsOfResponse(response:ComponentFramework.WebApi.RetrieveMultipleResponse | null):IPersonaProps[]{
        let itemsArray : IPersonaProps[] = [];
        if(response){
            response.entities.map((element, index) => {
                itemsArray.push(
                    {
                        id:element[this.relatedEntityValue],
                        text:element[this.relatedEntityDisplay],
                        key:index
                    }
                )
            });
        }
        return itemsArray;
   }

    public async getEntity(value:string){
        let data : IPersonaProps[] = [];

        if(this.debugMode){
           data = this.exampleData.filter(data => data.id?.toUpperCase() === value.toUpperCase());
        }else{
            let fullQuery = "?$filter=contains("+ this.relatedEntityValue +",'"+value+"') "+ this.getFilteredQuery(); 
            let response = this.retriveRecords(this.entityName, fullQuery);
            return await response.then((response) => {
               return data = this.transformItemsOfResponse(response);
            })

        }
        return data;
    }

    public async getFirstElementSelected(value:string){
        let data : IPersonaProps[] = [];

        if(this.debugMode){
           data = this.exampleData.filter(data => data.id?.toUpperCase() === value.toUpperCase());
        }else{
            let fullQuery = "?$filter="+ this.relatedEntityValue +" eq '"+value+"'" + this.getFilteredQuery(); 
            let response = this.retriveRecords(this.entityName, fullQuery);
            return await response.then((response) => {
               return data = this.transformItemsOfResponse(response);
            })

        }
        return data;
    }

    public async getEntityList(value:string){
        let data : IPersonaProps[] = []
        if(this.debugMode){
           data = this.exampleData;
        }else{
            let fullQuery = "?$filter=contains("+ this.relatedEntityDisplay +",'"+value+"') "+ this.getFilteredQuery(); 
            let response = this.retriveRecords(this.entityName, fullQuery);
            return await response.then((response) => {
              return  data = this.transformItemsOfResponse(response);
            })
        }
        if(value) return data.filter(entity => entity.text?.toUpperCase().includes(value.toUpperCase())) || [];
        else return data;
    }
}