
import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { IProps, MultySelectControl } from "./MultySelector";

export class MultySelect implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _existingValues: any
	private _value: any;
	private _notifyOutputChanged:() => void;
	private _container: HTMLDivElement;
	private props: IProps = 
	{ 
		value : "", 
		onChange : this.notifyChange.bind(this),
		onSearch : this.notifySearch.bind(this),	
		initialValues : undefined,	
		records: [],
		displayValueField: "",
		displayFieldLabel: "",
		orderby:"",
		columns: "",
		odatafilter: "",
		topCount: "",
		filterField: "",
		entityName: "",
		isControlVisible: true,
		isControlDisabled: true	
	};
	private _context: ComponentFramework.Context<IInputs>;

	constructor()
	{

	}

	public async init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
        
		this._context = context;
		this._notifyOutputChanged = notifyOutputChanged;
		this._container = document.createElement("div");
		this.props.value = context.parameters.value.raw || "";	
		this.props.entityName = context.parameters.entityName.raw || "";
		this.props.filterField = context.parameters.filterField.raw || "";
		this.props.orderby = context.parameters.orderby.raw || "";
		this.props.topCount = context.parameters.topCount.raw || "";
		this.props.columns = context.parameters.columns.raw || "";
		this.props.odatafilter = context.parameters.odatafilter.raw || "";
		this.props.displayFieldLabel = context.parameters.displayFieldLabel.raw || "";
		this.props.displayValueField = context.parameters.displayValueField.raw || "";
		console.log("init");
		if(this.props.value.length > 0)
		{
			this.props.initialValues = await this.onLoad();
			console.log(JSON.stringify(this.props.initialValues));
			this.updateView(context);			
		}	
		else
		{
			this.props.initialValues =[];
		}	
		
		container.appendChild(this._container);
		
	}

	notifyChange(newValue: string) 
	{
		this._value = newValue;
		this._notifyOutputChanged();
	}

	async notifySearch(newValue: string)
	{
		var filteredOdata = this.props.odatafilter ? " "+ this.props.odatafilter : "";
		var query = `?$select=${this.props.columns}&$filter=contains(${this.props.filterField}, '${newValue}')${filteredOdata}&$top=${this.props.topCount}`;
		if(this.props.orderby != "") query += `&$orderby=${this.props.orderby}`;

		console.log(this.props.entityName,query);

		return this._context.webAPI.retrieveMultipleRecords(this.props.entityName,query)
		.then(function (results) {		
				return results?.entities;		
		})
	}

	public async onLoad()
	{			
			var count = 0;
			var qs = `?$select=${this.props.columns}&$filter=`;		

			this.props.value.split(",").forEach(c=>{			
				if (count > 0)
				{
					qs = qs + ' or ' + this.props.displayValueField + " eq '" + c + "'";
				}
				else
				{
					qs = qs + this.props.displayValueField + " eq '" + c + "'";
				}
				count++;
			});
				return this._context.webAPI.retrieveMultipleRecords(this.props.entityName,qs)
				.then(function (results) {		
				return results?.entities;
			});				
	}

	private renderElement()
	{
		if(this.props.initialValues != undefined)
		{
			this.props.isControlDisabled = this._context.mode.isControlDisabled;
			this.props.isControlVisible = this._context.mode.isVisible;
	
				ReactDOM.render(
					React.createElement(MultySelectControl, this.props)
					, this._container
				);
		}
		
	}
    
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		this._value = context.parameters.value.raw;
		this.props.value = this._value;
		this.props.orderby = context.parameters.orderby.raw;
		this.props.topCount = context.parameters.topCount.raw;
		this.props.columns = context.parameters.columns.raw;
		this.props.filterField = context.parameters.filterField.raw;
		this.props.odatafilter = context.parameters.odatafilter.raw;
		this.props.displayFieldLabel = context.parameters.displayFieldLabel.raw;
		this.props.displayValueField = context.parameters.displayValueField.raw;
		this.props.entityName =context.parameters.entityName.raw;
			
		this.renderElement();
	}
    
	public getOutputs(): IOutputs
	{
		return {
			value : this._value
		};
	}
    
	public destroy(): void
	{
	}
}
	