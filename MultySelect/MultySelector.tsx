import * as React from "react";
import AsyncSelect from 'react-select/async';

export interface IProps {
	initialValues: any;
	isControlVisible: boolean;
	isControlDisabled: boolean;
	displayValueField: any;
	displayFieldLabel: any;	
	columns: any;
	odatafilter:any;
	topCount: any;
	orderby:any;
	filterField: any;
	entityName: any;
    value: string;
	onChange: (value:string) => void;
	onSearch: (value:string) => void;	
	records: any
}

export interface IState {
    value: string;
}

export class MultySelectControl extends React.Component<IProps, IState> {
		
    constructor(props: Readonly<IProps>) {
		super(props);	
		this.state = { value: props.value};     
	}

	componentWillReceiveProps(p: IProps) 
	{
		this.setState({value : (p.value)});
    }

	onChange = (ob: any) =>
	{
		if (ob == null) {
			this.setState({value : "" });
			this.props.onChange("");
			return;
		}

		var res = ob.map((e: any) => e[this.props.displayValueField]).join(",");

		this.setState({value : res });
		this.props.onChange(res);
	}

	loadOptions = async (inputValue: string) => {
		const res = this.props.onSearch(inputValue);
		return res;		
	}	
	
	public render(): React.JSX.Element 
	{		
		const selectStyles = { menuPortal: (zindex: any) => ({ ...zindex, zIndex: 9999}) };

		if (this.props.isControlVisible)
		{
        return (
			<div>
			<AsyncSelect
			isMulti={true}
			menuPortalTarget={document.body}
			defaultOptions
			styles={selectStyles}
			getOptionLabel={e => e[this.props.displayFieldLabel]}
			getOptionValue={e => e[this.props.displayValueField]}	
            //@ts-ignore
			loadOptions={this.loadOptions}
			isDisabled={this.props.isControlDisabled}
			onChange={this.onChange}
			defaultValue={this.props.initialValues}
			/></div>
		)
		}
		else{
			return (<></>);
		}
	}
}


