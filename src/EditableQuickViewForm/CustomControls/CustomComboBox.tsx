import * as React from "react";
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

initializeIcons(undefined, { disableWarnings: true });

export interface IComboBoxState extends React.ComponentState, IComboBoxProps { currentValue: number | undefined, currentError: string|undefined }
export interface IComboBoxProps {
	id: string,
	type: string,
	label: string,
	required: boolean,
	disabled: boolean,
	value: number | undefined,
	inputChanged:Function|null;
	inputError:Function|null;
	options:Array<any>
	isRequiredErrorMessage: string
}
export class CustomComboBox  extends React.Component<IComboBoxProps, IComboBoxState> {
  	constructor(props: IComboBoxProps) {
      	super(props);

		this.state = {
			id: props.id,
			type: props.type,
			label: props.label,
			required: props.required,
			disabled: props.disabled,
			value: props.value,
			inputChanged: props.inputChanged,
			inputError:  props.inputError,
			options: props.options,
			isRequiredErrorMessage: props.isRequiredErrorMessage,
			currentError: undefined,
			currentValue: props.value,
		};

		this.items = props.options.map(o => {
			return {key: o.value, text: o.label}
		});
	}
	  
	private items:IDropdownOption[];

  	render(): JSX.Element {
		return (
			<div>
				<Dropdown
					placeholder="---"
					ariaLabel=""
					// label={this.props.label}
					// required={this.props.required}
					defaultSelectedKey={this.props.value}
					disabled={this.props.disabled}
					options={this.items}
					// id={this.props.id}
					onChange={this.inputChanged}
					onBlur={this.inputBlur}
				/>
				{!!this.state.currentError && 
				<MessageBar messageBarType={MessageBarType.error} isMultiline={false}>{this.state.currentError}</MessageBar>}
			</div>
		);
  	}

	private inputChanged = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption | undefined, index?: number | undefined): void => {
		var val:string | number | string[] | undefined = undefined;
		if(option && option.key) val = option.key as number;
		this.setState({ currentValue: val });
		
		if(this.props.inputChanged){
			if(this.props.type == "boolean"){
				this.props.inputChanged(this.props.id, this.props.type, "", val);
			}else{
				this.props.inputChanged(this.props.id, this.props.type, "", val || undefined);
			}
		}
	};

	private inputBlur = (): void => {
		var error:string|undefined = undefined;

		if(this.state.required && !this.state.currentValue){
			error = this.state.isRequiredErrorMessage;
		}

		this.setState({ currentError: error });
		
		if(this.props.inputError){
			this.props.inputError(this.props.id??"", this.state.currentError);
		}
	}
}