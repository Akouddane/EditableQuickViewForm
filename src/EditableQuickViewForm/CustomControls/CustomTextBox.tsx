import * as React from "react";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

initializeIcons(undefined, { disableWarnings: true });

const controlClass = mergeStyleSets({
	control: {
	  margin: '0 0 15px 0',
	  maxWidth: '300px',
	},
});

export interface ITextFieldState extends React.ComponentState, ITextFieldProps { currentValue: string, currentError: string|undefined }
export interface ITextFieldProps {
	id: string,
	type: string,
	format: string,
	label: string,
	required: boolean,
	disabled: boolean,
	value: string|undefined,
	randomid: string,
	inputChanged:Function|null,
	inputError:Function|null,
	isRequiredErrorMessage: string,
	invalidInputErrorMessage: string,
    currency: string,
	decimalPrecision: number,
	decimalSeparator: string,
}
export class CustomTextBox  extends React.Component<ITextFieldProps, ITextFieldState> {
  	constructor(props: ITextFieldProps) {
      	super(props);

		var currValue = props.value??"";
		if(["money", "decimal", "float"].indexOf(props.type??"") > -1){
			var value = parseFloat(props.value??"");
			if(!isNaN(value)){
				currValue = this.formatNumber(value, props.type == "money")
			}
		}
		

		this.state = {
			id: props.id,
			type: props.type,
			format: props.format,
			label: props.label,
			required: props.required,
			disabled: props.disabled,
			value: props.value,
			randomid: props.randomid,
			inputChanged: props.inputChanged,
			inputError:  props.inputError,
			isRequiredErrorMessage: props.isRequiredErrorMessage,
			invalidInputErrorMessage: props.invalidInputErrorMessage,
			currency: props.currency,
			decimalPrecision: props.decimalPrecision,
			decimalSeparator: props.decimalSeparator,
			
			currentError: undefined,
			currentValue: currValue,
		};

  	}

	//TODO : use props.format for (Text, Email, Phone, URL, TextArea, Ticker)
	// Email : Accounts
	// Phone : Phone
	// URL : Globe
	// Ticker : Financial
  	render(): JSX.Element {
		return (
			<div>
				<TextField
					placeholder="---"
					id={this.props.randomid}
					name={this.props.randomid}
					disabled={this.props.disabled}
					value={this.state.currentValue}
					iconProps={{ iconName: this.getIcon() }}
					multiline={this.isMultiLine()}
					onChange={this.inputChanged}
					onFocus={this.inputFocus}
					onBlur={this.inputBlur}
				/>
				{!!this.state.currentError && 
				<MessageBar messageBarType={MessageBarType.error} isMultiline={false}>{this.state.currentError}</MessageBar>}
			</div>
		);
  	}

	private inputChanged = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string | undefined): void => {
		this.setState({ currentValue: newValue || '' });
		if(this.props.inputChanged){
			this.props.inputChanged(this.props.id??"", this.props.type, this.props.format, newValue || '');
		}
	};

	private inputBlur = (): void => {
		var error:string|undefined = undefined;

		if(!this.state.currentValue){
			if(this.props.inputChanged){
				this.props.inputChanged(this.props.id??"", this.props.type, this.props.format, undefined);
			}
			if(this.state.required)
				error = this.props.isRequiredErrorMessage;
		}
		else{
			switch(this.state.type){
				case "money":
				case "decimal":
				case "float":
					var value = parseFloat(this.state.currentValue);
					if(isNaN(value)){
						error = this.props.invalidInputErrorMessage;
						if(this.props.inputChanged){
							this.props.inputChanged(this.props.id??"", this.props.type, this.props.format, undefined);
						}
					}else{
						this.setState({ currentValue: this.formatNumber(value, this.state.type == "money") });
						if(this.props.inputChanged){
							this.props.inputChanged(this.props.id??"", this.props.type, this.props.format, value);
						}		
					}
					break;
				case "integer":
					var value = parseInt(this.state.currentValue);
					if(isNaN(value)){
						error = this.props.invalidInputErrorMessage;
						if(this.props.inputChanged){
							this.props.inputChanged(this.props.id??"", this.props.type, this.props.format, null);
						}
					}else{
						this.setState({ currentValue: value.toString() });
						if(this.props.inputChanged){
							this.props.inputChanged(this.props.id??"", this.props.type, this.props.format, value);
						}		
					}
					break;
			}
		}


		this.setState({ currentError: error });
		
		if(this.props.inputError){
			this.props.inputError(this.props.id??"", this.state.currentError);
		}
	}
	private inputFocus = (): void => {
		switch(this.state.type){
			case "money":
			case "decimal":
			case "float":
				var value = parseFloat(this.unFormatNumber(this.state.currentValue||""));
				if(isNaN(value)){
					this.setState({ currentValue: "" });
				}else{
					this.setState({ currentValue: value.toString() });
				}
				break;
		}
	}

	private getIcon(){
		switch(this.props.format){
			case "Email": return "Accounts";
			case "Phone": return "Phone";
			case "Url": return "Globe";
			case "Ticker": return "Financial";
		}
		return "Text";
	}
	private isMultiLine(){
		return this.props.format == "TextArea";
	}

	private replaceAll(str: string, rep: string, by: string){
		while(str.indexOf(rep) > -1)
			str = str.replace(rep, by);
		return str;
	}
	private unFormatNumber(formattedValue: string){
		return formattedValue ? this.replaceAll(this.replaceAll(formattedValue, this.props.currency, ""), " ", "").trim() : "";
	}
	private formatNumber(number: number, isMoney: boolean){
		var decimal = this.props.decimalSeparator;
		var thousands = " ";
		var currency = this.props.currency;

		var formatted = number.toFixed(this.props.decimalPrecision);
		formatted = this.replaceAll(formatted, ".", "_");
		formatted = this.replaceAll(formatted, ",", "_");
		formatted = this.replaceAll(formatted, "_", decimal);
		formatted = formatted.replace(/\B(?=(\d{3})+(?!\d))/g, thousands);

		if(isMoney) formatted = formatted + " " + currency;

		return formatted;
	}
}