import * as React from "react";
import { IBasePickerSuggestionsProps, IPickerItemProps, CompactPeoplePicker, ValidationState, Autofill } from 'office-ui-fabric-react/lib/Pickers';
import { IPersonaProps, Persona } from 'office-ui-fabric-react/lib/Persona';
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

export interface ILookup{
	entityType: string, 
	name: string, 
	id: string,
	entityImg: string|undefined,
}
export interface IPersonaObj extends IPersonaProps { lookup: ILookup }
export interface ILookupBoxState extends React.ComponentState, ILookupBoxProps { currentValue: IPersonaProps[]|undefined, currentError: string|undefined }
export interface ILookupBoxProps {
	id: string,
	type: string,
	label: string,
	required: boolean,
	disabled: boolean,
	value: IPersonaProps[]|undefined,
	randomid: string,
	inputChanged:Function|null,
	inputError:Function|null,
	inputAutocomplete: Function|null,
	isRequiredErrorMessage: string,
	loadingText:string,
    suggestionsHeaderText: string,
    noResultsFoundText: string
}
export class CustomLookupBox  extends React.Component<ILookupBoxProps, ILookupBoxState> {
	
  	constructor(props: ILookupBoxProps) {
      	super(props);

		this.state = {
			id: props.id,
			type: props.type,
			label: props.label,
			required: props.required,
			disabled: props.disabled,
			value: props.value,
			randomid: props.randomid,
			inputChanged: props.inputChanged,
			inputError:  props.inputError,
			inputAutocomplete: props.inputAutocomplete,
			currentError: undefined,
			currentValue: props.value,
			isRequiredErrorMessage: props.isRequiredErrorMessage,
			loadingText: props.loadingText,
			suggestionsHeaderText: props.suggestionsHeaderText,
			noResultsFoundText: props.noResultsFoundText,
		};

		this.suggestionProps.suggestionsHeaderText = props.suggestionsHeaderText;
		this.suggestionProps.mostRecentlyUsedHeaderText = props.suggestionsHeaderText;
		this.suggestionProps.suggestionsAvailableAlertText = props.suggestionsHeaderText;
		this.suggestionProps.suggestionsContainerAriaLabel = props.suggestionsHeaderText;
		this.suggestionProps.noResultsFoundText = props.noResultsFoundText;
		this.suggestionProps.loadingText = props.loadingText;
  	}
	 
	private suggestionProps: IBasePickerSuggestionsProps = {
		suggestionsHeaderText: 'Suggested records',
		mostRecentlyUsedHeaderText: 'Suggested records',
		noResultsFoundText: 'No results found',
		loadingText: 'Loading',
		showRemoveButtons: false,
		suggestionsAvailableAlertText: 'Suggestions available',
		suggestionsContainerAriaLabel: 'Suggested records',
	};

  	render(): JSX.Element {
		return (
			<div>
				<CompactPeoplePicker
					inputProps={{
						placeholder: "---",
						id:this.props.randomid,
						name:this.props.randomid,
					}}
					// onRenderItem={this.onRenderItem}
					className={'ms-PeoplePicker'}
					itemLimit={1}
					defaultSelectedItems={this.props.value}
					getTextFromItem={this.getTextFromItem}
					disabled={this.props.disabled}
					onValidateInput={this.validateInput}
					onResolveSuggestions={this.onFilterChanged}
					onEmptyInputFocus={this.onFilterEmpty}
					pickerSuggestionsProps={this.suggestionProps}
					onChange={this.inputChanged}
					onBlur={this.inputBlur}
					
				/>
				{!!this.state.currentError && 
				<MessageBar messageBarType={MessageBarType.error} isMultiline={false}>{this.state.currentError}</MessageBar>}
			</div>
		);
  	}

	
	private onRenderItem(props: IPickerItemProps<IPersonaProps>):JSX.Element{
		if(props.item.imageUrl){
			return <Persona 
				imageUrl={props.item.imageUrl}
				text={props.item.text}
				/>
		}else{

		}
		return <div></div>;
	}
	
	private onFilterChanged = (filterText: string, currentPersonas?: IPersonaProps[] | undefined, limitResults?: number) => 
	{
		if(this.props?.inputAutocomplete){
			return this.props?.inputAutocomplete(this.props.id??"", filterText||"");
		}
	};
	private onFilterEmpty = (selectedItems?: IPersonaProps[] | undefined) => {
		if(this.props?.inputAutocomplete){
			return this.props?.inputAutocomplete(this.props.id??"", "", 3);
		}
	}
	private inputChanged = (items?: IPersonaProps[] | undefined) => {
		var error = undefined;
		this.setState({currentValue:items});
		if(items && items.length > 0){
			if(this.props.inputChanged){
				var val = undefined;
				if(items.length > 0){
					val = items.map(a => (a as IPersonaObj).lookup);
				}
				this.props.inputChanged(this.props.id??"", this.props.type, "", val);
			}
		}
		else{
			if(this.props.inputChanged){
				this.props.inputChanged(this.props.id??"", this.props.type, "", []);
			}

			if(this.props.required){
				error = this.props.isRequiredErrorMessage;
			}
		}

		this.setState({currentError : error});
	}
	private inputBlur = (event: React.FocusEvent<HTMLInputElement | Autofill>) => {
		var error = undefined;

		if(this.props.required){
			if(!this.state.currentValue || this.state.currentValue.length == 0){
				error = this.props.isRequiredErrorMessage;
			}
		}

		this.setState({currentError : error});
	}
	
	private getTextFromItem(persona: IPersonaProps): string {
		return persona.text as string;
	}
	private validateInput(input: string): ValidationState {
		if (input.indexOf('@') !== -1) {
		  return ValidationState.valid;
		} else if (input.length > 1) {
		  return ValidationState.warning;
		} else {
		  return ValidationState.invalid;
		}
	}

	

	
	
}