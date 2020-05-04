import * as React from 'react';
import { IDatePickerProps, CustomDatePicker } from "./CustomDatePicker";
import { ITextFieldProps, CustomTextBox } from "./CustomTextBox"
import { IComboBoxProps, CustomComboBox } from "./CustomComboBox"
import { ILookupBoxProps, CustomLookupBox } from './CustomLookupBox';
import { DefaultButton, BaseButton, Button } from 'office-ui-fabric-react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

export interface IFormError{
	id: string,
	errorText: string,
	errorType: MessageBarType
}
export interface ILoader{
    loading: boolean,
    text: string
}
export interface IFormProps {
    title: string,
    controls: Array<IFormControl>,
    canSave: boolean,
    save: Function,
    loader: ILoader,
    saveError: string,

    formErrors: Array<IFormError>,
}
export interface IComboValue{
    value: Number,
    label: string
}
export interface IControlStrings{
//     months: [ "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" ],
//     shortMonths: [ "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" ],
//     days: [ "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" ],
//     shortDays: ["S", "M", "T", "W", "T", "F", "S"],
//     goToToday: "Go to today",
//     prevMonthAriaLabel: "Go to previous month",
//     nextMonthAriaLabel: "Go to next month",
//     prevYearAriaLabel: "Go to previous year",
//     nextYearAriaLabel: "Go to next year",
//     closeButtonAriaLabel: "Close date picker",
//     isRequiredErrorMessage: 'Field is required.',
//     invalidInputErrorMessage: 'Invalid date format.',
    months: string[],
    shortMonths: string[],
    days: string[],
    shortDays: string[],
    goToToday: string,
    prevMonthAriaLabel: string,
    nextMonthAriaLabel: string,
    prevYearAriaLabel: string,
    nextYearAriaLabel: string,
    closeButtonAriaLabel: string,
    isRequiredErrorMessage: string,
    invalidInputErrorMessage: string,
    dateFormat: string,
    currency: string,
    decimalPrecision: number,
    decimalSeparator: string,

    savingText:string,
    searchingText:string,
    loadingText:string,
    suggestionsHeaderText: string,
    noResultsFoundText: string,

    lookupEmptyMessage: string,
    formNotFoundMessage: string,
    formEntityMismatchMessage: string
}
export interface IFormControl{
    id: string,
    type: string,
    format: string,
    label: string,
    recommended: boolean,
    required: boolean,
    disabled: boolean,
    value: any,
    randomid: string,

    //Events
    inputChanged:Function|null,
    inputError:Function|null,

    //Dates
    dateFormat:string|undefined,

    //Lookups
    inputAutocomplete:Function|null,

    //Picklists
    options: Array<IComboValue>,

    strings: IControlStrings,
}
 
export interface IFormState extends React.ComponentState, IFormProps {  }
export class Form extends React.Component<IFormProps> {
    constructor(props: IFormProps) {
        super(props);
    }
    render() {
        try{
            return (
                
                <div className="mainDiv">
                    {
                        this.props.formErrors.length > 0 &&
                        this.getMessages()
                    }
                    {
                        this.props.formErrors.length == 0 &&
                        this.props.loader.loading &&
                        <div className="loader">
                            <table className="loaderContent">
                                <tbody>
                                    <tr>
                                        <td>
                                            <img className="loaderProgress" src="/_imgs/advfind/progress.gif" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td className="loaderText">{this.props.loader.text}</td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    }
                    {
                        this.props.formErrors.length == 0 &&
                        !this.props.loader.loading &&
                        <div className="formHeaderRow">
                            <div className="formTitle">{this.props.title}</div>
                            <div className="formButtonRow">{
                                    this.props.canSave && 
                                    <DefaultButton 
                                        iconProps= {{ iconName: 'Save' }}
                                        className="formButton" 
                                        onClick={this.saveForm.bind(this)}>Save</DefaultButton>
                                }
                            </div>
                        </div>
                    }
                    {
                        this.props.formErrors.length == 0 &&
                        this.props.saveError && 
                        <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>{this.props.saveError}</MessageBar>
                    }
                    <div className="container-fluid">
                        {
                        this.props.formErrors.length == 0 &&
                        !this.props.loader.loading && 
                        this.getFormContent()}
                    </div>
                </div>
                
            );
        }catch(exc){
            debugger;
        }
    }

    private saveForm(event: React.MouseEvent<HTMLAnchorElement | HTMLButtonElement | HTMLDivElement | BaseButton | Button | HTMLSpanElement, MouseEvent>){
        if(this.props.save)
            this.props.save();
    }
   
    private getMessages(){
        return this.props.formErrors.map(err => {
            var className = "";
            switch(err.errorType){
                case MessageBarType.info: className = "infoMessageBar-text"; break;
            }
            return <MessageBar className={className} messageBarType={err.errorType} isMultiline={true}>{err.errorText}</MessageBar>
        })
    }
    private getFormContent(){
        return this.props.controls.map(field => {
            switch(field.type){
                case "datetime":
                    return (
                        <div className="formRow" key={field.id}>
                            <div className="formLabel">
                                <div className="formLabelText">{field.label}</div>
                                {field.required && <div className="formLabelRequired">*</div>}
                                {field.recommended && <div className="formLabelRecommended">+</div>}
                            </div>
                            <div className="formField">{this.getDateControl(field)}</div>
                        </div>
                    );
                case "string":
                case "money":
                case "decimal":
                case "float":
                case "integer":
                        return (
                        <div className="formRow" key={field.id}>
                            <div className="col-xs-3 formLabel">
                                <div className="formLabelText">{field.label}</div>
                                {field.required && <div className="formLabelRequired">*</div>}
                                {field.recommended && <div className="formLabelRecommended">+</div>}
                            </div>
                            <div className="col-xs-9 formField">{this.getTextControl(field)}</div>
                        </div>
                    );
                case "picklist":
                case "boolean":
                    return (
                        <div className="formRow" key={field.id}>
                            <div className="col-xs-3 formLabel">
                                <div className="formLabelText">{field.label}</div>
                                {field.required && <div className="formLabelRequired">*</div>}
                                {field.recommended && <div className="formLabelRecommended">+</div>}
                            </div>
                            <div className="col-xs-9 formField">{this.getComboControl(field)}</div>
                        </div>
                    );
                case "lookup":
                    return (
                        <div className="formRow" key={field.id}>
                            <div className="col-xs-3 formLabel">
                                <div className="formLabelText">{field.label}</div>
                                {field.required && <div className="formLabelRequired">*</div>}
                                {field.recommended && <div className="formLabelRecommended">+</div>}
                            </div>
                            <div className="col-xs-9 formField">{this.getLookupControl(field)}</div>
                        </div>
                    );
                default: 
                    return (
                        <div key={field.id}>
                            
                        </div>
                    ); 
            }
            
        });
    }
    private getDateControl(field: IFormControl){
        var date:Date|undefined = undefined;
        if(field.value){
            date = new Date(field.value); 
            if(date.toString() == "Invalid Date") date = undefined;
        }

        var p:IDatePickerProps = {
            id: field.id,
            type: field.type,
            format: field.format,
            label: field.label,
            required: field.required,
            disabled: field.disabled,
            value: date,
            inputChanged: field.inputChanged,
            inputError: field.inputError,
            dateStrings: field.strings,
            dateFormat: field.strings.dateFormat,
            isRequiredErrorMessage: field.strings.isRequiredErrorMessage,
        } 
        return React.createElement(CustomDatePicker, p)
    }
    private getTextControl(field: IFormControl){
        var str:string = "";
        if(field.value) str = field.value.toString();

        var p:ITextFieldProps = {
            id: field.id,
            type: field.type,
            format: field.format,
            label: field.label,
            required: field.required,
            disabled: field.disabled,
            value: str,
            randomid: field.randomid,
            inputChanged: field.inputChanged,
            inputError: field.inputError,
            isRequiredErrorMessage: field.strings.isRequiredErrorMessage,
            invalidInputErrorMessage: field.strings.invalidInputErrorMessage,
            currency: field.strings.currency,
            decimalPrecision: field.strings.decimalPrecision,
            decimalSeparator: field.strings.decimalSeparator
        } 
        return React.createElement(CustomTextBox, p)
    }
    private getComboControl(field: IFormControl){
        var val:number|undefined = undefined;
        if(field.type == "boolean"){
            val = field.value ? 1 : 0;
        }else{
            if(field.value){
                var n = parseInt(field.value.toString());
                if(!isNaN(n)) val = n;
            }
        }

        var p:IComboBoxProps = {
            id: field.id,
            type: field.type,
            label: field.label,
            required: field.required,
            disabled: field.disabled,
            value: val,
            inputChanged: field.inputChanged,
            inputError: field.inputError,
            options: field.options,
            isRequiredErrorMessage: field.strings.isRequiredErrorMessage
        } 
        return React.createElement(CustomComboBox, p)
    }
    private getLookupControl(field: IFormControl){
        var date:Date|undefined = new Date(field.value);
        if(date.toString() == "Invalid Date") date = undefined;

        var p:ILookupBoxProps = {
            id: field.id,
            type: field.type,
            label: field.label,
            required: field.required,
            disabled: field.disabled,
            inputChanged: field.inputChanged,
            inputError: field.inputError,
            value: field.value,
            randomid: field.randomid,
            inputAutocomplete: field.inputAutocomplete,
            isRequiredErrorMessage: field.strings.isRequiredErrorMessage,
            loadingText: field.strings.searchingText,
            noResultsFoundText: field.strings.noResultsFoundText,
            suggestionsHeaderText: field.strings.suggestionsHeaderText,
        } 
        return React.createElement(CustomLookupBox, p)
    }
}