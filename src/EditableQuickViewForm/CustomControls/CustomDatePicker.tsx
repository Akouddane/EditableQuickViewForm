import * as React from "react";
import { DatePicker, DayOfWeek, IDatePickerStrings } from "office-ui-fabric-react/lib/DatePicker";
import { VirtualizedComboBox, IComboBox, IComboBoxOption } from "office-ui-fabric-react/lib/ComboBox";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import * as moment from "moment";


initializeIcons(undefined, { disableWarnings: true });

const controlClass = mergeStyleSets({
	control: {
	  margin: '0 0 15px 0',
	  maxWidth: '300px',
	},
});

export interface IDatePickerState extends React.ComponentState, IDatePickerProps { currentError: string|undefined, time: string|undefined }
export interface IDatePickerProps {
	id: string,
	type: string,
	format: string,
	label: string,
	required: boolean,
	disabled: boolean,
	value: Date|undefined,
	inputChanged:Function|null,
	inputError:Function|null,
	dateStrings: IDatePickerStrings|null,
	isRequiredErrorMessage: string,
	dateFormat: string,
}
export class CustomDatePicker extends React.Component<IDatePickerProps, IDatePickerState> {
  	constructor(props: IDatePickerProps) {
      	super(props);

		if(props.dateStrings != null){
			this._defaultDayPickerStrings = props.dateStrings;
		}

		var time = "00:00";
		if(this.props.value && this.props.format == "datetime"){
			var m = moment(this.props.value);
			if(m.isValid()){
				time = m.format("HH:mm");
			}
		}else{
			time = "";
		}

		this.state = {
			id: props.id,
			type: props.type,
			format: props.format,
			label: props.label,
			required: props.required,
			disabled: props.disabled,
			value: props.value,
			inputChanged: props.inputChanged,
			inputError: props.inputError,
			dateStrings: props.dateStrings,
			isRequiredErrorMessage: props.isRequiredErrorMessage,
			dateFormat: props.dateFormat,
			currentValue: props.value,
			currentError: "",
			time: time
		};

		
  	}

  	private _defaultDayPickerStrings: IDatePickerStrings = {
		months: [ "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" ],
		shortMonths: [ "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" ],
		days: [ "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" ],
		shortDays: ["S", "M", "T", "W", "T", "F", "S"],
		goToToday: "Go to today",
		prevMonthAriaLabel: "Go to previous month",
		nextMonthAriaLabel: "Go to next month",
		prevYearAriaLabel: "Go to previous year",
		nextYearAriaLabel: "Go to next year",
		closeButtonAriaLabel: "Close date picker",
		isRequiredErrorMessage: 'Field is required.',
  		invalidInputErrorMessage: 'Invalid date format.',
	};

  	render(): JSX.Element {
		return (
			<div>
				{
					this.props.format == "date" && 
					<DatePicker
						strings={this._defaultDayPickerStrings}
						firstWeekOfYear={1}
						placeholder="---"
						disabled={this.props.disabled}
						value={this.state.currentValue}
						allowTextInput={true}
						// id={this.props.id}
						firstDayOfWeek={DayOfWeek.Monday}
						formatDate={this.formatDate.bind(this)}
						parseDateFromString={this.parseDate.bind(this)}
						showWeekNumbers={false}
						onSelectDate={this.onSelectDate}
					/>
				}
				{
					this.props.format == "datetime" && 
					<div>
						<DatePicker
							strings={this._defaultDayPickerStrings}
							firstWeekOfYear={1}
							placeholder="---"
							disabled={this.props.disabled}
							value={this.state.currentValue}
							allowTextInput={true}
							// id={this.props.id}
							firstDayOfWeek={DayOfWeek.Monday}
							formatDate={this.formatDate.bind(this)}
							parseDateFromString={this.parseDate.bind(this)}
							showWeekNumbers={false}
							onSelectDate={this.onSelectDate}
							styles={this.mixedDateStyles}
							style={{display: 'inline-block'}}
						/>
						<VirtualizedComboBox
							placeholder="---"
							ariaLabel=""
							selectedKey={this.state.time}
							allowFreeform={false}
							disabled={this.props.disabled}
							options={this.getTimeItems()}
							autoComplete={"on"}
							buttonIconProps={{iconName: "BufferTimeBoth"}}
							onChange={this.timeChanged.bind(this)}
							styles={this.mixedComboStyles}
							style={{display: 'inline-block'}}
						/>
					</div>
					
				}
				{!!this.state.currentError && 
				<MessageBar messageBarType={MessageBarType.error} isMultiline={false}>{this.state.currentError}</MessageBar>}
			</div>
		);
  	}

	private mixedDateStyles = {
		root: {
		  maxWidth: '50%',
		  width: '50%',
		},
	};
	private mixedComboStyles = {
		container: {
		  maxWidth: '50%',
		  width: '50%',
		},
	};

	private getTimeItems(){
		var list = [];
		for(var i = 0; i < 24; i++){
			for(var j = 0; j < 60; j++){
				var h = i > 9 ? i.toString() : "0" + i;
				var m = j > 9 ? j.toString() : "0" + j;
				list.push({key: h + ":" + m, text: h + ":" + m});
			}
		}
		return list;
	}
	private formatDate(date:Date|undefined) {
		if(date){
			return moment(date).format(this.props.dateFormat);
		}
			// return date.toLocaleDateString(navigator.languages && navigator.languages[0])
		return ""
	}
	private parseDate = (val: string): Date|null => {
		var d = moment(val, this.props.dateFormat);
		if(d.isValid()){
			return d.toDate();
		}
		return null;
	};
	private onSelectDate = (value: Date | null | undefined): void => {
		if(this.props.inputChanged){
			this.setState({currentValue: value});
			this.props.inputChanged(this.props.id??"", this.props.type, this.props.format, value);
		}
		var error = undefined;

		if(!value){
			if(this.props.required)
				error = this.props.isRequiredErrorMessage;
		}
		
		this.setState({currentError: error});

		if(this.props.inputError){
			this.props.inputError(this.props.id??"", this.state.currentError);
		}
	};
	private timeChanged(event: React.FormEvent<IComboBox>, option?: IComboBoxOption | undefined, index?: number | undefined, value?: string | undefined){
		if(this.state.currentValue && option){
			var newDate = moment(moment(this.state.currentValue).format("YYYY-MM-DD") + " " + option.key).toDate();
			if(this.props.inputChanged){
				this.setState({currentValue: newDate, time: option.text });
				this.props.inputChanged(this.props.id??"", this.props.type, this.props.format, newDate);
			}
		}
	}
}