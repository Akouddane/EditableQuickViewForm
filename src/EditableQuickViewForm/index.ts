import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import * as moment from "moment";
import * as X2JS from "x2js";

import { IFormControl, IControlStrings, Form, ILoader, IFormError } from "./CustomControls/Form";
import { IPersonaObj, ILookup } from './CustomControls/CustomLookupBox';
import { MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

interface ILookupRequestGroup{
	max: number|undefined,
	pattern: string,
	requests: Array<ILookupRequest>,
	resolve: (value?: IPersonaObj[] | PromiseLike<IPersonaObj[]> | undefined) => void
}
interface ILookupRequest{
	entity: string,
	done: boolean,
	results: Array<IPersonaObj>,
}
interface IEntityMeta{
	entity: string,
	done: boolean,
	metadata: ComponentFramework.PropertyHelper.EntityMetadata|null
}
interface ILookupImage{
	attribute: string,
	lookup: ILookup|null,
	done: boolean
}
interface IUpdateData{
	attribute: string,
	type: string,
	format: string,
	value: any
}



export class EditableQuickViewForm implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;
	private _formControls: Array<IFormControl>;
	private _dataObject: any;
	private _updateObject: Array<IUpdateData>;
	private _dataErrors: any;
	private _lookupsImages: Array<ILookupImage>;
	private _loader: ILoader;
	private _initialization: boolean;

	private _lookupMetadatas:Array<IEntityMeta>;
	private _entityMetadata:ComponentFramework.PropertyHelper.EntityMetadata;
	private _formDefinedControls: Array<any>;
	private _formEntity: string;
	private _formId: string;
	private _formTitle: string;
	private _fieldValue: ILookup|null;
	private _saveError: string;
	private _formErrors: Array<IFormError>;
	

	private _defaultMaxLookupRequest = 5;

	private lookupTypes:Array<string> = ["lookup","customer","owner"];
	private controlsStrings:IControlStrings = {
		months: [ "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" ],
		shortMonths: [ "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" ],
		days: [ "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" ],
		shortDays: ["S", "M", "T", "W", "T", "F", "S"],
		goToToday: "Today",
		prevMonthAriaLabel: "Prev month",
		nextMonthAriaLabel: "Next month",
		prevYearAriaLabel: "Prev year",
		nextYearAriaLabel: "Next year",
		closeButtonAriaLabel: "Close",
		dateFormat: "MM/DD/YYYY",
		isRequiredErrorMessage: 'This field is required',
		invalidInputErrorMessage: 'Invalid input format',
		currency: '€',
		decimalPrecision: 2,
		decimalSeparator: ",",
		savingText: "Saving...",
		searchingText: "Searching...",
		loadingText: "Loading...",
		noResultsFoundText: "No result",
		suggestionsHeaderText: "Results found",
		lookupEmptyMessage: "Fill '#fieldName' field to activate this content",
		formNotFoundMessage: "Something went wrong in retrieving systemform with id #formid",
		formEntityMismatchMessage: "Lookup Type (#lookupentity) does not match to entity's form (#formentity)",
	};
	private errorTypes = {
		lookupEmptyError: "lookupEmpty",
		formNotFoundError: "formNotFound",
		formEntityMismatch: "formEntityMismatch",
		runtimeError: "runtimeError"
	};
	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		this._initialization = true;
		this._context = context;
		this._container = container;
		this._container.classList.add("customFormContainer");

		this._dataObject = {};
		this._updateObject = [];
		this._dataErrors = {};
		this._formErrors = [];

		this._fieldValue = this.getLookup();
		this._formId = context.parameters.quickFormId.raw??"";
		this._formTitle = context.parameters.formTitle.raw??"";
		this._formControls = [];
		this._loader = {loading: false, text: ""};
		

		this.controlsStrings.isRequiredErrorMessage = this.getResourceString("isRequiredErrorMessage", this.controlsStrings.isRequiredErrorMessage);
		this.controlsStrings.invalidInputErrorMessage = this.getResourceString("invalidInputErrorMessage", this.controlsStrings.invalidInputErrorMessage);
		this.controlsStrings.lookupEmptyMessage = this.getResourceString("lookupEmptyMessage", this.controlsStrings.lookupEmptyMessage);
		
		this.controlsStrings.loadingText = this.getResourceString("loadingText", this.controlsStrings.loadingText);
		this.controlsStrings.searchingText = this.getResourceString("searchingText", this.controlsStrings.searchingText);
		this.controlsStrings.savingText = this.getResourceString("savingText", this.controlsStrings.savingText);
		
		this.controlsStrings.noResultsFoundText = this.getResourceString("noResultsFoundText", this.controlsStrings.noResultsFoundText);
		this.controlsStrings.suggestionsHeaderText = this.getResourceString("suggestionsHeaderText", this.controlsStrings.suggestionsHeaderText);
		
		this.controlsStrings.goToToday = this.getResourceString("goToTodayText", this.controlsStrings.goToToday);
		this.controlsStrings.prevMonthAriaLabel = this.getResourceString("prevMonthText", this.controlsStrings.prevMonthAriaLabel);
		this.controlsStrings.nextMonthAriaLabel = this.getResourceString("nextMonthText", this.controlsStrings.nextMonthAriaLabel);
		this.controlsStrings.prevYearAriaLabel = this.getResourceString("prevYearText", this.controlsStrings.prevYearAriaLabel);
		this.controlsStrings.nextYearAriaLabel = this.getResourceString("nextYearText", this.controlsStrings.nextYearAriaLabel);
		this.controlsStrings.closeButtonAriaLabel = this.getResourceString("closeButtonText", this.controlsStrings.closeButtonAriaLabel);


		this.controlsStrings.lookupEmptyMessage = this.controlsStrings.lookupEmptyMessage.replace("#fieldName", this.getLookupDisplayName());

		if(this._fieldValue == null){
			this.setError({ id: this.errorTypes.lookupEmptyError, errorText: this.controlsStrings.lookupEmptyMessage, errorType: MessageBarType.info});
		}
		
		this.controlsStrings.days = this._context.userSettings.dateFormattingInfo.dayNames;
		this.controlsStrings.shortDays = this._context.userSettings.dateFormattingInfo.abbreviatedDayNames;
		this.controlsStrings.months = this._context.userSettings.dateFormattingInfo.monthNames;
		this.controlsStrings.shortMonths = this._context.userSettings.dateFormattingInfo.abbreviatedMonthNames;
		this.controlsStrings.dateFormat = this._context.userSettings.dateFormattingInfo.shortDatePattern.toUpperCase();
		this.controlsStrings.currency = this._context.userSettings.numberFormattingInfo.currencySymbol;
		this.controlsStrings.decimalPrecision = this._context.userSettings.numberFormattingInfo.currencyDecimalDigits;
		this.controlsStrings.decimalSeparator = this._context.userSettings.numberFormattingInfo.currencyDecimalSeparator;

		this._loader.text = this.controlsStrings.loadingText;
		this._loader.loading = true;
		this._saveError = "";
		this.renderForm();

		this._formDefinedControls = [];
		var self = this;

		
		/*
			TODO : 
			Use this._context.mode.contextInfo for hosting entity (entityId, entityTypeName, entityRecordName)
			and Xrm.Page.getAttribute("primarycontactid").addOnchange ... for detecting changes in lookup Field
		*/


		if(this._formId){
			this.controlsStrings.formNotFoundMessage = this.controlsStrings.formNotFoundMessage.replace("#formid", this._formId);
			this._context.webAPI.retrieveRecord("systemform", this._formId, "?$select=formjson,formxml").then(
				entity => {
					debugger;
					self._formEntity = entity.objecttypecode;
					self._formDefinedControls = this.getFieldsFromXml(entity.formxml);

					self._lookupMetadatas = [];
					
					self._context.utils.getEntityMetadata(self._formEntity, self._formDefinedControls.map(a => a.attribute)).then(meta => {
						self._entityMetadata = meta;
						self._lookupMetadatas.push({entity: self._formEntity, done: true, metadata: meta});
						

						//@ts-ignore
						meta.Attributes.getAll().forEach(attr => {
							var am = attr.attributeDescriptor;
							var field = self._formDefinedControls.filter(a => a.attribute == am.LogicalName);
							if(field.length > 0){
								field[0].required = am.RequiredLevel == 2;
								field[0].recommended = am.RequiredLevel == 3;
								field[0].validForUpdate = am.IsValidForUpdate;
							}
						});
						
						//@ts-ignore
						var lookups = meta.Attributes.getAll().filter(a => self.lookupTypes.indexOf(a.attributeDescriptor.Type) > -1);
						//@ts-ignore
						lookups.forEach(l => {
							l.attributeDescriptor.Targets.forEach((ent: string) => {
								if(self._lookupMetadatas.filter(a => a.entity == ent).length == 0){
									self._lookupMetadatas.push({entity: ent, done: false, metadata: null});
								}
							})
						});

						self.getNeededMetadata();
					});
				},
				error => {
					this.setError({ id: this.errorTypes.formNotFoundError, errorText: this.controlsStrings.formNotFoundMessage, errorType: MessageBarType.blocked});
					debugger;
				});
		}
	}
	private getFieldsFromJson(formjson: string){
		var _formDefinedControls:Array<any> = [];
		var form = JSON.parse(formjson);
		//@ts-ignore
		form.Tabs.$values.forEach(tab => {
			//@ts-ignore
			tab.Columns.$values.forEach(col => {
				//@ts-ignore
				col.Sections.$values.forEach(sect => {
					//@ts-ignore
					sect.Rows.$values.forEach(row => {
						//@ts-ignore
						row.Cells.$values.forEach(cell => {
							if(cell.Control && cell.Control.DataFieldName && cell.Control.Visible){
								_formDefinedControls.push(
									{
										attribute: cell.Control.DataFieldName, 
										disabled: cell.Control.Disabled, 
										params: cell.Control.Parameters
									});
							}
						});
					});
				});
			});
		});
		return _formDefinedControls;
	}
	private getFieldsFromXml(formxml: string){
		var _formDefinedControls:Array<any> = [];
		try{
			var x2js = new X2JS();
			var form = x2js.xml2js(formxml);
	
			var self = this;
	
			//@ts-ignore
			self.getAsArray(form.form.tabs.tab).forEach(tab => {
				//@ts-ignore
				self.getAsArray(tab.columns.column).forEach(col => {
					//@ts-ignore
					self.getAsArray(col.sections.section).forEach(sect => {
						//@ts-ignore
						self.getAsArray(sect.rows.row).forEach(row => {
							//@ts-ignore
							self.getAsArray(row.cell).forEach(cell => {
								if(cell.control && cell.control._datafieldname){
									var visible = cell.control._visible;
									if(visible != "false"){
										var label = "";
										if(cell.labels && cell.labels.label){
											//@ts-ignore
											var currLabel = self.getAsArray(cell.labels.label).filter(a => a._languagecode == self._context.userSettings.languageId.toString())
											if(currLabel.length > 0){
												label = currLabel[0]._description;
											}
										}
										_formDefinedControls.push(
											{
												attribute: cell.control._datafieldname, 
												disabled: cell.control._disabled == "true", 
												label: label
											});
									}
								}
							});
						});
					});
				});
			});
		}catch(exc){
			debugger;
		}
		
		return _formDefinedControls;
	}

	private getNeededMetadata(){
		var self = this;
		var metaToRetrieve = self._lookupMetadatas.filter(a => !a.done);
		if(metaToRetrieve.length > 0){
			var l = metaToRetrieve[0];
			self._context.utils.getEntityMetadata(l.entity, []).then(
				m => {
					l.done = true;
					l.metadata = m;
					self.getNeededMetadata();
				},
				error => {
					this.setError({ id: this.errorTypes.runtimeError, errorText: error.message, errorType: MessageBarType.error});
				})
		}else{
			self.createForm();
		}
	}
	private createForm(){
		var self = this;
		self._formControls = [];
		self._formDefinedControls.forEach(ctrl => {
			if(self._formControls.filter(a => a.id == ctrl.attribute).length == 0){
				var attrMeta = self._entityMetadata.Attributes.get(ctrl.attribute);
				if(attrMeta){
					var type = attrMeta.attributeDescriptor.Type;
					if(self.lookupTypes.indexOf(type) > -1)
						type = "lookup";
	
					var field:IFormControl = {
						id: attrMeta.LogicalName,
						label: ctrl.label ? ctrl.label : attrMeta.DisplayName,
						type: type,
						format: attrMeta.attributeDescriptor?.Format??"",
						required: ctrl.required,
						recommended: ctrl.recommended,
						disabled: ctrl.disabled || !ctrl.validForUpdate,
						value: undefined,
						randomid: this.getRandom(),
						dateFormat: self.controlsStrings.dateFormat,
						strings: self.controlsStrings,
						options: attrMeta.attributeDescriptor.OptionSet ? attrMeta.attributeDescriptor.OptionSet.map((o: { Value: any; Label: any; }) => {return {value: o.Value, label: o.Label}}) : [],
						inputChanged: this.inputValueChanged.bind(this), 
						inputError: this.inputErrorChanged.bind(this), 
						inputAutocomplete: this.inputAutocomplete.bind(this),
					}
					self._formControls.push(field);
				}
			}
		});

		self._initialization = false;
		self.updateLayout();
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		var newValue = this.getLookup();

		if(this._fieldValue != newValue){
			this._fieldValue = newValue;
			
			if(newValue){
				this.updateLayout();
			}else{
				this.setError({ id: this.errorTypes.lookupEmptyError, errorText: this.controlsStrings.lookupEmptyMessage, errorType: MessageBarType.info});
				this._saveError = "";
				this.updateLayout();
			}
		}

		// if(newValue){
		// 	if(this._fieldValue != newValue){
		// 		this._fieldValue = newValue;
		// 		this.updateLayout();
		// 	}
		// }else{
		// 	this._fieldValue = null;
		// 	if()
		// 	this._formErrors.push({ id: this.errorTypes.lookupEmptyError, errorText: this.controlsStrings.formNotFoundMessage, errorType: MessageBarType.info});
		// 	this._saveError = "";
		// 	this.renderForm();
		// }
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		ReactDOM.unmountComponentAtNode(this._container);
	}

	private getResourceString(id:string, def: string){
		try{
			var str = this._context.resources.getString(id);
			if(str != id) return str;
		}catch(exc){
			debugger;
		}
		return def;
	}
	private removeError(id: string) {
		this._formErrors = this._formErrors.filter(a => a.id != id);
	}
	private getAsArray(prop:any){
		if(!Array.isArray(prop)){
			prop = [prop];
		}
		return prop;
	}
	private setError(error: IFormError){
		if(this._formErrors.filter(a => a.id == error.id).length == 0){
			this._formErrors.push(error);
		}
	}
	private getLookupDisplayName():string{
		try{
			//@ts-ignore
			return this._context.parameters.lookupAttribute.attributes.DisplayName;
		}catch(exc){
			debugger;
			return "";
		}
	}
	private getLookup(): ILookup|null{
		if(this._context.parameters.lookupAttribute.raw){
			var array = this._context.parameters.lookupAttribute.raw as unknown as Array<ILookup>;
			if(array.length > 0){
				return array[0];
			}
		}
		return null;
	}
	
	private updateLayout(){
		var self = this;
		if(this._initialization || this._formControls.length == 0) return;

		var error = false;

		if(this._fieldValue == null){
			error = true;
			this.setError({ id: this.errorTypes.lookupEmptyError, errorText: this.controlsStrings.lookupEmptyMessage, errorType: MessageBarType.info});
		}else{
			this.removeError(this.errorTypes.lookupEmptyError);

			if(this._fieldValue.entityType != this._formEntity){
				error = true;
				this.setError({ id: this.errorTypes.formEntityMismatch, errorText: this.controlsStrings.formEntityMismatchMessage.replace("#lookupentity", this._fieldValue.entityType).replace("#formentity", this._formEntity), errorType: MessageBarType.error});
			}else{
				this.removeError(this.errorTypes.formEntityMismatch);
			}
		}

		if(error){
			this._loader.loading = false;
			this.renderForm();			
			return;
		}

		this._loader.loading = true;
		this._loader.text = this.controlsStrings.loadingText;
		this._saveError = "";
		this.renderForm();

		var sel = this._formControls.
		map(a => {
			if(self.lookupTypes.indexOf(a.type) > -1){
				return "_" + a.id + "_value";
			}else{
				return a.id
			}
		}).join(",");
		if(self._fieldValue){
			self.getCurrency(function(){
				self.removeError(self.errorTypes.runtimeError);
				self._context.webAPI.retrieveRecord(self._fieldValue?.entityType??"", self._fieldValue?.id??"", "?$select=" + sel).then(
					entity => {
						self._lookupsImages = [];
	
						self._updateObject = [];
						self._formControls.forEach(field => {
							if(self.lookupTypes.indexOf(field.type) > -1){
								var lookupValue = self.getLookupValue(entity, field.id, null);
								self._dataObject[field.id] = lookupValue;
								if(lookupValue){
									self._lookupsImages.push({
										attribute: field.id,
										lookup: lookupValue,
										done: false
									});
								}else{
									field.value = self.lookupToPersonaArray(lookupValue);
								}
							}
							else{
								self._dataObject[field.id] = self.getAttributeValue(entity, field.id, undefined);
								field.value = self.getAttributeValue(entity, field.id, undefined);
							}
						});
	
						self.getLookupsImages();
					},
					error => {
						self.setError({ id: self.errorTypes.runtimeError, errorText: error.message, errorType: MessageBarType.error});
						debugger;
					}
				)
			});
			
		}else{
			this._loader.loading = false;
			this.renderForm();
		}
	}
	private getLookupsImages(){
		var self = this;
		var next = self._lookupsImages.filter(a => a.done == false);
		if(next.length > 0){
			var req = next[0];
			var meta = self._lookupMetadatas.filter(a => a.metadata?.LogicalName == req.lookup?.entityType);
			if(meta.length > 0){
				var field = this._formControls.filter(a => a.id == req.attribute);
				if(field.length > 0){
					if(meta[0].metadata?.PrimaryImageAttribute){
						self._context.webAPI.retrieveRecord(req.lookup?.entityType??"", req.lookup?.id??"", "?$select=" + meta[0].metadata?.PrimaryImageAttribute).then(
							entity => {
								req.done = true;
								var img = self.getAttributeValue(entity, meta[0].metadata?.PrimaryImageAttribute, null);
								if(img){
									if(req.lookup){ 
										req.lookup.entityImg = "data:image/png;base64," + img;
										field[0].value = self.lookupToPersonaArray(req.lookup); 
									}
								}else{
									field[0].value = self.lookupToPersonaArray(req.lookup); 
								}
								self.getLookupsImages();
							},
							error => {
								debugger;
								req.done = true;
								self.getLookupsImages();
							}
						);
					}else{
						req.done = true;
						self.getLookupsImages();
					}	
					return;
				}
			}
			req.done = true;
			self.getLookupsImages();
		}else{
			this._loader.loading = false;
			self.renderForm();
		}
	}
	private getCurrency(callback: Function){
		var self = this;
		try{
			self._context.webAPI.retrieveRecord(self._fieldValue?.entityType??"", self._fieldValue?.id??"", "?$expand=transactioncurrencyid($select=currencysymbol)").then(
				result => {
					if (result.hasOwnProperty("transactioncurrencyid")) {
						if(result["transactioncurrencyid"]){
							var symbol = result["transactioncurrencyid"]["currencysymbol"];
							self.controlsStrings.currency = symbol;
						}
					}
					callback();
				},
				error => {
					callback();
				}
			)
		}catch(exc){
			callback();
		}
	}
	
	private inputValueChanged(id: string, type: string, format: string, newValue: any) 
	{
		var self = this;

		var updDatas = this._updateObject.filter(a => a.attribute == id);
		var updData = undefined; 
		if(updDatas.length > 0){
			updData = updDatas[0];	
			updData.value = newValue;
		}
		else{
			updData = { attribute: id, type: type, format: format, value: newValue };
			this._updateObject.push(updData);
		}

		if(this._dataObject.hasOwnProperty(id)){
			if(!this.isDataObjectDifferent(id)){
				this._updateObject = this._updateObject.filter(a => a.attribute != id);
			}
		}

		this._saveError = "";
		this.renderForm();
	}
	private isDataObjectDifferent(id: string): boolean{
		var update = this._updateObject.filter(a => a.attribute == id);
		if(update.length == 0){
			return !!this._dataObject[id];
		}else{
			if(this._dataObject[id]){
				if(this.lookupTypes.indexOf(update[0].type) > -1){
					var currLookup = this._dataObject[id] as ILookup;
					var newLookup = update[0].value as Array<ILookup>;
					return newLookup.length == 0 || currLookup.id != newLookup[0].id;
				}else{
					return update[0].value != this._dataObject[id];
				}
			}else{
				return !!update[0].value;
			}
		}
	}
	private getValueForUpdate(type: string, format: string, value: any):any{
		switch (type){
			case "string":
			case "memo":
				return value;
			case "money":
			case "decimal":
			case "float":
				var n = parseFloat(value);
				return isNaN(n) ? null : n;
			case "integer":
				var n = parseInt(value);
				return isNaN(n) ? null : n;
			case "boolean":
				var n = parseInt(value);
				return n == 1;
			case "lookup":
				var lookup = value as Array<ILookup>;
				if(lookup && lookup.length > 0){
					var meta = this._lookupMetadatas.filter(a => a.entity == lookup[0].entityType)[0]; 
					
					return "/" + meta.metadata?.EntitySetName + "(" + lookup[0].id + ")";
				}else{
					return null;
				}
				
			case "picklist":
			case "state":
			case "status":
				var n = parseInt(value);
				return isNaN(n) ? null : n;
			case "datetime":
				var d = moment(value);
				if(d.isValid()){
					if(format == "date"){
						return d.format("YYYY-MM-DD");
					}					
					if(format == "datetime"){
						return d.utc().format("YYYY-MM-DDTHH:mm:ss.000Z");
					}
				}
				else
					return null;
		}

		return null;
	}
	private inputErrorChanged(id: string, error: string){
		if(error){
			this._dataErrors[id] = error;
		}else{
			delete this._dataErrors[id];
		}
	}
	private inputAutocomplete(attribute: string, pattern: string, max: number|undefined){
		var self = this;
		return new Promise<IPersonaObj[]>((resolve, reject) => {
			var field = self._context
			var attrMeta = self._entityMetadata.Attributes.get(attribute);
			if(attrMeta){
				var targets = attrMeta.attributeDescriptor.Targets as Array<string>;
				if(targets){
					var requestGroup:ILookupRequestGroup = {
						max: max,
						pattern: pattern,
						resolve: resolve,
						requests: targets.map(a => { return {
							done: false,
							entity: a,
							results: []
						}})
					};

					self.searchLookups(requestGroup);
				}else{
					resolve([]);
				}
			}else{
				resolve([]);
			}
		})
	}


	private getAttributeValue(entity: ComponentFramework.WebApi.Entity, attribute: string, def: any):any{
		if(entity.hasOwnProperty("_" + attribute + "_value")) return entity["_" + attribute + "_value"];
		if(entity.hasOwnProperty(attribute)) return entity[attribute];
		return def;
	}
	private getAttributeLabel(entity: ComponentFramework.WebApi.Entity, attribute: string, def: any):any{
		if(entity.hasOwnProperty("_" + attribute + "_value")) 
		{
			if(entity.hasOwnProperty("_" + attribute + "_value" + "@OData.Community.Display.V1.FormattedValue")){
				return entity["_" + attribute + "_value" + "@OData.Community.Display.V1.FormattedValue"];
			}
			return entity[attribute];
		}
		if(entity.hasOwnProperty(attribute)) 
		{
			if(entity.hasOwnProperty(attribute + "@OData.Community.Display.V1.FormattedValue")){
				return entity[attribute + + "@OData.Community.Display.V1.FormattedValue"];
			}
			return entity[attribute];
		}
		return def;
	}
	private getLookupValue(entity: ComponentFramework.WebApi.Entity, attribute: string, def: any):ILookup|null{
		if(entity.hasOwnProperty("_" + attribute + "_value")) 
		{
			if(this.getAttributeValue(entity, attribute, undefined)){
				return {
					id: entity["_" + attribute + "_value"],
					name: entity["_" + attribute + "_value" + "@OData.Community.Display.V1.FormattedValue"],
					entityType: entity["_" + attribute + "_value" + "@Microsoft.Dynamics.CRM.lookuplogicalname"],
					entityImg: undefined,
				}
			}
		}
		return null;
	}
	private lookupToPersonaArray(lookup: ILookup|null):IPersonaObj[]|undefined{
		if(lookup){
			//crmSymbolFont entity-symbol TransactionCurrency
			return [{
				lookup: lookup,
				text: lookup.name, 
				initialsColor: "#F44336",
				className: "customPersona",
				imageUrl: lookup.entityImg,
				imageInitials: ""
			}];
		}
		return undefined;
	}
	private searchLookups(reqGroup: ILookupRequestGroup){
		var self = this;
		var remaining = reqGroup.requests.filter(a => !a.done);
		if(remaining.length > 0){
			var curr = remaining[0];
			var meta = self._lookupMetadatas.filter(a => a.entity == curr.entity)[0].metadata;
			var sel = meta?.PrimaryIdAttribute + "," + meta?.PrimaryNameAttribute;
			if(meta?.PrimaryImageAttribute){
				sel += "," + meta?.PrimaryImageAttribute;
			}

			var reqMax = self._defaultMaxLookupRequest;
			if(reqGroup.max) reqMax = reqGroup.max;

			self._context.webAPI.retrieveMultipleRecords(curr.entity, "?$select=" + sel + "&$filter=contains(" + meta?.PrimaryNameAttribute + ",'" + reqGroup.pattern + "')", reqMax).
			then(
				results =>{
					var data = results.entities.map(a => {
						var img = a[meta?.PrimaryImageAttribute];
						return {
							lookup: {
								id: a[meta?.PrimaryIdAttribute],
								name: a[meta?.PrimaryNameAttribute], 
								entityType: curr.entity,
								entityImg: img ? ("data:image/png;base64," + img) : undefined,
							},
							text: a[meta?.PrimaryNameAttribute], 
							initialsColor: "#F44336",
							className: "customPersona",
							imageUrl: img ? ("data:image/png;base64," + img) : undefined
						};
					});
					curr.results = data;
					curr.done = true;
					self.searchLookups(reqGroup);
				},
				error => {
					debugger;
					curr.results = [];
					curr.done = true;
					self.searchLookups(reqGroup);
				});
		}else{
			var finalData:Array<IPersonaObj> = [];
			reqGroup.requests.forEach(r => {
				r.results.forEach(l => {
					finalData.push(l);
				});
			});
			finalData.sort(function(a,b){
				var av = a ? a.text??"" : "";
				var bv = a ? b.text??"" : "";
				if(av > bv) return 1;
				else return -1
			});

			reqGroup.resolve(finalData);
		}
	}
	private getRandom():string {
		var result           = '';
		var characters       = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
		var charactersLength = characters.length;
		for ( var i = 0; i < 10; i++ ) {
		   result += characters.charAt(Math.floor(Math.random() * charactersLength));
		}
		return result;
	}
	private renderForm(){
		try{
			ReactDOM.render(React.createElement(Form, 
				{
					title: this._formTitle, 
					formErrors: this._formErrors,
					saveError: this._saveError, 
					controls: this._formControls, 
					loader: this._loader, 
					canSave: this.canSave(), 
					save: this.saveForm.bind(this) 
				}),this._container);
		}catch(exc){
			debugger;
		}
	}
	
	private canSave():boolean{
		var canSave = true;
		try{
			this._formControls.forEach(field => {
				var upd = this._updateObject.filter(a => a.attribute == field.id);
				if(field.required){
					if(upd.length == 0){
						if(!this._dataObject[field.id]) canSave = false;
					}
					else if(!upd[0].value || (Array.isArray(upd[0].value) && upd[0].value.length == 0)) {
						canSave = false;
					}
				}
			});
			if(this._updateObject.length == 0) canSave = false;
		}catch(exc){
			debugger;
		}
		
		return canSave;
	}
	private saveForm(){
		this._loader.text = this.controlsStrings.savingText;
		this._loader.loading = true;
		this._saveError = "";
		this.renderForm();

		var obj:any = {};
		var self = this;
		this._updateObject.forEach(upd => {
			if(self.lookupTypes.indexOf(upd.type) > -1){
				var lookup = upd.value as Array<ILookup>;
				if(lookup.length == 0){
					obj["_" + upd.attribute + "_value"] = null;
				}else{
					var suffix = "";
					var attMeta = self._entityMetadata.Attributes.get(upd.attribute).attributeDescriptor;
					if(attMeta.Type == "customer"){
						suffix = "_" + lookup[0].entityType;
					}
					obj[upd.attribute + suffix+ "@odata.bind"] = self.getValueForUpdate(upd.type, upd.format, lookup);
				}
			}else{
				obj[upd.attribute] = self.getValueForUpdate(upd.type, upd.format, upd.value);
			}
		});

		var self = this;
		this._context.webAPI.updateRecord(this._fieldValue?.entityType??"", this._fieldValue?.id??"", obj).then(
			ref => {
				self.updateLayout();
			},
			error => {
				debugger;
				self._updateObject = [];
				self._loader.loading = false;
				self._saveError = error.message;
				self.renderForm();
			}
		)

	}
}