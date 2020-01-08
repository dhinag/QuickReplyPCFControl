import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { ITextFieldProps, QuickReplyText } from "./QuickReplyText";

export class QuickReplyComponent
  implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private notifyOutputChanged: () => void;
  private _container: HTMLDivElement;
  private quickReplyTextDataList : HTMLDataListElement;
  private dlOption : HTMLOptionElement;
  private _context: ComponentFramework.Context<IInputs>;
  private controlContainer: HTMLDivElement;

  private readonly entityName: string = "dg_quickreply";
  private readonly expandedTextFieldName : string = "dg_expandedtext";
  private readonly shortenedTextFieldName : string = "dg_name";
  private readonly textAreaFieldLimit : number = 100 ;

  private props : ITextFieldProps = {
    multiline : false,
    textFieldChanged: this.textFieldChanged.bind(this),
    textFieldBlured: this.textFieldBlured.bind(this),
    textAreaFieldLimit : this.textAreaFieldLimit
  }

  private textFieldChanged(newValue: string) {
    if (this.props.inputText !== newValue) {
      this.props.inputText = newValue;
      this.notifyOutputChanged();
    }
  }

  private textFieldBlured() { 
     this._context.webAPI.retrieveMultipleRecords(this.entityName, "?$select=" + this.expandedTextFieldName +"&$filter=contains(" + this.shortenedTextFieldName + ", '" + this.props.inputText + "')")
		.then( res => {
			  let expandedText: string = res.entities[0][this.expandedTextFieldName];  
        this.props.inputText = expandedText;        
				this.notifyOutputChanged();		 	
       });       
    }

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ) {
    this._context = context;
    this._container = container; 

    if (context.parameters.QuickReplyText.raw != null && context.parameters.QuickReplyText.raw.length > this.textAreaFieldLimit)
    {
      this.props.multiline = true;
    }
    
    ReactDOM.render(
      React.createElement(QuickReplyText, this.props),
      this._container
    );

    if (context.parameters.QuickReplyText.raw == null)
    this.props.inputText = "";

    this.quickReplyTextDataList = document.createElement("datalist");		
    this.quickReplyTextDataList.id = "options";
    context.webAPI.retrieveMultipleRecords(this.entityName, "?$select=" + this.shortenedTextFieldName)
		.then( res => {     
			res.entities.forEach(key => {
				this.dlOption = document.createElement("option");
				this.dlOption.value = key[this.shortenedTextFieldName];
				this.quickReplyTextDataList.appendChild(this.dlOption);
			});				 	
  });

    this.controlContainer = document.createElement("div");		
    this.controlContainer.appendChild(this.quickReplyTextDataList);
    this._container.appendChild(this.controlContainer);
    this.notifyOutputChanged = notifyOutputChanged;
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   */
  public updateView(context: ComponentFramework.Context<IInputs>): void {
    if (context.parameters.QuickReplyText.raw != null && context.parameters.QuickReplyText.raw.length > this.textAreaFieldLimit)
    {
      this.props.multiline = true;    
    }
    else{
      this.props.multiline = false;
    }
    this.props.inputText = context.parameters.QuickReplyText.raw;
    ReactDOM.render(
      React.createElement(QuickReplyText, this.props),
      this._container
    );  
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
   */
  public getOutputs(): IOutputs {
     return {      
      QuickReplyText : this.props.inputText
      };
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    ReactDOM.unmountComponentAtNode(this._container);
  };
}
