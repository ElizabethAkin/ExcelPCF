import {IInputs, IOutputs} from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;
import * as $ from 'jquery';
import './js/htmlson.js';
import BatchPostRecords from "./utils/BatchPostRecords";

declare var Xrm: any;

export class PCFexcel implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private mainContainer: HTMLDivElement;
	private _notifyOutputChanged: () => void;
	private context: any;
	private contextObj: ComponentFramework.Context<IInputs>;
	private _entityName: string;
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
		this._notifyOutputChanged = notifyOutputChanged;
		this.contextObj = context;

		if (context.parameters.entityName != null)
		this._entityName = context.parameters.entityName.raw || "";

		this.mainContainer = document.createElement("div");
		this.mainContainer.innerHTML = `
        	<input type="file" onchange="importf(this)" style="margin-bottom:10px"/>
		`;

		var tableElement = document.createElement("div");
		tableElement.innerHTML = `
			<table id="tableWithDataFromExcel"></table>
		`;
		
		var buttonLoadFields = document.createElement("input");
		buttonLoadFields.id = "buttonLoadFieldsId";
		buttonLoadFields.type="button";
		buttonLoadFields.value="Load fields";
		buttonLoadFields.onclick = () => this.loadFields();

		this.mainContainer.appendChild(buttonLoadFields);
		this.mainContainer.appendChild(tableElement);

		var scriptXLSXPlugin = document.createElement("script");
		scriptXLSXPlugin.type = "text/javascript";
		scriptXLSXPlugin.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js";
		container.appendChild(scriptXLSXPlugin);

		var scriptElement = document.createElement("script");
		scriptElement.type = "text/javascript";
		scriptElement.innerHTML = `
		var wb;
		function importf(obj) {
			if(!obj.files) {
				return;
			}
			var f = obj.files[0];
			var reader = new FileReader();
			reader.onload = function(e) {
				var data = e.target.result;
				wb = XLSX.read(data, {
					type: 'binary'
				});
				var data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
				var tableWithData = $('#tableWithDataFromExcel').htmlson({data: data});
				$('#tableWithDataFromExcel').addClass("paleBlueRows");
				$("#buttonLoadFieldsId").css("visibility","visible");
			};
			reader.readAsBinaryString(f);
		}`;

		var createButton = document.createElement("input");
		createButton.id = "buttonCreateRecocrdsId";
		createButton.type="button";
		createButton.value="Create recods";
		createButton.style.visibility="hidden";
		createButton.onclick = () => this.createRecords();

		this.mainContainer.appendChild(createButton);

		container.appendChild(scriptElement);
		container.appendChild(this.mainContainer);
	}

	public loadFields() {
		$("#buttonLoadFieldsId").css("visibility","hidden");
		$("#tableWithDataFromExcel thead tr").append("<th>Check to save</th>");
		$("#tableWithDataFromExcel tbody tr").each(function(){
			$(this).find("td").each(function(){
				var curval = $(this).html().toString();
                $(this).html("<input type='text' value='"+curval+"'>");
			});
			$(this).append('<input type="checkbox" checked>');
		});
	   $("#buttonCreateRecocrdsId").css("visibility","visible");
	}

	public createRecords(){
		$("#tableWithDataFromExcel tbody tr").each(function(){
			var values = "";
			$(this).find("td").each(function(){
				values+=$(this).find("input[type='text']").val()+",";
                console.log(values.slice(0, -1));
			});
            $(this).find("input[type='checkbox']").val(values);
		});

		var url: string = (<any>Xrm).Utility.getGlobalContext().getClientUrl();
		var attributes = new Array();
		$("#tableWithDataFromExcel thead tr th").each(function(){
			attributes.push($(this).html());
		});
		attributes.pop();

		var batchRequest = new BatchPostRecords(url + "/api/data/v9.1/");

		$("#tableWithDataFromExcel tbody tr").each(function(){
			var jsonDataString = "{"
			$(this).find("input[type='checkbox']:checked").each(function(){
				var value = $(this).val()?.toString() || "";
				const valuesArray = value.split(",");
				for(var i = 0; i < attributes.length; i++)
				{
					jsonDataString += "'"+ attributes[i] +"':'" + valuesArray[i] + "',";
				}
				jsonDataString = jsonDataString.slice(0, -1);
			});
			jsonDataString += "}"
			if(jsonDataString!="{}")
			{
				batchRequest.addRequestItem(jsonDataString);
				
			}
		});

		batchRequest.sendRequest();
	}

	
	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		// Add code to update control view
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
		// Add code to cleanup control if necessary
	}

}