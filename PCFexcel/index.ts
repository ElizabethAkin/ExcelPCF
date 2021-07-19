import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as $ from 'jquery';
import './js/htmlson.js';
import BatchPostRecords from "./utils/BatchPostRecords";

declare var Xrm: any;

interface IDropdownOption{
	key:string, text:string, type:string, schemaName:string
}

export class PCFexcel implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private mainContainer: HTMLDivElement;
	private _notifyOutputChanged: () => void;
	private context: any;
	private contextObj: ComponentFramework.Context<IInputs>;
	private _entitySchemaName: string;
	private _entityCollectionSchemaName: string;
	private _url: string;
	private options: IDropdownOption[]=[];
	private areAttributesMapped: boolean;
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

		this._url = (<any>Xrm).Utility.getGlobalContext().getClientUrl();

		if (context.parameters.entitySchemaName != null)
		this._entitySchemaName = context.parameters.entitySchemaName.raw || "";
		this._entityCollectionSchemaName = context.parameters.entityCollectionSchemaName.raw || "";

		this.mainContainer = document.createElement("div");
		this.mainContainer.innerHTML = `
        	<input type="file" onchange="importf(this)" style="margin-bottom:10px"/>
		`;

		var tableElement = document.createElement("div");
		tableElement.classList.add("tableDiv");
		tableElement.innerHTML = `
			<table id="tableWithDataFromExcel"></table>
		`;
		
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
				var data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval:""});
				var tableWithData = $('#tableWithDataFromExcel').htmlson({data: data});
				$("#tableWithDataFromExcel thead tr").append("<th id='CheckboxesHeaderId' style='position: sticky; right: 0;top:0;background:#1C6EA4;border-bottom: 2px solid #444444; z-index: 3'>Select to create</th>");
				
			};
			reader.readAsBinaryString(f);

			var checkExist = setInterval(function() {
				if ($("#tableWithDataFromExcel tbody tr td").length) {
					
					$("#tableWithDataFromExcel tbody tr").each(function(){
						$(this).find("td").each(function(){
							var curval = $(this).html().toString();
							$(this).html("<input type='text' value='"+curval+"'>");
						});
						$(this).append('<td style="right: 0; background: #b1d4f2; position: sticky; z-index: 0"><input type="checkbox" checked ></td>');
					});
					$('#CheckboxesHeaderId').click(function()
					{
						$("#tableWithDataFromExcel tbody tr td:last-child input[type='checkbox']").each(function(){
							$(this).click();
						})
					});
					$('#tableWithDataFromExcel').addClass("paleBlueRows");
					$("#messageResponse").css("visibility","hidden");
				   clearInterval(checkExist);
				}
			 }, 100);

			
	 		$("#buttonCreateRecocrdsId").css("visibility","visible");
	 		$("#loadDropdownsButtonId").css("visibility","visible");
	 		$("#changeAttributesButtonId").css("visibility","visible");
		}`;

		var loadDropdownsButton = document.createElement("input");
		loadDropdownsButton.id = "loadDropdownsButtonId";
		loadDropdownsButton.type="button";
		loadDropdownsButton.value="Load Dropdowns";
		loadDropdownsButton.style.visibility="hidden";
		loadDropdownsButton.onclick = () => this.loadDropdowns();

		var changeAttributesButton = document.createElement("input");
		changeAttributesButton.id = "changeAttributesButtonId";
		changeAttributesButton.type="button";
		changeAttributesButton.value="Map attributes";
		changeAttributesButton.style.visibility="hidden";
		//changeAttributesButton.style.visibility="hidden";
		changeAttributesButton.onclick = () => this.changeAttributes();
		var createButton = document.createElement("input");
		createButton.id = "buttonCreateRecocrdsId";
		createButton.type="button";
		createButton.value="Create records";
		createButton.style.visibility="hidden";
		createButton.onclick = () => this.createRecords();

		var messageResponse = document.createElement("div");
		messageResponse.id = "messageResponse";
		messageResponse.style.visibility="hidden";

		this.mainContainer.appendChild(loadDropdownsButton);
		this.mainContainer.appendChild(changeAttributesButton);
		this.mainContainer.appendChild(messageResponse);
		this.mainContainer.appendChild(createButton);

		container.appendChild(scriptElement);
		container.appendChild(this.mainContainer);
	}
	
	public changeAttributes(){
		this.areAttributesMapped = true;
		$("#changeAttributesButtonId").css("visibility","hidden");
		this.populateAttributeMapping(this._entitySchemaName).then(
			function(){
				var table = document.getElementById("tableWithDataFromExcel") as HTMLTableElement;
				
    			for( var j = 0; j < table.rows[0].cells.length-1; j++)
    			{
    			    var excelColumn = table.rows[0].cells[j];
    			    var excelColumnName = $(excelColumn).text();
				
    			    var idx = table.rows[1].cells[j].cellIndex;
				
    			    var head = $(table).find('tr:last-child th').eq(idx);
    			    var select = $(head).find('select');
    			    console.log(select);
				
    			    $(select).find("option").filter(function() {
    			        let str = $(this).text();
    			        return str.includes(excelColumnName) 
    			    }).prop('selected', true);
    			}
			}
		);
		
	}

	public loadDropdowns(){
		var attributes = "";
		$("#tableWithDataFromExcel thead tr th:not(:last-child) ").each(function(){
			attributes += $(this).html()+",";
		});
		attributes = attributes.substring(0, attributes.​length - 1);
		debugger;
		this.contextObj.webAPI.retrieveMultipleRecords(this._entitySchemaName, "?$select=" + attributes + "&$top=1").then(
			function successCallback(value: any){

			}, 
			function errorCallback(error: any){
				$("#messageResponse").css("visibility","visible");
				$("#messageResponse").html("<p>Import failed: "+error.message+"</p>");
				$("#messageResponse").css("color","red");
			});
	}

	public createRecords(){
		var entityCollectionSchemaName = this._entityCollectionSchemaName;

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
		if(this.areAttributesMapped)
		{
			$("#tableWithDataFromExcel thead tr:last-child th select").each(function(){
				var opt = $(this);
				var type = $(opt).find(":selected").attr("data-attributetype");
				var schemaName = $(opt).find(":selected").attr("data-schemaName");

				if(type == "#Microsoft.Dynamics.CRM.LookupAttributeMetadata")
				{
					attributes.push(schemaName+"@odata.bind");
				}
				else
				{
					attributes.push($(opt).val());
				}

			});
		}
		else{
			$("#tableWithDataFromExcel thead tr th").each(function(){
				debugger;
				attributes.push($(this).html());
				attributes.pop();
			});
		}

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
				batchRequest.addRequestItem(jsonDataString,entityCollectionSchemaName);
				
			}
		});

		var result: XMLHttpRequest = batchRequest.sendRequest();
		if(result.status == 200)
		{
			$("#tableWithDataFromExcel tbody tr").each(function(){
				if($(this).find("input[type='checkbox']:checked").length)
				{
					$(this).remove();
				};
				
				$("#messageResponse").css("visibility","visible");
				$("#messageResponse").html("Import was successful!");
				$("#messageResponse").css("color","green");
			});
			if(!$("#tableWithDataFromExcel tbody tr").length)
			{
				$("#tableWithDataFromExcel").empty()
			}
		}
		else
		{
			$("#messageResponse").css("visibility","visible");
			$("#messageResponse").html("Import failed: "+result.responseText);
			$("#messageResponse").css("color","red");
		}
		console.log(result);
	}

	private async populateAttributeMapping(entity:string) {
		let selectOption = document.createElement("option");
		if (entity!==""){


			var a = await this.getAttributes(entity);
			var result = JSON.parse(a);
			var options: IDropdownOption[]=[];
			
			// format all the options into a usable record
			for (var i = 0; i < result.value.length; i++) {

				if (result.value[i].DisplayName !== null && result.value[i].DisplayName.UserLocalizedLabel !== null) {
					var text = result.value[i].DisplayName.UserLocalizedLabel.Label + " (" + result.value[i].LogicalName + ")";
					var option: IDropdownOption = { key: result.value[i].LogicalName, text: text, type: result.value[i]["@odata.type"], schemaName:result.value[i].SchemaName }
					options.push(option);
					
				}
			}
			options.sort((a, b) => a.text.localeCompare(b.text));
			var dropdownControl = document.createElement("select");
			dropdownControl.style.maxWidth = '150px';

			// add a top level empty option in case it's needed
			for (let i = 0; i < options.length; i++) {
				
				selectOption = document.createElement("option");
				selectOption.innerHTML = options[i].text;
				selectOption.value = options[i].key;
				selectOption.setAttribute("data-attributeType",options[i].type);
				selectOption.setAttribute("data-schemaName",options[i].schemaName);

				dropdownControl.add(selectOption);
			}
			$("#tableWithDataFromExcel thead").append("<tr id='attributesHeaderId'></tr>");
			$("#tableWithDataFromExcel thead tr th:not(:last-child) ").each(function(){
				$("#attributesHeaderId").append("<th style='max-width:150px'>"+dropdownControl.outerHTML+"</th>");
			});
			
			$("#attributesHeaderId").append("<th style='right: 0; background: #b1d4f2; position: sticky; z-index: 3'></th>");
				
			}
		
	}		


	private async getAttributes(entity: string):Promise<string> {
		var req = new XMLHttpRequest();
		var baseUrl=this._url;
		return new Promise(function (resolve, reject) {

			req.open("GET", baseUrl + "/api/data/v9.2/EntityDefinitions(LogicalName='"+entity+"')/Attributes?$select=LogicalName,SchemaName,DisplayName,AttributeType", true);
			req.onreadystatechange = function () {
				
				if (req.readyState !== 4) return;
				if (req.status >= 200 && req.status < 300) {
					try {
						var result = JSON.parse(req.responseText);
						if (parseInt(result.StatusCode) < 0) {
							reject({
								status: result.StatusCode,
								statusText: result.StatusMessage
							});
						}
						resolve(req.responseText);
					}
					catch (error) {
						throw error;
					}	
				} else {
					reject({
						status: req.status,
						statusText: req.statusText
					});
				}	
			};
			req.setRequestHeader("OData-MaxVersion", "4.0");
			req.setRequestHeader("OData-Version", "4.0");
			req.setRequestHeader("Accept", "application/json");
			req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
			req.send();
		});
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