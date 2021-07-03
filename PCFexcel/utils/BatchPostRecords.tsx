export default class BatchPostRecords{
	    public apiUrl: string;
		public uniqueId: string;
		public batchItemHeader: string;
		public content: Array<string>;
	constructor(apiUrl:string)
	{
		this.apiUrl = apiUrl;
		this.uniqueId = "batch_" + (new Date().getTime());
		this.batchItemHeader = "--" + 
        this.uniqueId + 
        "\nContent-Type: application/http\nContent-Transfer-Encoding:binary";
		this.content = [];
	}
	public addRequestItem(entity: Object): void{
		this.content.push(this.batchItemHeader);
		this.content.push("");
		this.content.push("POST " + this.apiUrl + "contacts" + " HTTP/1.1");
		this.content.push("Content-Type: application/json;type=entry");
		this.content.push("");
		this.content.push(JSON.stringify(entity));
	}
	public sendRequest(): void{
		this.content.push("");
		this.content.push("--" + this.uniqueId + "--");
		this.content.push(" ");
	
		var xhr = new XMLHttpRequest();
		xhr.open("POST", encodeURI(this.apiUrl + "$batch"));
		xhr.setRequestHeader("Content-Type", "multipart/mixed;boundary=" + 
			this.uniqueId);
		xhr.setRequestHeader("Accept", "application/json");
		xhr.setRequestHeader("OData-MaxVersion", "4.0");
		xhr.setRequestHeader("OData-Version", "4.0");
		xhr.addEventListener("load", 
			function() { 
				console.log("Batch request response code: " + xhr.status); 
			});
		
		xhr.send(this.content.join("\n"));
	}
}