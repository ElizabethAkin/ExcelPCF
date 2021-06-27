function BatchPostEntity(){ 
    batchApiUrl = url + 
        "/api/data/v9.1/"; 
    batchuniqueId = "batch_" + (new Date().getTime()); 
    batchItemHeader = "--" +  
        batchuniqueId +  
        "\nContent-Type: application/http\nContent-Transfer-Encoding:binary"; 
    batchContent = []; 
} 
BatchPostEntity.prototype.addRequestItem = function(entity) { 
    batchContent.push(batchItemHeader); 
    batchContent.push(""); 
    batchContent.push("POST " + batchApiUrl + "contacts" + " HTTP/1.1"); 
    batchContent.push("Content-Type: application/json;type=entry"); 
    batchContent.push(""); 
    batchContent.push(JSON.stringify(entity)); 
} 
 
BatchPostEntity.prototype.sendRequest = function() { 
    batchContent.push(""); 
    batchContent.push("--" + batchuniqueId + "--"); 
    batchContent.push(" "); 
 
    var xhr = new XMLHttpRequest(); 
    xhr.open("POST", encodeURI(batchApiUrl + "$batch")); 
    xhr.setRequestHeader("Content-Type", "multipart/mixed;boundary=" + batchuniqueId); 
    xhr.setRequestHeader("Accept", "application/json"); 
    xhr.setRequestHeader("OData-MaxVersion", "4.0"); 
    xhr.setRequestHeader("OData-Version", "4.0"); 
    xhr.addEventListener("load",  
        function() {  
            console.log("Batch request response code: " + xhr.status);  
        }); 
 
    xhr.send(this.content.join("\n")); 
}