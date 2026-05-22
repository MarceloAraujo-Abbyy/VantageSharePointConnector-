
Context.LogMessage("Custom Activity - Export SP Start");

// SharePoint client_credentials 
var sp_client_id = Context.GetSecret("sp-client-id"); 				
var sp_client_secret = Context.GetSecret("sp-client-secret"); 		
var sp_tenant_id =  Context.GetSecret("sp-tenant-id"); 											
var sp_site_id = Context.GetSecret("sp-site-id"); 					
var sp_drive_id = Context.GetSecret("sp-drive-id"); 					
var sp_list_id = Context.GetSecret("sp-list-id"); 					

// get SharePoint access token 
var access_token = getToken();

for (var doc of Context.Transaction.Documents) {
    // name file in Sharepoint 
    var filename = doc.DocumentId+".pdf";
	// upload to SharePoint 
	var ret_upload = uploadFile(doc, filename);
	// filter to get id of file 
	var item_id = filterSharePointList(ret_upload.filename);
	// create mapping info SharePoint x Vantage 
	var data = createUpdateItemData(doc);
	// update SharePoint List 
	updateItem(item_id,data);

}

function createUpdateItemData(doc) {

	const parInvoiceNumber = doc.getField("Invoice Number").Value;
	const parTotal = doc.getField("Total").Value;
	const parInvoiceDate = doc.getField("Invoice Date").Value;

    // {<SharePointFieldInternalName> : <VantageField.Value>, ... } 
	let dataObject = {
		
		InvoiceNumber: parInvoiceNumber,
		Total: parTotal,
		DocDate: parInvoiceDate
	}

	Context.LogMessage('update data info: ' + JSON.stringify(dataObject))
	return JSON.stringify(dataObject);
}

function getToken() {
	const request = Context.CreateHttpRequest();
	request.Url = 'https://login.microsoftonline.com/' + sp_tenant_id + '/oauth2/v2.0/token';
	request.Method = 'POST';
	const loginContent = 'client_id=' + sp_client_id + '&scope=https://graph.microsoft.com/.default&client_secret=' + sp_client_secret + '&grant_type=client_credentials';
	request.SetStringContent(loginContent, ContentEncoding.UTF8, "application/x-www-form-urlencoded");
	request.Send();

	let jsonResponse = 'ERROR'; 
	if (request.ResponseText)
		jsonResponse = JSON.parse(request.ResponseText);
	if (request.Status == 200) {
		Context.LogMessage(`Auth token requested successfully`);
		return jsonResponse['access_token'];
	}
	throw `SharePoint token request error: ${jsonResponse['error_description']}`;
}

function uploadFile(doc, filename) {

	var exportResult = doc.Exports.GetByFormat(ExportFormat.Pdf);

	const request = Context.CreateHttpRequest();
	request.Url = 'https://graph.microsoft.com/v1.0/sites/' + sp_site_id + "/drives/" + sp_drive_id + '/root:/' + filename +  ':/content';
	request.Method = 'PUT';
	request.AuthToken = access_token;
	request.AuthScheme = 'Bearer';
	request.SetFileContent(exportResult);
	request.Send();
	if (request.Status == 201 || request.Status == 200) {
		const item_id = JSON.parse(request.ResponseText).id;
		Context.LogMessage(`file upload successfully. filename: ` + filename + ' item_id: ' + item_id);
		return {"filename": filename, "item_id": item_id};
	} 
	else 
		throw `SharePoint upload error: ${request.ResponseText}`;

}

function filterSharePointList(filename) {
	
	const request = Context.CreateHttpRequest();
	request.Url = "https://graph.microsoft.com/v1.0/sites/" + sp_site_id + "/lists/" + sp_list_id + "/items?$filter=fields/FileLeafRef eq '"+filename+"'&$expand=fields";
	request.Method = 'GET';
	request.AuthToken = access_token;
	request.AuthScheme = 'Bearer';
	request.SetHeader("Prefer","HonorNonIndexedQueriesWarningMayFailRandomly")
	request.Send();
	if (request.Status == 201 || request.Status == 200) {
		var jsonResponse = JSON.parse(request.ResponseText);
		Context.LogMessage('SharePoint filered successfully. item_id: ' + jsonResponse['value'][0]['id']);
		return jsonResponse['value'][0]['id'];
	} else 
		throw `SharePoint filter error`;

}



function updateItem(item_id, data) {
	const request = Context.CreateHttpRequest();
	request.Url = 'https://graph.microsoft.com/v1.0/sites/'+sp_site_id+'/lists/'+sp_list_id+'/items/'+item_id+'/fields';
	request.Method = 'PATCH';
	request.AuthToken = access_token;
	request.AuthScheme = 'Bearer';
	request.SetStringContent(data);
	request.Send();
	if (request.Status == 201 || request.Status == 200) {
		Context.LogMessage(`SharePoint list updated successfully`);
	} else 
		throw `SharePoint update item error`;
}
