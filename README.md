````md
# ABBYY Vantage SharePoint Export Connector

Microsoft SharePoint Online Integration using Microsoft Graph API

---

# Overview

This document describes how to configure and use a custom ABBYY Vantage export activity to automatically upload processed documents and metadata into Microsoft SharePoint using Microsoft Graph API.

The connector performs the following operations:

1. Authenticate using OAuth2 Client Credentials
2. Upload PDF files to SharePoint
3. Retrieve SharePoint Item IDs
4. Update SharePoint metadata columns using ABBYY extracted fields

---

# Features

- Export ABBYY processed documents directly to SharePoint Online
- Upload PDF files automatically
- Update SharePoint metadata columns
- Use Microsoft Graph API
- Secure authentication using Azure App Registration
- Fully customizable field mapping

---

# Architecture

```text
ABBYY Vantage
      |
      |  Custom Export Activity
      |
      v
Microsoft Graph API
      |
      +--> SharePoint Site
               |
               +--> Document Library
                        |
                        +--> PDF File
                        +--> Metadata Columns
```

---

# Prerequisites

## ABBYY Requirements

- Access to create Environment Variables / Secrets
- Access to create Skills

## Microsoft Requirements

- Microsoft 365 Tenant
- SharePoint Online
- Azure App Registration
- Microsoft Graph API permissions

---

# Azure App Registration

## Step 1 — Create App Registration

Open Azure Portal:

https://portal.azure.com

Navigate to:

```text
Azure Active Directory
→ App Registrations
→ New Registration
```

Configure:

| Field | Value |
|---|---|
| Name | ABBYY Vantage SharePoint Export |
| Supported Account Types | Single Tenant |
| Redirect URI | Empty |

Click **Register**.

---

## Step 2 — Save Application Information

Save the following values:

| Azure Field | Usage |
|---|---|
| Application (client) ID | Client ID |
| Directory (tenant) ID | Tenant ID |

---

## Step 3 — Create Client Secret

Navigate to:

```text
Certificates & Secrets
→ New Client Secret
```

Save the generated secret value.

---

# Microsoft Graph API Permissions

Navigate to:

```text
API Permissions
→ Add Permission
→ Microsoft Graph
→ Application Permissions
```

Add the following permissions:

| Permission |
|---|
| Sites.Selected |

Click:

```text
Grant Admin Consent
```

---

# SharePoint Library Preparation

Create or select a SharePoint Document Library.

Example:

```text
Invoices
```

Create metadata columns as needed.

Example:

| Display Name | Internal Name | Type |
|---|---|---|
| Invoice Number | InvoiceNumber | Single line text |
| Total | Total | Currency |
| Invoice Date | DocDate | Date |

> IMPORTANT  
> The connector uses SharePoint **Internal Names**, not display names.

---

# Postman Collection

The included Postman collection retrieves:

- Site ID
- Drive ID
- List ID
- SharePoint columns

Recommended tool:

https://www.postman.com

---

# Import Collection

Import:

```text
SharePoint API.postman_collection.json
```

---

# Configure Variables

Set:

| Variable | Description |
|---|---|
| sp_client_id | Azure Client ID |
| sp_secret_id | Azure Client Secret |
| sp_tenant_id | Azure Tenant ID |

---

# Execute Requests in Order

Run sequentially:

1. Get Token
2. Get Site ID
3. Get Drive ID
4. Get List ID
5. Get List Columns

---

# CURL Examples

---

# Get OAuth Token

```bash
curl --location --request POST \
'https://login.microsoftonline.com/<TENANT_ID>/oauth2/v2.0/token' \
--header 'Content-Type: application/x-www-form-urlencoded' \
--data-urlencode 'client_id=<CLIENT_ID>' \
--data-urlencode 'client_secret=<CLIENT_SECRET>' \
--data-urlencode 'grant_type=client_credentials' \
--data-urlencode 'scope=https://graph.microsoft.com/.default'
```

## Parameters

| Parameter | Description |
|---|---|
| TENANT_ID | Azure Tenant ID |
| CLIENT_ID | Azure Application Client ID |
| CLIENT_SECRET | Azure Client Secret |
| grant_type | Must be `client_credentials` |
| scope | Microsoft Graph API scope |

---

# Get SharePoint Site ID

```bash
curl --location --request GET \
'https://graph.microsoft.com/v1.0/sites/<SHAREPOINT_DOMAIN>:/sites/<SITE_NAME>' \
--header 'Authorization: Bearer <ACCESS_TOKEN>'
```

Example:

```bash
curl --location --request GET \
'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/Invoices' \
--header 'Authorization: Bearer eyJ0eXAiOiJKV1Qi...'
```

---

# Get Drive ID

```bash
curl --location --request GET \
'https://graph.microsoft.com/v1.0/sites/<SITE_ID>/drive' \
--header 'Authorization: Bearer <ACCESS_TOKEN>'
```

---

# Get List ID

```bash
curl --location --request GET \
'https://graph.microsoft.com/v1.0/sites/<SITE_ID>/lists?$filter=displayName eq '\''Documents'\''' \
--header 'Authorization: Bearer <ACCESS_TOKEN>'
```

---

# Get SharePoint Columns

```bash
curl --location --request GET \
'https://graph.microsoft.com/v1.0/sites/<SITE_ID>/lists/<LIST_ID>/columns' \
--header 'Authorization: Bearer <ACCESS_TOKEN>'
```

---


# ABBYY Vantage Configuration

## Create Secrets

Navigate to:

```text
Administration
→ Secrets
```

Create:

| Secret Name | Description |
|---|---|
| sp-client-id | Azure Client ID |
| sp-client-secret | Azure Client Secret |
| sp-tenant-id | Azure Tenant ID |
| sp-site-id | SharePoint Site ID |
| sp-drive-id | SharePoint Drive ID |
| sp-list-id | SharePoint List ID |

---

# Create Custom Export Activity

Navigate to:

```text
Skills
→ Activities
→ New Custom Activity or use Output Activity
```

Type:

```text
Export Activity
```

Paste the JavaScript connector code.

Make sure select PDF format as output.
---

# Connector Script

## Current Mapping

```javascript
function createUpdateItemData(doc) {

	const parInvoiceNumber = doc.getField("Invoice Number").Value;
	const parTotal = doc.getField("Total").Value;
	const parInvoiceDate = doc.getField("Invoice Date").Value;

	let dataObject = {

		InvoiceNumber: parInvoiceNumber,
		Total: parTotal,
		DocDate: parInvoiceDate
	}

	Context.LogMessage('update data info: ' + JSON.stringify(dataObject))
	return JSON.stringify(dataObject);
}
```

Adjust the dataObject to map all fields required to be exported in SharePoint. 

---

# SharePoint Mapping Rules

## IMPORTANT

The left side must be the SharePoint Internal Name.

Example:

```javascript
VendorName: parVendor
```

| SharePoint Column | ABBYY Field |
|---|---|
| VendorName | Vendor |
| InvoiceNumber | Invoice Number |
| DocDate | Invoice Date |

---

# Adding Additional Fields

## Step 1 — Create SharePoint Column

Example:

| Display Name | Internal Name |
|---|---|
| Vendor Name | VendorName |

---

## Step 2 — Read ABBYY Field

```javascript
const parVendor = doc.getField("Vendor").Value;
```

---

## Step 3 — Add Mapping

```javascript
let dataObject = {

	InvoiceNumber: parInvoiceNumber,
	Total: parTotal,
	DocDate: parInvoiceDate,
	VendorName: parVendor
}
```

---

# Complete Mapping Example

```javascript
function createUpdateItemData(doc) {

	const parInvoiceNumber = doc.getField("Invoice Number").Value;
	const parTotal = doc.getField("Total").Value;
	const parInvoiceDate = doc.getField("Invoice Date").Value;
	const parVendor = doc.getField("Vendor").Value;
	const parPoNumber = doc.getField("PO Number").Value;

	let dataObject = {

		InvoiceNumber: parInvoiceNumber,
		Total: parTotal,
		DocDate: parInvoiceDate,
		VendorName: parVendor,
		PONumber: parPoNumber
	}

	Context.LogMessage('update data info: ' + JSON.stringify(dataObject))
	return JSON.stringify(dataObject);
}
```

---

# Supported SharePoint Field Types

| Type | Supported |
|---|---|
| Text | Yes |
| Number | Yes |
| Currency | Yes |
| Date | Yes |
| Yes/No | Yes |

## Complex Types

The following types may require additional formatting:

- People Picker
- Lookup
- Managed Metadata
- Taxonomy
- Multi-select fields

---

# Troubleshooting

---

## 401 Unauthorized

Possible causes:

- Invalid client secret
- Missing Graph permissions
- Admin consent not granted

---

## 403 Forbidden

Possible causes:

- SharePoint permissions denied
- Site access restrictions

---

## 404 Not Found

Possible causes:

- Invalid Site ID
- Invalid Drive ID
- Invalid List ID

---

## Metadata Not Updating

Possible causes:

- Wrong SharePoint Internal Name
- Unsupported field type

Use:

```http
GET /columns
```

to validate internal names.

---

# Security Recommendations

- Store credentials as ABBYY Secrets
- Never hardcode secrets
- Rotate Azure secrets periodically
- Use least privilege permissions

---

# Example Processing Flow

```text
Input Invoice PDF
        |
        v
ABBYY Extraction
        |
        v
Custom Export Activity
        |
        +--> Upload PDF to SharePoint
        |
        +--> Update Metadata
        |
        v
SharePoint Library
```

---

# Included Files

| File | Description |
|---|---|
| SharePoint API.postman_collection.json | Graph API helper collection |
| script sharepoint export.js | ABBYY connector script |

---


# Conclusion

This connector enables seamless integration between ABBYY Vantage and Microsoft SharePoint Online using Microsoft Graph API, allowing automated document archiving and metadata synchronization directly into SharePoint document libraries.
````

