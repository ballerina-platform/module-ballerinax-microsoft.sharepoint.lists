# Stale Items Archival Audit

This example demonstrates how to automate a SharePoint list compliance audit by identifying stale list items (not modified since a defined cutoff date), reviewing their version history, and archiving them by updating their fields with an `Archived` status and archive date.

## Prerequisites

1. **Microsoft SharePoint Setup**
   > Refer to the [Microsoft SharePoint Lists connector setup guide](https://central.ballerina.io/ballerinax/microsoft.sharepoint.lists/latest) to register an Azure AD application and obtain the required credentials.

2. Ensure your SharePoint list contains the custom fields `Status`, `ArchiveDate`, and `ArchiveComment` to support the archival workflow.

3. Create a `Config.toml` file in the project root with your credentials:

```toml
clientId = "<Your Client ID>"
clientSecret = "<Your Client Secret>"
tenantId = "<Your Tenant ID>"
siteId = "<Your Site ID>"
listId = "<Your List ID>"
```

## Run the Example

Execute the following command to run the example. The script will print its audit and archival progress to the console.

```shell
bal run
```
