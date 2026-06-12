# Compliance Snapshot Verification

This example demonstrates how to automate compliance archiving for SharePoint document sets by creating a point-in-time version snapshot and verifying its metadata fields — providing an immutable audit record suitable for regulatory compliance purposes.

## Prerequisites

1. **Microsoft SharePoint Setup**
   > Refer to the [Microsoft SharePoint Lists connector setup guide](https://central.ballerina.io/ballerinax/microsoft.sharepoint.lists/latest) to register an Azure AD application and obtain the required credentials.

2. For this example, create a `Config.toml` file in the project root with your credentials:

```toml
tenantId = "<Your Tenant ID>"
clientId = "<Your Client ID>"
clientSecret = "<Your Client Secret>"
siteId = "<Your SharePoint Site ID>"
listId = "<Your SharePoint List ID>"
listItemId = "<Your List Item ID>"
```

## Run the Example

Execute the following command to run the example. The script will print its progress and a compliance workflow summary to the console.

```shell
bal run
```
