# Content Type Compliance Audit

This example demonstrates how to perform a governance audit on SharePoint list content types by retrieving all content types, checking their column counts against a configurable minimum threshold, and reporting detailed column definitions for any non-compliant content types.

## Prerequisites

1. **Microsoft SharePoint Setup**
   > Refer to the [Microsoft SharePoint Lists connector setup guide](https://central.ballerina.io/ballerinax/microsoft.sharepoint.lists/latest) to obtain the required credentials and configure access to your SharePoint site.

2. For this example, create a `Config.toml` file in the project root with your credentials:

```toml
tenantId = "<Your Tenant ID>"
clientId = "<Your Client ID>"
clientSecret = "<Your Client Secret>"
siteId = "<Your Site ID>"
listId = "<Your List ID>"
minimumRequiredColumns = 5
```

## Run the Example

Execute the following command to run the example. The script will print its audit progress and compliance summary to the console.

```shell
bal run
```
