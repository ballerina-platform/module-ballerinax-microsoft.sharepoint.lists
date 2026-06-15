## Overview

[Microsoft SharePoint](https://www.microsoft.com/en-us/microsoft-365/sharepoint/collaboration) is a collaborative platform that enables organizations to create, manage, and share content, documents, and data through customizable sites, lists, and libraries deeply integrated with the Microsoft 365 ecosystem.

The `ballerinax/microsoft.sharepoint.lists` package offers APIs to connect and interact with the [Microsoft SharePoint Lists API](https://learn.microsoft.com/en-us/graph/api/resources/list?view=graph-rest-1.0) endpoints, specifically based on [Microsoft Graph REST API v1.0](https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0).

## Setup guide

To use the Microsoft SharePoint Lists connector, you must have access to the Microsoft SharePoint API through a [Microsoft Azure developer account](https://portal.azure.com/) and obtain the necessary OAuth 2.0 credentials (Client ID, Client Secret, and Tenant ID). If you do not have a Microsoft account, you can sign up for one [here](https://signup.microsoft.com/).

### Step 1: Create a Microsoft Account and Set Up SharePoint

1. Navigate to the [Microsoft 365 website](https://www.microsoft.com/en-us/microsoft-365) and sign up for an account or log in if you already have one.

2. Ensure you have a Microsoft 365 Business Basic, Business Standard, Business Premium, or an Enterprise (E1, E3, or E5) plan, as access to SharePoint and its APIs requires an active Microsoft 365 subscription that includes SharePoint Online.

### Step 2: Register an Application and Generate API Credentials

1. Log in to the [Microsoft Azure Portal](https://portal.azure.com/) using your Microsoft account.

2. In the left-hand navigation menu, select **Azure Active Directory** (or search for **Microsoft Entra ID** in the top search bar).

3. In the left panel, select **App registrations**, then click **+ New registration**.

4. Fill in the registration form:
   - Enter a name for your application.
   - Select the appropriate **Supported account types** (e.g., *Accounts in this organizational directory only*).
   - Optionally, provide a Redirect URI, then click **Register**.

5. Once the app is registered, note down the **Application (client) ID** and the **Directory (tenant) ID** displayed on the app's Overview page — these are your Client ID and Tenant ID.

6. In the left panel of your app registration, navigate to **Certificates & secrets**, then select **Client secrets** and click **+ New client secret**.

7. Provide a description and select an expiry duration, then click **Add**. Copy the generated **Value** immediately — this is your Client Secret.

8. In the left panel, navigate to **API permissions**, click **+ Add a permission**, select **Microsoft Graph**, and grant the required SharePoint-related permissions (e.g., `Sites.Read.All`, `Sites.ReadWrite.All`). Click **Grant admin consent** to activate the permissions.

9. Construct your `tokenUrl` using the **Directory (tenant) ID** noted in step 5:

   ```
   https://login.microsoftonline.com/<tenantId>/oauth2/v2.0/token
   ```

> **Tip:** You must copy and store the Client Secret value somewhere safe. It won't be visible again in the Azure Portal after you navigate away from the page, for security reasons.

## Quickstart
To use the `microsoft.sharepoint.lists` connector in your Ballerina application, update the `.bal` file as follows:

### Step 1: Import the module

```ballerina
import ballerinax/microsoft.sharepoint.lists;
```

### Step 2: Instantiate a new connector

1. Create a `Config.toml` file and configure the obtained credentials:

```toml
clientId = "<Your_Client_Id>"
clientSecret = "<Your_Client_Secret>"
tokenUrl = "<Your_Token_Url>"
```

2. Create a `lists:ConnectionConfig` and initialize the client:

```ballerina
configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string tokenUrl = ?;

final lists:Client listsClient = check new ({
    auth: {
        tokenUrl,
        clientId,
        clientSecret,
        scopes: ["https://graph.microsoft.com/.default"]
    }
});
```

### Step 3: Invoke the connector operation

Now, utilize the available connector operations.

#### Create a new list

```ballerina
public function main() returns error? {
    lists:MicrosoftGraphList newList = {
        displayName: "Project Tasks",
        list: {
            template: "genericList",
            contentTypesEnabled: false
        }
    };

    lists:MicrosoftGraphList response = check listsClient->createList("contoso.sharepoint.com,abc123,def456", newList);
}
```

### Step 4: Run the Ballerina application

```bash
bal run
```

## Examples
The `microsoft.sharepoint.lists` connector provides practical examples illustrating usage in various scenarios. Explore these [examples](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples), covering the following use cases:

1. [Compliance snapshot verification](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples/compliance-snapshot-verification) - Demonstrates how to capture and verify compliance snapshots of SharePoint list data.
2. [Sharepoint webhook subscription lifecycle](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples/sharepoint-webhook-subscription-lifecycle) - Illustrates creating, verifying, and extending SharePoint webhook subscriptions to manage event-driven notifications.
3. [Content type compliance audit](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples/content-type-compliance-audit) - Demonstrates how to audit SharePoint list items for adherence to expected content type configurations.
4. [Stale items archival audit](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples/stale-items-archival-audit) - Illustrates identifying and archiving stale items from SharePoint lists based on inactivity criteria.
