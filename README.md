# microsoft.sharepoint.lists

[![Build](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/actions/workflows/ci.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/actions/workflows/ci.yml)
[![Trivy](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/actions/workflows/trivy-scan.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/actions/workflows/trivy-scan.yml)
[![GraalVM Check](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/actions/workflows/build-with-bal-test-graalvm.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/actions/workflows/build-with-bal-test-graalvm.yml)
[![GitHub Last Commit](https://img.shields.io/github/last-commit/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/commits/master)
[![GitHub Issues](https://img.shields.io/github/issues/ballerina-platform/ballerina-library/module/microsoft.sharepoint.lists.svg?label=Open%20Issues)](https://github.com/ballerina-platform/ballerina-library/labels/module%microsoft.sharepoint.lists)

## Overview
[Microsoft SharePoint](https://www.microsoft.com/en-us/microsoft-365/sharepoint/collaboration) is a cloud-based collaboration and content management platform that enables organizations to create, manage, and share information through sites, lists, libraries, and workflows across teams and enterprises.

The `ballerinax/microsoft.sharepoint.lists` package offers APIs to connect and interact with [Microsoft SharePoint Lists API](https://learn.microsoft.com/en-us/graph/api/resources/list?view=graph-rest-1.0) endpoints, specifically based on [Microsoft Graph REST API v1.0](https://learn.microsoft.com/en-us/graph/overview).

## Setup guide
To use the Microsoft SharePoint Lists connector, you must have access to the Microsoft SharePoint API through a [Microsoft Azure developer account](https://portal.azure.com/) and obtain the necessary OAuth 2.0 credentials (client ID, client secret, and tenant ID). If you do not have a Microsoft account, you can sign up for one [here](https://signup.microsoft.com/).

### Step 1: Create a Microsoft Account and Set Up SharePoint

1. Navigate to the [Microsoft 365 website](https://www.microsoft.com/en-us/microsoft-365) and sign up for an account or log in if you already have one.

2. Ensure you have a Microsoft 365 Business Basic, Business Standard, Business Premium, or an Enterprise (E1, E3, E5) plan, as access to SharePoint and its API requires an active Microsoft 365 subscription that includes SharePoint Online.

### Step 2: Register an Application and Generate API Credentials

1. Log in to the [Microsoft Azure Portal](https://portal.azure.com/) using your Microsoft account.

2. In the left-hand navigation panel, select **Azure Active Directory** (or search for **Microsoft Entra ID** in the top search bar).

3. In the left menu, select **App registrations**, then click **New registration**.

4. Enter a name for your application, choose the appropriate **Supported account types** (e.g., single tenant or multi-tenant), and click **Register**.

5. Once the application is registered, note down the **Application (client) ID** and the **Directory (tenant) ID** displayed on the app's Overview page — these are your `clientId` and `tenantId`.

6. In the left menu of your registered app, navigate to **Certificates & secrets**, then click **New client secret**. Provide a description and an expiry period, then click **Add**.

7. Copy the generated **client secret value** immediately — this is your `clientSecret`.

8. To grant the necessary permissions, go to **API permissions** in the left menu, click **Add a permission**, select **Microsoft Graph**, and add the required SharePoint permissions such as `Sites.Read.All`, `Sites.ReadWrite.All`, or `Lists.ReadWrite.All` depending on your use case. Click **Grant admin consent** to activate the permissions.

> **Tip:** You must copy and store the client secret value somewhere safe immediately after generation. It won't be visible again in the Azure Portal for security reasons.

## Quickstart
To use the `microsoft.sharepoint.lists` connector in your Ballerina application, update the `.bal` file as follows:

### Step 1: Import the module

```ballerina
import ballerina/oauth2;
import ballerinax/microsoft.sharepoint.lists as mssplists;
```

### Step 2: Instantiate a new connector

1. Create a `Config.toml` file and configure the obtained credentials:

```toml
clientId = "<Your_Client_Id>"
clientSecret = "<Your_Client_Secret>"
refreshToken = "<Your_Refresh_Token>"
```

2. Create a `mssplists:ConnectionConfig` and initialize the client:

```ballerina
configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string refreshToken = ?;

final mssplists:Client msSpListsClient = check new({
    auth: {
        clientId,
        clientSecret,
        refreshToken
    }
});
```

### Step 3: Invoke the connector operation

Now, utilize the available connector operations.

#### Create a new list

```ballerina
public function main() returns error? {
    mssplists:MicrosoftGraphList newList = {
        displayName: "Project Tasks",
        list: {
            template: "genericList",
            contentTypesEnabled: false
        }
    };

    mssplists:MicrosoftGraphList response = check msSpListsClient->createList("contoso.sharepoint.com,abc123,def456", newList);
}
```

### Step 4: Run the Ballerina application

```bash
bal run
```

## Examples
The `microsoft.sharepoint.lists` connector provides practical examples illustrating usage in various scenarios. Explore these [examples](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples), covering the following use cases:

1. [Compliance snapshot verification](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples/compliance-snapshot-verification) - Demonstrates how to capture and verify compliance snapshots of SharePoint list data.
2. [Sharepoint webhook subscription lifecycle](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples/sharepoint-webhook-subscription-lifecycle) - Illustrates creating, renewing, and deleting SharePoint webhook subscriptions to manage event-driven notifications.
3. [Content type compliance audit](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples/content-type-compliance-audit) - Demonstrates how to audit SharePoint list items for adherence to expected content type configurations.
4. [Stale items archival audit](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples/stale-items-archival-audit) - Illustrates identifying and archiving stale items from SharePoint lists based on inactivity criteria.

## Useful Links

* For more information go to the [`microsoft.sharepoint.lists` package](https://central.ballerina.io/ballerinax/microsoft.sharepoint.lists/latest).
* For example demonstrations of the usage, go to [Ballerina By Examples](https://ballerina.io/learn/by-example/).
* Chat live with us via our [Discord server](https://discord.gg/ballerinalang).
* Post all technical questions on Stack Overflow with the [#ballerina](https://stackoverflow.com/questions/tagged/ballerina) tag.
