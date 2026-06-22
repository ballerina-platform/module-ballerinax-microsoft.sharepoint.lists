# Ballerina Microsoft SharePoint Lists connector

[![Build](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/actions/workflows/ci.yml/badge.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/actions/workflows/ci.yml)
[![GitHub Last Commit](https://img.shields.io/github/last-commit/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists.svg)](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/commits/master)
[![GitHub Issues](https://img.shields.io/github/issues/ballerina-platform/ballerina-library/module/microsoft.sharepoint.lists.svg?label=Open%20Issues)](https://github.com/ballerina-platform/ballerina-library/labels/module%microsoft.sharepoint.lists)

## Overview

[Microsoft SharePoint](https://www.microsoft.com/en-us/microsoft-365/sharepoint/collaboration) is a cloud-based collaboration and content management platform that enables organizations to create, manage, and share information through sites, lists, libraries, and workflows across teams and enterprises.

The `ballerinax/microsoft.sharepoint.lists` package offers APIs to connect and interact with [Microsoft SharePoint Lists API](https://learn.microsoft.com/en-us/graph/api/resources/list?view=graph-rest-1.0) endpoints, specifically based on [Microsoft Graph REST API v1.0](https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0).

## Setup guide

To use the Microsoft SharePoint Lists connector, you must have access to the Microsoft SharePoint API through a [Microsoft Azure developer account](https://portal.azure.com/) and obtain client credentials by registering an application in Microsoft Entra ID. If you do not have a Microsoft account, you can sign up for one at the [Microsoft account sign-up page](https://account.microsoft.com/account).

### Step 1: Create a Microsoft Account and Set Up SharePoint Access

1. Navigate to the [Microsoft 365 website](https://www.microsoft.com/en-us/microsoft-365) and sign up for an account or log in if you already have one.

2. Ensure you have a Microsoft 365 Business Basic, Business Standard, Business Premium, or an Enterprise (E1, E3, or E5) plan, as SharePoint Online and its API capabilities are restricted to users on these plans.

### Step 2: Register an Application and Generate Credentials

1. Log in to the [Microsoft Azure Portal](https://portal.azure.com/) using your Microsoft account credentials.

2. In the left-hand navigation menu, select **Microsoft Entra ID** in the top search bar.

3. In the left panel, navigate to **App registrations** and click **New registration**.

   ![New application registration](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/refs/heads/main/docs/resources/new-application-registration.png)

4. Enter a name for your application, select the appropriate **Supported account types** (e.g., "Single tenant only"), and click **Register**.

   ![Application registration details](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/refs/heads/main/docs/resources/application-registration-details.png)

5. Once the application is registered, note down the **Application (client) ID** and **Directory (tenant) ID** from the Overview page.

   ![Client ID and Tenant ID](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/refs/heads/main/docs/resources/client-id-and-tenant-id.png)

6. Navigate to **Certificates & secrets** in the left panel, click **New client secret**, provide a description and expiry period, then click **Add**. Copy the generated **client secret value** immediately.

   ![Create client secret](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/refs/heads/main/docs/resources/create-client-secret.png)

7. Navigate to **API permissions** in the left panel and click **Add a permission**.

   ![Add API permission](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/refs/heads/main/docs/resources/add-api-permission.png)

8. Select **Microsoft Graph** from the available API options.

   ![Microsoft Graph API permission](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/refs/heads/main/docs/resources/microsoft-graph-api-permission.png)

9. Select **Application permissions**, then search for and add the following permissions depending on your use case, then click **Add permissions**.

   | Permission | Operations covered |
   |---|---|
   | `Sites.Read.All` | Read sites, lists, columns, content types, drives, analytics |
   | `Sites.ReadWrite.All` | Create and update lists, list items, drives, and content |
   | `Sites.Manage.All` | Update site properties, create/delete columns and content types |
   | `Sites.FullControl.All` | Manage site permissions |

   > **Tip:** Grant only the permissions your application actually requires. For read-only use cases, `Sites.Read.All` is sufficient. For full connector coverage, add all four.

   ![API site permissions](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/refs/heads/main/docs/resources/api-site-permissions.png)

10. Click **Grant admin consent** to approve the permissions for your organization.

    ![Grant admin consent](https://raw.githubusercontent.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/refs/heads/main/docs/resources/grant-admin-consent.png)

11. Construct the `tokenUrl` using the **Directory (tenant) ID** obtained in step 5:

```text
https://login.microsoftonline.com/<TENANT_ID>/oauth2/v2.0/token
```

This is the OAuth 2.0 token endpoint the connector uses to exchange your `clientId` and `clientSecret` for an access token with the `https://graph.microsoft.com/.default` scope.

## Quickstart

To use the `microsoft.sharepoint.lists` connector in your Ballerina application, update the `.bal` file as follows:

### Step 1: Import the module

```ballerina
import ballerinax/microsoft.sharepoint.lists;
```

### Step 2: Instantiate a new connector

1. Create a `Config.toml` file and configure the credentials obtained above:

```toml
clientId = "<CLIENT_ID>"
clientSecret = "<CLIENT_SECRET>"
tokenUrl = "https://login.microsoftonline.com/<TENANT_ID>/oauth2/v2.0/token"
```

2. Instantiate a `lists:Client` with the obtained credentials.

```ballerina
configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string tokenUrl = ?;

final lists:Client listsClient = check new ({
    auth: {
        clientId,
        clientSecret,
        tokenUrl,
        scopes: "https://graph.microsoft.com/.default"
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

## Build from the source

### Prerequisites

1. Download and install Java SE Development Kit (JDK) version 17. You can download it from either of the following sources:
   - [Oracle JDK](https://www.oracle.com/java/technologies/downloads/)
   - [OpenJDK](https://adoptium.net/)

   > **Note:** Set the `JAVA_HOME` environment variable to the path of the JDK installation.

2. Download and install [Ballerina Swan Lake](https://ballerina.io/).

3. Download and install [Docker](https://www.docker.com/get-started).

### Build options

Execute the commands below to build from the source.

1. To build the package:

   ```bash
   ./gradlew clean build
   ```

2. To run the tests:

   ```bash
   ./gradlew clean test
   ```

3. To build without the tests:

   ```bash
   ./gradlew clean build -x test
   ```

4. To run tests against different environments:

   ```bash
   ./gradlew clean test -Pgroups=<Comma separated groups/test cases>
   ```

5. To debug the package with a remote debugger:

   ```bash
   ./gradlew clean build -Pdebug=<port>
   ```

6. To debug with the Ballerina language:

   ```bash
   ./gradlew clean build -PbalJavaDebug=<port>
   ```

7. Publish the generated artifacts to the local Ballerina Central repository:

   ```bash
   ./gradlew clean build -PpublishToLocalCentral=true
   ```

8. Publish the generated artifacts to the Ballerina Central repository:

   ```bash
   ./gradlew clean build -PpublishToCentral=true
   ```

## Contributing to Ballerina

As an open source project, Ballerina welcomes contributions from the community.

For more information, go to the [contribution guidelines](https://github.com/ballerina-platform/ballerina-lang/blob/master/CONTRIBUTING.md).

## Code of conduct

All contributors are encouraged to read the [Ballerina Code of Conduct](https://ballerina.io/code-of-conduct).


## Useful links

* For more information go to the [`microsoft.sharepoint.lists` package](https://central.ballerina.io/ballerinax/microsoft.sharepoint.lists/latest).
* For example demonstrations of the usage, go to [Ballerina By Examples](https://ballerina.io/learn/by-example/).
* Chat live with us via our [Discord server](https://discord.gg/ballerinalang).
* Post all technical questions on Stack Overflow with the [#ballerina](https://stackoverflow.com/questions/tagged/ballerina) tag.
