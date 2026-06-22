# Examples

The `microsoft.sharepoint.lists` connector provides practical examples illustrating usage in various scenarios. Explore these [examples](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples), covering use cases like compliance snapshot verification, SharePoint webhook subscription lifecycle, content type compliance audit, and stale items archival audit.

1. [Compliance snapshot verification](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples/compliance-snapshot-verification) - Verify compliance snapshots by retrieving and validating list data from SharePoint against defined compliance criteria.

2. [SharePoint webhook subscription lifecycle](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples/sharepoint-webhook-subscription-lifecycle) - Manage SharePoint webhook subscriptions, including creation, verification, and extension.

3. [Content type compliance audit](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples/content-type-compliance-audit) - Audit SharePoint list items to ensure they conform to the required content types and compliance standards.

4. [Stale items archival audit](https://github.com/ballerina-platform/module-ballerinax-microsoft.sharepoint.lists/tree/main/examples/stale-items-archival-audit) - Identify and audit stale items in SharePoint lists that are candidates for archival based on inactivity or age.

## Prerequisites

1. Generate Microsoft SharePoint credentials to authenticate the connector as described in the [Setup guide](https://central.ballerina.io/ballerinax/microsoft.sharepoint.lists/latest#setup-guide).

2. For each example, create a `Config.toml` file with the required configuration. The following fields are common across all examples:

    ```toml
    tenantId = "<Your_Tenant_Id>"
    clientId = "<Your_Client_Id>"
    clientSecret = "<Your_Client_Secret>"
    siteId = "<Your_SharePoint_Site_Id>"
    listId = "<Your_SharePoint_List_Id>"
    ```

    Some examples require additional fields. Refer to the `main.bal` of each example for the full list of required configurables.

## Running an Example

Execute the following commands to build an example from the source:

* To build an example:

    ```bash
    bal build
    ```

* To run an example:

    ```bash
    bal run
    ```
