# SharePoint Webhook Subscription Lifecycle

This example demonstrates how to manage the complete lifecycle of a SharePoint list webhook subscription using the Microsoft SharePoint Lists connector. The script creates a webhook subscription to monitor list changes, verifies the subscription details, and extends its expiration date to ensure uninterrupted notifications.

## Prerequisites

1. **Microsoft SharePoint Setup**
   > Refer to the [Microsoft SharePoint Lists connector setup guide](https://central.ballerina.io/ballerinax/microsoft.sharepoint.lists/latest) to obtain the required credentials, including your tenant ID, client ID, and client secret.

2. **Notification Endpoint**
   > You must have a publicly accessible HTTPS endpoint that SharePoint can send webhook notifications to. This URL must respond to SharePoint's validation request during subscription creation.

3. For this example, create a `Config.toml` file in the project root with your credentials:

```toml
tenantId = "<Your Tenant ID>"
clientId = "<Your Client ID>"
clientSecret = "<Your Client Secret>"
siteId = "<Your SharePoint Site ID>"
listId = "<Your SharePoint List ID>"
notificationUrl = "<Your Notification URL>"
```

## Run the Example

Execute the following command to run the example. The script will print its progress to the console as it creates, verifies, and updates the webhook subscription.

```shell
bal run
```
