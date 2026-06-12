import ballerina/io;
import ballerinax/microsoft.sharepoint.lists;

configurable string tenantId = ?;
configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string siteId = ?;
configurable string listId = ?;
configurable string notificationUrl = ?;

public function main() returns error? {
    lists:Client sharepointClient = check new ({
        auth: {
            tokenUrl: string `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
            clientId: clientId,
            clientSecret: clientSecret,
            scopes: ["https://graph.microsoft.com/.default"]
        }
    });

    io:println("=== Webhook Subscription Lifecycle Management for SharePoint List ===");
    io:println("");

    io:println("Step 1: Creating webhook subscription for the project tracking list...");

    lists:MicrosoftGraphSubscription subscriptionPayload = {
        changeType: "created,updated",
        notificationUrl: notificationUrl,
        expirationDateTime: "2025-12-31T23:59:59Z",
        clientState: "secretClientState-12345"
    };

    lists:MicrosoftGraphSubscription createdSubscription = check sharepointClient->createSubscription(
        siteId,
        listId,
        subscriptionPayload
    );

    io:println("Subscription created successfully!");

    string|() createdId = createdSubscription?.id;
    string|() createdChangeType = createdSubscription?.changeType;
    string|() createdNotificationUrl = createdSubscription?.notificationUrl;
    string|() createdExpiration = createdSubscription?.expirationDateTime;
    string|() createdClientState = createdSubscription?.clientState;

    io:println("  Subscription ID       : ", createdId ?: "N/A");
    io:println("  Change Type           : ", createdChangeType ?: "N/A");
    io:println("  Notification URL      : ", createdNotificationUrl ?: "N/A");
    io:println("  Expiration DateTime   : ", createdExpiration ?: "N/A");
    io:println("  Client State          : ", createdClientState ?: "N/A");
    io:println("");

    string subscriptionId = createdId ?: "";
    if subscriptionId == "" {
        return error("Failed to retrieve subscription ID from created subscription.");
    }

    io:println("Step 2: Verifying subscription details using getSubscription...");

    lists:MicrosoftGraphSubscription retrievedSubscription = check sharepointClient->getSubscription(
        siteId,
        listId,
        subscriptionId,
        {},
        {
            dollarSelect: ["id", "changeType", "notificationUrl", "expirationDateTime", "clientState", "resource"]
        }
    );

    io:println("Subscription verified successfully!");

    string|() retrievedId = retrievedSubscription?.id;
    string|() retrievedChangeType = retrievedSubscription?.changeType;
    string|() retrievedNotificationUrl = retrievedSubscription?.notificationUrl;
    string|() retrievedExpiration = retrievedSubscription?.expirationDateTime;
    string|() retrievedClientState = retrievedSubscription?.clientState;

    io:println("  Subscription ID       : ", retrievedId ?: "N/A");
    io:println("  Change Type           : ", retrievedChangeType ?: "N/A");
    io:println("  Notification URL      : ", retrievedNotificationUrl ?: "N/A");
    io:println("  Expiration DateTime   : ", retrievedExpiration ?: "N/A");
    io:println("  Client State          : ", retrievedClientState ?: "N/A");
    io:println("");

    io:println("Step 3: Extending subscription expiration date to ensure continuity...");

    lists:MicrosoftGraphSubscription updatePayload = {
        expirationDateTime: "2026-06-30T23:59:59Z"
    };

    error? updateResult = sharepointClient->updateSubscription(
        siteId,
        listId,
        subscriptionId,
        updatePayload
    );

    if updateResult is error {
        io:println("Failed to update subscription: ", updateResult.message());
        return updateResult;
    }

    io:println("Subscription expiration extended successfully!");
    io:println("  Subscription ID       : ", subscriptionId);
    io:println("  New Expiration Date   : 2026-06-30T23:59:59Z");
    io:println("");

    io:println("=== Webhook Subscription Lifecycle Management Completed ===");
    io:println("");
    io:println("Summary:");
    io:println("  - Created a webhook subscription to monitor list changes (items created/updated)");
    io:println("  - Verified subscription details including ID, notification URL, and expiration");
    io:println("  - Extended the subscription expiration to ensure uninterrupted notifications");
    io:println("");
    io:println("The external system at '", notificationUrl, "' will now receive real-time");
    io:println("notifications whenever items are added or modified in the SharePoint list.");
    io:println("Use the clientState value to validate incoming notifications for security.");
}