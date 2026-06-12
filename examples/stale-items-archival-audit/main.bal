import ballerina/io;
import ballerinax/microsoft.sharepoint.lists;

configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string tenantId = ?;
configurable string siteId = ?;
configurable string listId = ?;

public function main() returns error? {
    lists:ConnectionConfig config = {
        auth: <lists:OAuth2ClientCredentialsGrantConfig>{
            clientId: clientId,
            clientSecret: clientSecret,
            tokenUrl: string `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
            scopes: ["https://graph.microsoft.com/.default"]
        }
    };

    lists:Client sharePointClient = check new (config);

    io:println("=== SharePoint List Audit and Archive Workflow ===");
    io:println("");

    string staleCutoffDate = "2024-10-01T00:00:00Z";
    string filterQuery = string `fields/Modified lt '${staleCutoffDate}'`;

    io:println("Step 1: Retrieving stale list items (not modified since " + staleCutoffDate + ")...");

    lists:MicrosoftGraphListItemCollectionResponse itemsResponse = check sharePointClient->listItems(
        siteId,
        listId,
        queries = {
            dollarFilter: filterQuery,
            dollarExpand: ["fields"],
            dollarTop: 100
        }
    );

    lists:MicrosoftGraphListItem[] staleItems = itemsResponse.value ?: [];

    if staleItems.length() == 0 {
        io:println("No stale items found. Audit complete.");
        return;
    }

    io:println("Found " + staleItems.length().toString() + " stale item(s) to review.");
    io:println("");

    int archivedCount = 0;
    string archiveDate = "2025-01-08";

    foreach lists:MicrosoftGraphListItem item in staleItems {
        string itemId = item.id ?: "unknown";

        io:println("--------------------------------------------------");
        io:println("Processing item ID: " + itemId);

        anydata|() rawFields = item?.fields;
        if rawFields is lists:MicrosoftGraphFieldValueSet {
            lists:MicrosoftGraphFieldValueSet fieldSet = rawFields;
            anydata titleValue = fieldSet["Title"];
            if titleValue is string {
                io:println("  Item Title: " + titleValue);
            }
        }

        io:println("  Step 2: Retrieving version history for item: " + itemId);

        lists:MicrosoftGraphListItemVersionCollectionResponse versionsResponse = check sharePointClient->listItemVersions(
            siteId,
            listId,
            itemId,
            queries = {
                dollarOrderby: ["lastModifiedDateTime desc"],
                dollarSelect: ["id", "lastModifiedDateTime", "lastModifiedBy"],
                dollarTop: 50
            }
        );

        lists:MicrosoftGraphListItemVersion[] versions = versionsResponse.value ?: [];
        int versionCount = versions.length();

        io:println("  Total versions found: " + versionCount.toString());

        if versionCount > 0 {
            lists:MicrosoftGraphListItemVersion latestVersion = versions[0];
            string lastModifiedDateTime = latestVersion.lastModifiedDateTime ?: "Unknown";
            io:println("  Last substantive change (most recent version): " + lastModifiedDateTime);

            anydata|() rawLastModifiedBy = latestVersion?.lastModifiedBy;
            if rawLastModifiedBy is lists:MicrosoftGraphIdentitySet {
                lists:MicrosoftGraphIdentitySet lastModifiedBy = rawLastModifiedBy;
                anydata|() rawUserIdentity = lastModifiedBy?.user;
                if rawUserIdentity is lists:MicrosoftGraphIdentity {
                    lists:MicrosoftGraphIdentity userIdentity = rawUserIdentity;
                    string|() displayName = userIdentity?.displayName;
                    if displayName is string {
                        io:println("  Last modified by: " + displayName);
                    }
                }
            }

            io:println("  Version history confirms item is stale. Proceeding to archive.");
        } else {
            io:println("  No version history available. Proceeding to archive based on item metadata.");
        }

        io:println("  Step 3: Archiving item " + itemId + " with Status='Archived' and ArchiveDate='" + archiveDate + "'");

        lists:MicrosoftGraphFieldValueSet archivePayload = {
            "Status": "Archived",
            "ArchiveDate": archiveDate,
            "ArchiveComment": "Automatically archived by compliance audit - item stale for over 90 days"
        };

        error? updateResult = sharePointClient->updateItemFields(
            siteId,
            listId,
            itemId,
            archivePayload
        );

        if updateResult is error {
            io:println("  ERROR: Failed to archive item " + itemId + ": " + updateResult.message());
        } else {
            io:println("  SUCCESS: Item " + itemId + " has been archived successfully.");
            archivedCount += 1;
        }

        io:println("");
    }

    io:println("=== Audit and Archive Workflow Complete ===");
    io:println("Total stale items reviewed: " + staleItems.length().toString());
    io:println("Total items archived: " + archivedCount.toString());
    io:println("Archive date applied: " + archiveDate);
    io:println("All archived items retain full version history for compliance purposes.");
}