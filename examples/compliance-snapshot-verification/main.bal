import ballerina/io;
import ballerinax/microsoft.sharepoint.lists;

configurable string tenantId = ?;
configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string siteId = ?;
configurable string listId = ?;
configurable string listItemId = ?;

public function main() returns error? {

    lists:ConnectionConfig connectionConfig = {
        auth: <lists:OAuth2ClientCredentialsGrantConfig>{
            tokenUrl: string `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
            clientId: clientId,
            clientSecret: clientSecret,
            scopes: ["https://graph.microsoft.com/.default"]
        }
    };

    lists:Client sharepointClient = check new (connectionConfig);

    io:println("=== Document Set Version Management for Compliance Archiving ===");
    io:println("");

    io:println("Step 1: Creating a document set version snapshot for compliance archiving...");
    io:println("Site ID: ", siteId);
    io:println("List ID: ", listId);
    io:println("List Item ID (Contract Document Set): ", listItemId);
    io:println("");

    lists:MicrosoftGraphDocumentSetVersion documentSetVersionPayload = {
        comment: "Point-in-time snapshot captured for compliance audit - Contract signing milestone",
        shouldCaptureMinorVersion: false
    };

    lists:MicrosoftGraphDocumentSetVersion createdVersion = check sharepointClient->createItemDocumentSetVersion(
        siteId,
        listId,
        listItemId,
        documentSetVersionPayload
    );

    io:println("Document set version snapshot created successfully.");
    io:println("Version ID: ", createdVersion?.id ?: "N/A");
    io:println("Comment: ", createdVersion?.comment ?: "N/A");
    io:println("");

    string documentSetVersionId = createdVersion?.id ?: "";

    if documentSetVersionId == "" {
        io:println("Warning: No version ID returned. Cannot retrieve version fields.");
        return;
    }

    io:println("Step 2: Retrieving fields of the newly created document set version to verify metadata...");
    io:println("Document Set Version ID: ", documentSetVersionId);
    io:println("");

    lists:MicrosoftGraphFieldValueSet versionFields = check sharepointClient->getItemDocumentSetVersionFields(
        siteId,
        listId,
        listItemId,
        documentSetVersionId
    );

    io:println("Document set version fields retrieved successfully.");
    io:println("Field Value Set ID: ", versionFields?.id ?: "N/A");
    io:println("");

    io:println("=== Compliance Archiving Workflow Summary ===");
    io:println("A point-in-time snapshot of the contract document set has been created.");
    io:println("Version ID: ", createdVersion?.id ?: "N/A");
    io:println("Snapshot Comment: ", createdVersion?.comment ?: "N/A");
    io:println("Field Metadata Verified: true");
    io:println("");
    io:println("The document set version is now preserved as an immutable audit record.");
    io:println("This snapshot can be referenced for regulatory compliance and audit trail purposes.");
    io:println("");
    io:println("=== Workflow Completed Successfully ===");
}