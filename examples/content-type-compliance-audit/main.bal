// Copyright (c) 2026, WSO2 LLC. (http://www.wso2.com).
//
// WSO2 LLC. licenses this file to you under the Apache License,
// Version 2.0 (the "License"); you may not use this file except
// in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
// KIND, either express or implied.  See the License for the
// specific language governing permissions and limitations
// under the License.

import ballerina/io;
import ballerinax/microsoft.sharepoint.lists;

configurable string tenantId = ?;
configurable string clientId = ?;
configurable string clientSecret = ?;
configurable string siteId = ?;
configurable string listId = ?;
configurable int minimumRequiredColumns = 5;

public function main() returns error? {
    lists:ConnectionConfig config = {
        auth: <lists:OAuth2ClientCredentialsGrantConfig>{
            tokenUrl: string `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
            clientId: clientId,
            clientSecret: clientSecret,
            scopes: ["https://graph.microsoft.com/.default"]
        }
    };

    lists:Client sharepointClient = check new (config);

    io:println("=== Content Type Governance Audit ===");
    io:println("Site ID: ", siteId);
    io:println("List ID: ", listId);
    io:println("Minimum Required Columns: ", minimumRequiredColumns);
    io:println("=====================================");

    io:println("\n[Step 1] Retrieving all content types for the list...");
    lists:MicrosoftGraphContentTypeCollectionResponse contentTypesResponse =
        check sharepointClient->listContentTypes(siteId, listId);

    lists:MicrosoftGraphContentType[] contentTypes = contentTypesResponse.value ?: [];

    if contentTypes.length() == 0 {
        io:println("No content types found for the specified list.");
        return;
    }

    io:println("Found ", contentTypes.length(), " content type(s) on the list.");

    string[] nonCompliantContentTypeIds = [];
    string[] nonCompliantContentTypeNames = [];

    io:println("\n[Step 2] Auditing column counts for each content type...");
    io:println("------------------------------------------------------------");

    foreach lists:MicrosoftGraphContentType contentType in contentTypes {
        string contentTypeId = contentType?.id ?: "unknown-id";
        string contentTypeName = contentType?.name ?: "Unknown Content Type";
        string contentTypeDescription = contentType?.description ?: "No description";

        io:println("\nContent Type: ", contentTypeName);
        io:println("  ID         : ", contentTypeId);
        io:println("  Description: ", contentTypeDescription);

        if contentTypeId == "unknown-id" {
            io:println("  [WARNING] Content type has no ID, skipping column count check.");
            continue;
        }

        string|error columnCountResult = sharepointClient->getContentTypeColumnsCount(
            siteId, listId, contentTypeId
        );

        if columnCountResult is error {
            io:println("  [ERROR] Failed to retrieve column count: ", columnCountResult.message());
            continue;
        }

        string columnCountStr = columnCountResult;
        io:println("  Column Count: ", columnCountStr);

        int|error columnCount = int:fromString(columnCountStr);
        if columnCount is error {
            io:println("  [WARNING] Could not parse column count value '", columnCountStr, "', skipping compliance check.");
            continue;
        }

        if columnCount < minimumRequiredColumns {
            io:println("  [NON-COMPLIANT] Content type has ", columnCount,
                " column(s), which is below the minimum required ", minimumRequiredColumns, ".");
            nonCompliantContentTypeIds.push(contentTypeId);
            nonCompliantContentTypeNames.push(contentTypeName);
        } else {
            io:println("  [COMPLIANT] Content type meets the minimum column requirement.");
        }
    }

    io:println("\n[Step 3] Retrieving detailed column definitions for non-compliant content types...");
    io:println("------------------------------------------------------------");

    if nonCompliantContentTypeIds.length() == 0 {
        io:println("All content types are compliant. No further action required.");
        return;
    }

    io:println("Non-compliant content types found: ", nonCompliantContentTypeIds.length());

    foreach int i in 0 ..< nonCompliantContentTypeIds.length() {
        string nonCompliantId = nonCompliantContentTypeIds[i];
        string nonCompliantName = nonCompliantContentTypeNames[i];

        io:println("\n--- Non-Compliant Content Type: ", nonCompliantName, " (ID: ", nonCompliantId, ") ---");

        lists:MicrosoftGraphColumnDefinitionCollectionResponse|error columnsResponse =
            sharepointClient->listContentTypeColumns(siteId, listId, nonCompliantId);

        if columnsResponse is error {
            io:println("  [ERROR] Failed to retrieve column definitions: ", columnsResponse.message());
            continue;
        }

        lists:MicrosoftGraphColumnDefinitionCollectionResponse columnsResponseVal = columnsResponse;
        lists:MicrosoftGraphColumnDefinition[] columns = columnsResponseVal.value ?: [];

        if columns.length() == 0 {
            io:println("  No columns found for this content type.");
        } else {
            io:println("  Existing Columns (", columns.length(), " total):");
            foreach lists:MicrosoftGraphColumnDefinition column in columns {
                string columnId = column?.id ?: "unknown-id";
                string columnName = column?.name ?: "Unknown Column";
                string columnDisplayName = column?.displayName ?: columnName;
                boolean isRequired = column?.required ?: false;
                boolean isHidden = column?.hidden ?: false;

                io:println("    - Column Name    : ", columnName);
                io:println("      Display Name   : ", columnDisplayName);
                io:println("      ID             : ", columnId);
                io:println("      Required       : ", isRequired);
                io:println("      Hidden         : ", isHidden);
            }
        }

        int missingColumns = minimumRequiredColumns - columns.length();
        if missingColumns > 0 {
            io:println("  [ACTION REQUIRED] This content type is missing at least ",
                missingColumns, " column(s) to meet compliance standards.");
            io:println("  Please review and add the required metadata columns to '", nonCompliantName, "'.");
        }
    }

    io:println("\n=== Governance Audit Complete ===");
    io:println("Summary:");
    io:println("  Total content types audited : ", contentTypes.length());
    io:println("  Non-compliant content types : ", nonCompliantContentTypeIds.length());
    io:println("  Compliant content types     : ", contentTypes.length() - nonCompliantContentTypeIds.length());

    if nonCompliantContentTypeIds.length() > 0 {
        io:println("\nNon-compliant content types requiring corrective action:");
        foreach int i in 0 ..< nonCompliantContentTypeIds.length() {
            io:println("  - ", nonCompliantContentTypeNames[i], " (ID: ", nonCompliantContentTypeIds[i], ")");
        }
    }
}
