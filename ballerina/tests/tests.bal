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

import ballerina/test;

configurable boolean isLiveServer = false;

configurable string clientId = "client-id";
configurable string clientSecret = "client-secret";
configurable string tenantId = "tenant-id";
configurable string siteId = "test-site-id";

final string subscriptionNotificationUrl = "https://example.com/webhook";
final string subscriptionExpiry = "2027-01-01T00:00:00Z";
final string updatedSubscriptionExpiry = "2027-06-01T00:00:00Z";

isolated string createdListId = "";
isolated string createdItemId = "";
isolated string createdSubscriptionId = "";

final Client sharepoint = check initClient();

isolated function initClient() returns Client|error {
    if isLiveServer {
        OAuth2ClientCredentialsGrantConfig auth = {
            tokenUrl: string `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
            clientId,
            clientSecret,
            scopes: ["https://graph.microsoft.com/.default"]
        };
        return check new Client({auth}, "https://graph.microsoft.com/v1.0/sites");
    }
    return check new Client({auth: {token: "test-token"}}, "http://localhost:9090");
}

// ===== List Tests =====

@test:Config {groups: ["live_test", "mock_test"]}
isolated function testCreateList() returns error? {
    MicrosoftGraphList payload = {
        displayName: "Test SharePoint List",
        list: <MicrosoftGraphListInfo>{template: "genericList"}
    };
    MicrosoftGraphList response = check sharepoint->createList(siteId, payload);
    lock {
        createdListId = response.id ?: "";
    }
    test:assertTrue(response.id !is (), msg = "List creation failed: no ID returned.");
    test:assertEquals(response["displayName"], "Test SharePoint List",
        msg = "List creation failed: unexpected display name.");
}

@test:Config {dependsOn: [testCreateList], groups: ["live_test", "mock_test"]}
isolated function testGetList() returns error? {
    string listId;
    lock {
        listId = createdListId;
    }
    MicrosoftGraphList response = check sharepoint->getList(siteId, listId);
    test:assertEquals(response.id, listId, msg = "Get list failed: unexpected list ID.");
}

@test:Config {groups: ["live_test", "mock_test"]}
isolated function testListLists() returns error? {
    MicrosoftGraphListCollectionResponse response = check sharepoint->listLists(siteId);
    test:assertTrue((response.value ?: []).length() > 0,
        msg = "List lists failed: expected at least one list.");
}

@test:Config {dependsOn: [testCreateList], groups: ["live_test", "mock_test"]}
isolated function testUpdateList() returns error? {
    string listId;
    lock {
        listId = createdListId;
    }
    MicrosoftGraphList updatePayload = {
        displayName: "Updated Test List"
    };
    check sharepoint->updateList(siteId, listId, updatePayload);
}

@test:Config {dependsOn: [testUpdateList, testDeleteItem, testDeleteSubscription], groups: ["live_test", "mock_test"]}
isolated function testDeleteList() returns error? {
    string listId;
    lock {
        listId = createdListId;
    }
    error? response = check sharepoint->deleteList(siteId, listId);
    test:assertEquals(response, ());
}

// ===== Item Tests =====

@test:Config {dependsOn: [testCreateList], groups: ["live_test", "mock_test"]}
isolated function testCreateItem() returns error? {
    string listId;
    lock {
        listId = createdListId;
    }
    MicrosoftGraphListItem payload = {
        fields: <MicrosoftGraphFieldValueSet>{}
    };
    MicrosoftGraphListItem response = check sharepoint->createItem(siteId, listId, payload);
    lock {
        createdItemId = response.id ?: "";
    }
    test:assertTrue(response.id !is (), msg = "Item creation failed: no ID returned.");
}

@test:Config {dependsOn: [testCreateItem], groups: ["live_test", "mock_test"]}
isolated function testGetItem() returns error? {
    string listId;
    string itemId;
    lock {
        listId = createdListId;
    }
    lock {
        itemId = createdItemId;
    }
    MicrosoftGraphListItem response = check sharepoint->getItem(siteId, listId, itemId);
    test:assertEquals(response.id, itemId, msg = "Get item failed: unexpected item ID.");
}

@test:Config {dependsOn: [testCreateList], groups: ["live_test", "mock_test"]}
isolated function testListItems() returns error? {
    string listId;
    lock {
        listId = createdListId;
    }
    MicrosoftGraphListItemCollectionResponse response = check sharepoint->listItems(siteId, listId);
    test:assertTrue((response.value ?: []).length() > 0,
        msg = "List items failed: expected at least one item.");
}

@test:Config {dependsOn: [testCreateItem], groups: ["live_test", "mock_test"]}
isolated function testUpdateItem() returns error? {
    string listId;
    string itemId;
    lock {
        listId = createdListId;
    }
    lock {
        itemId = createdItemId;
    }
    MicrosoftGraphListItem updatePayload = {
        fields: <MicrosoftGraphFieldValueSet>{}
    };
    check sharepoint->updateItem(siteId, listId, itemId, updatePayload);
}

@test:Config {dependsOn: [testUpdateItem], groups: ["live_test", "mock_test"]}
isolated function testDeleteItem() returns error? {
    string listId;
    string itemId;
    lock {
        listId = createdListId;
    }
    lock {
        itemId = createdItemId;
    }
    error? response = check sharepoint->deleteItem(siteId, listId, itemId);
    test:assertEquals(response, ());
}

// ===== Subscription Tests =====

@test:Config {dependsOn: [testCreateList], groups: ["live_test", "mock_test"]}
isolated function testCreateSubscription() returns error? {
    string listId;
    lock {
        listId = createdListId;
    }
    MicrosoftGraphSubscription payload = {
        notificationUrl: subscriptionNotificationUrl,
        expirationDateTime: subscriptionExpiry,
        'resource: string `sites/${siteId}/lists/${listId}/items`,
        changeType: "updated"
    };
    MicrosoftGraphSubscription response = check sharepoint->createSubscription(siteId, listId, payload);
    lock {
        createdSubscriptionId = response.id ?: "";
    }
    test:assertTrue(response.id !is (), msg = "Subscription creation failed: no ID returned.");
    test:assertEquals(response.changeType, "updated",
        msg = "Subscription creation failed: unexpected change type.");
}

@test:Config {dependsOn: [testCreateSubscription], groups: ["live_test", "mock_test"]}
isolated function testGetSubscription() returns error? {
    string listId;
    string subId;
    lock {
        listId = createdListId;
    }
    lock {
        subId = createdSubscriptionId;
    }
    MicrosoftGraphSubscription response = check sharepoint->getSubscription(siteId, listId, subId);
    test:assertEquals(response.id, subId, msg = "Get subscription failed: unexpected subscription ID.");
}

@test:Config {dependsOn: [testCreateList], groups: ["live_test", "mock_test"]}
isolated function testListSubscriptions() returns error? {
    string listId;
    lock {
        listId = createdListId;
    }
    MicrosoftGraphSubscriptionCollectionResponse response = check sharepoint->listSubscriptions(siteId, listId);
    test:assertTrue((response.value ?: []).length() > 0,
        msg = "List subscriptions failed: expected at least one subscription.");
}

@test:Config {dependsOn: [testCreateSubscription], groups: ["live_test", "mock_test"]}
isolated function testUpdateSubscription() returns error? {
    string listId;
    string subId;
    lock {
        listId = createdListId;
    }
    lock {
        subId = createdSubscriptionId;
    }
    MicrosoftGraphSubscription updatePayload = {
        expirationDateTime: updatedSubscriptionExpiry
    };
    check sharepoint->updateSubscription(siteId, listId, subId, updatePayload);
}

@test:Config {dependsOn: [testUpdateSubscription], groups: ["live_test", "mock_test"]}
isolated function testDeleteSubscription() returns error? {
    string listId;
    string subId;
    lock {
        listId = createdListId;
    }
    lock {
        subId = createdSubscriptionId;
    }
    error? response = check sharepoint->deleteSubscription(siteId, listId, subId);
    test:assertEquals(response, ());
}

// ===== Negative Tests =====

@test:Config {groups: ["live_test", "mock_test"]}
isolated function testGetNonExistentList() returns error? {
    MicrosoftGraphList|error response = sharepoint->getList(siteId, "non-existent-id");
    test:assertTrue(response is error, msg = "Expected error for non-existent list ID.");
}

@test:Config {groups: ["live_test", "mock_test"]}
isolated function testCreateListMissingDisplayName() returns error? {
    MicrosoftGraphList payload = {};
    MicrosoftGraphList|error response = sharepoint->createList(siteId, payload);
    test:assertTrue(response is error, msg = "Expected error when creating list without displayName.");
}

@test:Config {dependsOn: [testCreateList], groups: ["live_test", "mock_test"]}
isolated function testDeleteNonExistentSubscription() returns error? {
    string listId;
    lock {
        listId = createdListId;
    }
    error? response = sharepoint->deleteSubscription(siteId, listId, "non-existent-sub");
    test:assertTrue(response is error, msg = "Expected error for deleting non-existent subscription.");
}
