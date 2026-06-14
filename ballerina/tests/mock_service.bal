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

import ballerina/http;

service on new http:Listener(9090) {

    // ===== List operations =====

    resource function get [string siteId]/lists()
        returns MicrosoftGraphListCollectionResponse|error {
        return {
            value: [
                {id: "mock-list-id-001", displayName: "Test SharePoint List"}
            ]
        };
    }

    resource function post [string siteId]/lists(@http:Payload MicrosoftGraphList payload)
        returns MicrosoftGraphList|http:BadRequest {
        if payload["displayName"] is () {
            return http:BAD_REQUEST;
        }
        return {id: "mock-list-id-001", displayName: "Test SharePoint List"};
    }

    resource function get [string siteId]/lists/[string listId]()
        returns MicrosoftGraphList|http:NotFound {
        if listId == "non-existent-id" {
            return http:NOT_FOUND;
        }
        return {id: listId, displayName: "Test SharePoint List"};
    }

    resource function patch [string siteId]/lists/[string listId](
            @http:Payload MicrosoftGraphList payload)
        returns http:NoContent|error {
        return http:NO_CONTENT;
    }

    resource function delete [string siteId]/lists/[string listId]()
        returns http:NoContent|error {
        return http:NO_CONTENT;
    }

    // ===== Item operations =====

    resource function get [string siteId]/lists/[string listId]/items()
        returns MicrosoftGraphListItemCollectionResponse|error {
        return {
            value: [
                {id: "mock-item-id-001"}
            ]
        };
    }

    resource function post [string siteId]/lists/[string listId]/items(
            @http:Payload MicrosoftGraphListItem payload)
        returns MicrosoftGraphListItem|error {
        return {id: "mock-item-id-001"};
    }

    resource function get [string siteId]/lists/[string listId]/items/[string itemId]()
        returns MicrosoftGraphListItem|http:NotFound {
        return {id: itemId};
    }

    resource function patch [string siteId]/lists/[string listId]/items/[string itemId](
            @http:Payload MicrosoftGraphListItem payload)
        returns http:NoContent|error {
        return http:NO_CONTENT;
    }

    resource function delete [string siteId]/lists/[string listId]/items/[string itemId]()
        returns http:NoContent|error {
        return http:NO_CONTENT;
    }

    // ===== Subscription operations =====

    resource function get [string siteId]/lists/[string listId]/subscriptions()
        returns MicrosoftGraphSubscriptionCollectionResponse|error {
        return {
            value: [
                {
                    id: "mock-sub-id-001",
                    notificationUrl: "https://example.com/webhook",
                    expirationDateTime: "2027-01-01T00:00:00Z",
                    'resource: string `sites/${siteId}/lists/${listId}/items`,
                    changeType: "updated"
                }
            ]
        };
    }

    resource function post [string siteId]/lists/[string listId]/subscriptions(
            @http:Payload MicrosoftGraphSubscription payload)
        returns MicrosoftGraphSubscription|http:BadRequest {
        if payload["notificationUrl"] is () {
            return http:BAD_REQUEST;
        }
        return {
            id: "mock-sub-id-001",
            notificationUrl: "https://example.com/webhook",
            expirationDateTime: "2027-01-01T00:00:00Z",
            'resource: string `sites/${siteId}/lists/${listId}/items`,
            changeType: "updated"
        };
    }

    resource function get [string siteId]/lists/[string listId]/subscriptions/[string subscriptionId]()
        returns MicrosoftGraphSubscription|http:NotFound {
        if subscriptionId == "non-existent-sub" {
            return http:NOT_FOUND;
        }
        return {
            id: subscriptionId,
            notificationUrl: "https://example.com/webhook",
            expirationDateTime: "2027-01-01T00:00:00Z",
            changeType: "updated"
        };
    }

    resource function patch [string siteId]/lists/[string listId]/subscriptions/[string subscriptionId](
            @http:Payload MicrosoftGraphSubscription payload)
        returns http:NoContent|error {
        return http:NO_CONTENT;
    }

    resource function delete [string siteId]/lists/[string listId]/subscriptions/[string subscriptionId]()
        returns http:NoContent|http:NotFound {
        if subscriptionId == "non-existent-sub" {
            return http:NOT_FOUND;
        }
        return http:NO_CONTENT;
    }
}
