/// <reference path="d365/window.d.ts" />
/// <reference path="d365/metadata.d.ts" />

interface String {
    /**
     * Trims leading and trailing curly braces from the string.
     * @returns {string} The trimmed string.
     */
    trimCurlyBraces(): string;

    /**
     * Formats the string injecting the arguments like in C# String.Format().
     * Example: let formattedString = "Hello, {0}!".format("World")
     * @param args The arguments.
     */
    format(...args: any[]): string;
}

String.prototype.trimCurlyBraces = function () { return this.replace(/^\{|\}$/g, ""); };

String.prototype.format = function (...args: any[]) {
    return this.replace(/{(\d+)}/g, function (match, number) {
        return typeof args[number] !== 'undefined'
            ? args[number]
            : match
            ;
    });
};

namespace Helper {

    export class Guid {

        public static Empty: string = '00000000-0000-0000-0000-000000000000';

        public static Equals(first: string, second: string): boolean {
            return Guid.format(first) === Guid.format(second);
        }

        public static format(value: string): string {
            if (value === '')
                return value;
            return '{' + value.trimCurlyBraces().toUpperCase() + '}';
        }

        public static newGuid() {
            var newGuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
                var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
                return v.toString(16);
            });
            return Guid.format(newGuid);
        };
    }

    export class Utils {

        public static prepareStatement(Query: string, Attributes: { [index: string]: any }): string {

            if (!Query || !Attributes) {
                return "";
            }

            for (let key of Object.keys(Attributes)) {
                if (Attributes[key] !== undefined && Attributes[key] !== null) {
                    let re = new RegExp("##" + key + "##", "g")
                    Query = Query.replace(re, Attributes[key]);
                }
            }

            return Query;
        }

        public static setCustomLookupView(LookupField: any, EntityName: string, ViewName: string, FieldName: string, FieldId: string, FetchXml: string, CellWidth?: number, ViewId?: string) {

            if (!EntityName || !ViewName || !FieldName || !FieldId || !FetchXml || !LookupField) {
                //console.log("[DEBUG] Wrong parameters"); //for debug purposes
                return;
            }
            CellWidth = (CellWidth || 250);

            ViewId = ViewId || Guid.newGuid();

            var layoutXml = `<grid name="resultset" jump="${FieldName}" select="1" icon="1" preview="1" object="1" >`;
            layoutXml += `<row name="result" id="${FieldId}" >`;
            layoutXml += `<cell name="${FieldName}" width="${CellWidth}" />`;
            layoutXml += '</row></grid>';
            LookupField.addCustomView(ViewId, EntityName, ViewName, FetchXml, layoutXml, true);
        }

        public static async getDescriptionBasedOnType(entityLogicalName: string, formLookupField: any)
        {
            let description: string = '';

            if (formLookupField) {
                const entity = await Helper.CRM.WebApi.retrieveRecord(entityLogicalName, formLookupField[0].id.trimCurlyBraces(), "?$select=moreone_name");
                description = entity.moreone_name;
            }
            return description;
        } 
    }

    export namespace CRM {

        export function getXrm(): D365.Sdk.WindowXrm {
            let xrms: D365.Sdk.WindowXrm[] = [];
            if (typeof (window.Xrm) !== "undefined")
                xrms.push(window.Xrm);
            if (window.parent && window.parent !== window && typeof (window.parent.Xrm) !== "undefined")
                xrms.push(window.parent.Xrm);
            if (window.opener && window.opener !== window && typeof (window.opener.Xrm) !== "undefined")
                xrms.push(window.opener.Xrm);
            if (xrms.length > 0) {
                for (let xrm of xrms) {
                    if (xrm.Page && xrm.Page.data)
                        return xrm;
                }
                return xrms[0];
            }
            throw new Error("Xrm is undefined");
        }

        export function getGlobalContext(): D365.Sdk.GlobalContext {
            return Helper.CRM.getXrm().Utility.getGlobalContext();
        }

        export function getFormContext(): D365.Sdk.FormContext {
            return Helper.CRM.getXrm().Page;
        }

        export function getUtility(): D365.Sdk.XrmUtility {
            return Helper.CRM.getXrm().Utility;
        }

        export async function dissociateRelationship(fromEntity: string, fromEntityId: string, relationshipName: string, toEntity: string, toEntityId: string, pluralEntityName = "") {
            const serverUrl = Helper.CRM.getGlobalContext().getClientUrl() + "/api/data/v9.0/";
            if (pluralEntityName == "")
                fromEntity += "s";
            else
                fromEntity = pluralEntityName;

            const deleteUrl = serverUrl + fromEntity + "(" + fromEntityId.replace("{", "").replace("}", "") + ")" + "/" + relationshipName + "/$ref?$id=" + serverUrl + toEntity + "s(" + toEntityId.replace("{", "").replace("}", "") + ")";
            var rq = new window.XMLHttpRequest();
            if (rq) {
                rq.open("DELETE", deleteUrl, false);
                rq.setRequestHeader("Accept", "application/json");
                rq.send(null);
            }
        }

        export class WebApi {
            public static createRecord(entityLogicalName: string, data: { [key: string]: any }): Promise<D365.Sdk.EntityReference> {
                return WebApi.initialize()
                    .then(() => Helper.CRM.getXrm().WebApi.createRecord(entityLogicalName, data));
            }

            public static deleteRecord(entityLogicalName: string, id: string): Promise<D365.Sdk.EntityReference> {
                return WebApi.initialize()
                    .then(() => Helper.CRM.getXrm().WebApi.deleteRecord(entityLogicalName, id));
            }

            public static retrieveRecord(entityLogicalName: string, id: string, options?: string): Promise<D365.Sdk.Entity> {
                return WebApi.initialize()
                    .then(() => Helper.CRM.getXrm().WebApi.retrieveRecord(entityLogicalName, id, options));
            }

            public static retrieveMultipleRecords(entityLogicalName: string, options?: string, maxPageSize?: number): Promise<D365.Sdk.WebApiRetrieveMultipleRecordsResponse> {
                return WebApi.initialize()
                    .then(() => Helper.CRM.getXrm().WebApi.retrieveMultipleRecords(entityLogicalName, options, maxPageSize));
            }

            public static updateRecord(entityLogicalName: string, id: string, data: { [key: string]: any }): Promise<D365.Sdk.EntityReference> {
                return WebApi.initialize()
                    .then(() => Helper.CRM.getXrm().WebApi.updateRecord(entityLogicalName, id, data));
            }

            public static execute(request: D365.Sdk.WebApiExecuteRequest): Promise<D365.Sdk.WebApiExecuteResponse> {
                return WebApi.initialize()
                    .then(() => Helper.CRM.getXrm().WebApi.execute(request));
            }

            public static executeMultiple(requests: (D365.Sdk.WebApiExecuteRequest | D365.Sdk.WebApiExecuteRequest[])[]): Promise<D365.Sdk.WebApiExecuteResponse[]> {
                return WebApi.initialize()
                    .then(() => Helper.CRM.getXrm().WebApi.executeMultiple(requests));
            }

            private static initialize(): Promise<void> {
                return Promise.resolve().then(() => {
                    // Workaround per bug V9.0: https://stackoverflow.com/questions/51416490/using-xrm-webapi-method-in-web-resource-opened-in-a-new-window
                    if (!(<any>window)["ENTITY_SET_NAMES"]) {
                        if (window.opener) {
                            (<any>window)["ENTITY_SET_NAMES"] = window.opener.top.ENTITY_SET_NAMES;
                        }
                    }
                });
            }
        }
    }
}