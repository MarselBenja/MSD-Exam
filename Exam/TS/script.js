"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
String.prototype.trimCurlyBraces = function () { return this.replace(/^\{|\}$/g, ""); };
String.prototype.format = function (...args) {
    return this.replace(/{(\d+)}/g, function (match, number) {
        return typeof args[number] !== 'undefined'
            ? args[number]
            : match;
    });
};
var Helper;
(function (Helper) {
    class Guid {
        static Equals(first, second) {
            return Guid.format(first) === Guid.format(second);
        }
        static format(value) {
            if (value === '')
                return value;
            return '{' + value.trimCurlyBraces().toUpperCase() + '}';
        }
        static newGuid() {
            var newGuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
                var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
                return v.toString(16);
            });
            return Guid.format(newGuid);
        }
        ;
    }
    Guid.Empty = '00000000-0000-0000-0000-000000000000';
    Helper.Guid = Guid;
    class Utils {
        static prepareStatement(Query, Attributes) {
            if (!Query || !Attributes) {
                return "";
            }
            for (let key of Object.keys(Attributes)) {
                if (Attributes[key] !== undefined && Attributes[key] !== null) {
                    let re = new RegExp("##" + key + "##", "g");
                    Query = Query.replace(re, Attributes[key]);
                }
            }
            return Query;
        }
        static setCustomLookupView(LookupField, EntityName, ViewName, FieldName, FieldId, FetchXml, CellWidth, ViewId) {
            if (!EntityName || !ViewName || !FieldName || !FieldId || !FetchXml || !LookupField) {
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
        static getDescriptionBasedOnType(entityLogicalName, formLookupField) {
            return __awaiter(this, void 0, void 0, function* () {
                let description = '';
                if (formLookupField) {
                    const entity = yield Helper.CRM.WebApi.retrieveRecord(entityLogicalName, formLookupField[0].id.trimCurlyBraces(), "?$select=moreone_name");
                    description = entity.moreone_name;
                }
                return description;
            });
        }
    }
    Helper.Utils = Utils;
    let CRM;
    (function (CRM) {
        function getXrm() {
            let xrms = [];
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
        CRM.getXrm = getXrm;
        function getGlobalContext() {
            return Helper.CRM.getXrm().Utility.getGlobalContext();
        }
        CRM.getGlobalContext = getGlobalContext;
        function getFormContext() {
            return Helper.CRM.getXrm().Page;
        }
        CRM.getFormContext = getFormContext;
        function getUtility() {
            return Helper.CRM.getXrm().Utility;
        }
        CRM.getUtility = getUtility;
        function dissociateRelationship(fromEntity, fromEntityId, relationshipName, toEntity, toEntityId, pluralEntityName = "") {
            return __awaiter(this, void 0, void 0, function* () {
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
            });
        }
        CRM.dissociateRelationship = dissociateRelationship;
        class WebApi {
            static createRecord(entityLogicalName, data) {
                return WebApi.initialize()
                    .then(() => Helper.CRM.getXrm().WebApi.createRecord(entityLogicalName, data));
            }
            static deleteRecord(entityLogicalName, id) {
                return WebApi.initialize()
                    .then(() => Helper.CRM.getXrm().WebApi.deleteRecord(entityLogicalName, id));
            }
            static retrieveRecord(entityLogicalName, id, options) {
                return WebApi.initialize()
                    .then(() => Helper.CRM.getXrm().WebApi.retrieveRecord(entityLogicalName, id, options));
            }
            static retrieveMultipleRecords(entityLogicalName, options, maxPageSize) {
                return WebApi.initialize()
                    .then(() => Helper.CRM.getXrm().WebApi.retrieveMultipleRecords(entityLogicalName, options, maxPageSize));
            }
            static updateRecord(entityLogicalName, id, data) {
                return WebApi.initialize()
                    .then(() => Helper.CRM.getXrm().WebApi.updateRecord(entityLogicalName, id, data));
            }
            static execute(request) {
                return WebApi.initialize()
                    .then(() => Helper.CRM.getXrm().WebApi.execute(request));
            }
            static executeMultiple(requests) {
                return WebApi.initialize()
                    .then(() => Helper.CRM.getXrm().WebApi.executeMultiple(requests));
            }
            static initialize() {
                return Promise.resolve().then(() => {
                    if (!window["ENTITY_SET_NAMES"]) {
                        if (window.opener) {
                            window["ENTITY_SET_NAMES"] = window.opener.top.ENTITY_SET_NAMES;
                        }
                    }
                });
            }
        }
        CRM.WebApi = WebApi;
    })(CRM = Helper.CRM || (Helper.CRM = {}));
})(Helper || (Helper = {}));
var Exam;
(function (Exam) {
    var TS;
    (function (TS) {
        let Form;
        function OnLoad(executionContext) {
            return __awaiter(this, void 0, void 0, function* () {
                Form = executionContext.getFormContext();
                TeamOnChange();
            });
        }
        TS.OnLoad = OnLoad;
        function TeamOnChange() {
            return __awaiter(this, void 0, void 0, function* () {
                let Team = Form.getControl("new_salesteamid").getAttribute().getValue();
                if (Team) {
                    if (Team[0].entityType == "cr260_skill") {
                        const confirmStrings = {
                            text: "You can not select a skill for this field. Please select a record from Teams Entity!",
                            title: `Warning!`,
                        };
                        const confirmOptions = {
                            height: 200,
                            width: 400,
                        };
                        Helper.CRM.getXrm().Navigation.openAlertDialog(confirmStrings, confirmOptions);
                    }
                }
            });
        }
        TS.TeamOnChange = TeamOnChange;
    })(TS = Exam.TS || (Exam.TS = {}));
})(Exam || (Exam = {}));
