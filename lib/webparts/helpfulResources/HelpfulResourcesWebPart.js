"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_1 = require("@pnp/sp");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var sp_lodash_subset_1 = require("@microsoft/sp-lodash-subset");
var HelpfulResourcesWebPart_module_scss_1 = require("./HelpfulResourcesWebPart.module.scss");
var strings = require("HelpfulResourcesWebPartStrings");
var sp_http_1 = require("@microsoft/sp-http");
//global vars
var userDept = "";
var HelpfulResourcesWebPart = (function (_super) {
    __extends(HelpfulResourcesWebPart, _super);
    function HelpfulResourcesWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        // main promoise method, 1st we get the department, second chain is a REST Call to query the list
        // third we get the list data and figure out the document libraries
        _this.getuser = new Promise(function (resolve, reject) {
            // SharePoint PnP Rest Call to get the User Profile Properties
            return sp_1.sp.profiles.myProperties.get().then(function (result) {
                var props = result.UserProfileProperties;
                var propValue = "";
                var userDepartment = "";
                props.forEach(function (prop) {
                    //this call returns key/value pairs so we need to look for the Dept Key
                    if (prop.Key == "Department") {
                        // set our global var for the users Dept.
                        userDept += prop.Value;
                    }
                });
                return result;
            }).then(function (result) {
                _this._getListData().then(function (response) {
                    _this._renderList(response.value);
                });
            });
        });
        return _this;
    }
    HelpfulResourcesWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n      <div class=\"" + HelpfulResourcesWebPart_module_scss_1.default.helpfulResources + "\">\n        <div class=\"" + HelpfulResourcesWebPart_module_scss_1.default.container + "\">\n          <div class=\"" + HelpfulResourcesWebPart_module_scss_1.default.row + "\">\n            <div class=\"" + HelpfulResourcesWebPart_module_scss_1.default.column + "\">\n              <span class=\"" + HelpfulResourcesWebPart_module_scss_1.default.title + "\">Welcome to SharePoint!</span>\n              <p class=\"" + HelpfulResourcesWebPart_module_scss_1.default.subTitle + "\">Customize SharePoint experiences using Web Parts.</p>\n              <p class=\"" + HelpfulResourcesWebPart_module_scss_1.default.description + "\">" + sp_lodash_subset_1.escape(this.properties.description) + "</p>\n              <a href=\"https://aka.ms/spfx\" class=\"" + HelpfulResourcesWebPart_module_scss_1.default.button + "\">\n                <span class=\"" + HelpfulResourcesWebPart_module_scss_1.default.label + "\">Learn more</span>\n              </a>\n            </div>\n            <h1>Helpful Resources</h1>\n            <h3><div id=\"HelpfulRes\"/></h3>\n          </div>\n        </div>\n      </div>";
    };
    Object.defineProperty(HelpfulResourcesWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    // main REST Call to the list...passing in the deaprtment into the call to 
    //return a single list item
    HelpfulResourcesWebPart.prototype._getListData = function () {
        return this.context.spHttpClient.get("https://girlscoutsrv.sharepoint.com/_api/web/lists/GetByTitle('TeamDashboardSettings')/Items?$filter=Title eq '" + userDept + "'", sp_http_1.SPHttpClient.configurations.v1)
            .then(function (response) {
            return response.json();
        });
    };
    HelpfulResourcesWebPart.prototype._renderList = function (items) {
        var html = '';
        var libHTML = '';
        var siteURL = "";
        //list name
        var helpfulResources = "";
        // items in the list
        var hrItems = "";
        items.forEach(function (item) {
            siteURL = item.DeptURL;
            helpfulResources = item.a85u;
        });
        //1st we need to override the current web to go to the department sites web
        var w = new sp_1.Web("https://girlscoutsrv.sharepoint.com" + siteURL);
        // then use PnP to query the list
        w.lists.getByTitle(helpfulResources).items
            .get()
            .then(function (data) {
            console.log(data);
            for (var x = 0; x < data.length; x++) {
                //console.log(data[x].URL);
                console.log(data[x].URL.Url);
                console.log(data[x].URL.Description);
                //hrItems += data[x].URL + '\r\n';
                // libHTML += `<p>${hrItems.toString()}</p>`;
            }
            //document.getElementById("HelpfulRes").innerText = hrItems;
        }).catch(function (e) { console.error(e); });
        var listContainer = this.domElement.querySelector('#ListItems');
        listContainer.innerHTML = html;
    };
    // this is required to use the SharePoint PnP shorthand REST CALLS
    HelpfulResourcesWebPart.prototype.onInit = function () {
        var _this = this;
        return _super.prototype.onInit.call(this).then(function (_) {
            sp_1.sp.setup({
                spfxContext: _this.context
            });
        });
    };
    HelpfulResourcesWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return HelpfulResourcesWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = HelpfulResourcesWebPart;

//# sourceMappingURL=HelpfulResourcesWebPart.js.map
