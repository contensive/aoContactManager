using System;
using System.Diagnostics;

using Contensive.BaseClasses;

namespace Contensive.Addons.ContactManager.Views {
    public class ContactManagerAddon : AddonBaseClass {
        // 
        // =====================================================================================
        /// <summary>
        /// AddonDescription
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        public override object Execute(CPBaseClass cp) {
            string result = "";
            var sw = new Stopwatch();
            sw.Start();
            try {
                // 
                // -- initialize application. If authentication needed and not login page, pass true
                using (var ae = new Controllers.ApplicationController(cp, false)) {
                    var request = new RequestModel(cp);
                    cp.Doc.AddRefreshQueryString("tab", request.TabNumber.ToString());
                    // 
                    request.FormID = request.FormID != constants.FormIdEnum.FormUnknown ? request.FormID : getDefaultFormId(request);
                    if (!string.IsNullOrEmpty(request.Button)) {
                        // 
                        // ----- Process Previous Forms
                        // 
                        switch (request.FormID) {
                            case constants.FormIdEnum.FormSearch: {
                                    // 
                                    request.FormID = SearchView.processRequest(cp, ae, request);
                                    break;
                                }
                            case constants.FormIdEnum.FormList: {
                                    // 
                                    request.FormID = ListView.processRequest(cp, ae, request);
                                    break;
                                }
                            case constants.FormIdEnum.FormDetails: {
                                    // 
                                    request.FormID = DetailView.processRequest(cp, request);
                                    break;
                                }
                        }
                    }
                    bool IsAdminPath = true;
                    // 
                    // ----- Output the next form
                    switch (request.FormID) {
                        case constants.FormIdEnum.FormDetails: {
                                // 
                                result += DetailView.getResponse(cp, ae, request.DetailMemberID);
                                break;
                            }
                        case constants.FormIdEnum.FormList: {
                                // 
                                result += ListView.getResponse(cp, ae);
                                break;
                            }

                        default: {
                                // 
                                result += SearchView.getResponse(cp, ae, IsAdminPath);
                                break;
                            }
                    }
                    // result = "<div class=ccbodyadmin>" & result & "</div>"
                    // 
                    // wrapper for style strength
                    result = "<div class=\"contactManager\">" + result + "</div>";
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
            return result;
        }
        // 
        // ========================================================================
        /// <summary>
        /// return the default formId
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        private constants.FormIdEnum getDefaultFormId(RequestModel request) {
            if (request.DetailMemberID != 0)
                return constants.FormIdEnum.FormDetails;
            return constants.FormIdEnum.FormList;
            // If ((ae.userProperties.contactSearchCriteria <> "") Or (ae.userProperties.contactGroupCriteria <> "")) Then Return FormIdEnum.FormList
            // Return FormIdEnum.FormSearch
        }
    }
}