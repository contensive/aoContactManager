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
                    request.FormID = request.FormID != Constants.FormIdEnum.FormUnknown ? request.FormID : getDefaultFormId(request);
                    if (!string.IsNullOrEmpty(request.Button)) {
                        // 
                        // ----- Process Previous Forms
                        // 
                        switch (request.FormID) {
                            case Constants.FormIdEnum.FormSearch: {
                                    // 
                                    request.FormID = SearchView.processRequest(cp, ae, request);
                                    break;
                                }
                            case Constants.FormIdEnum.FormList: {
                                    // 
                                    request.FormID = ListView.processRequest(cp, ae, request);
                                    break;
                                }
                            case Constants.FormIdEnum.FormDetails: {
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
                        case Constants.FormIdEnum.FormDetails: {
                                // 
                                result += DetailView.getResponse(cp, ae, request.DetailMemberID);
                                break;
                            }
                        case Constants.FormIdEnum.FormList: {
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
        private Constants.FormIdEnum getDefaultFormId(RequestModel request) {
            if (request.DetailMemberID != 0)
                return Constants.FormIdEnum.FormDetails;
            return Constants.FormIdEnum.FormList;
            // If ((ae.userProperties.contactSearchCriteria <> "") Or (ae.userProperties.contactGroupCriteria <> "")) Then Return FormIdEnum.FormList
            // Return FormIdEnum.FormSearch
        }
    }
}