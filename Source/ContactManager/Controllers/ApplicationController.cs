using System;
using System.Collections.Generic;
using Contensive.BaseClasses;

namespace Contensive.Addons.ContactManager.Controllers {
    // 
    // ====================================================================================================
    /// <summary>
    /// 
    /// </summary>
    /// <remarks></remarks>
    public class ApplicationController : IDisposable {
        // 
        private readonly CPBaseClass cp;
        // 
        // ====================================================================================================
        /// <summary>
        /// Errors accumulated during rendering.
        /// </summary>
        /// <returns></returns>
        public List<PackageErrorModel> packageErrorList { get; set; } = new List<PackageErrorModel>();
        // '
        // '====================================================================================================
        // ''' <summary>
        // ''' data accumulated during rendering
        // ''' </summary>
        // ''' <returns></returns>
        // Public Property packageNodeList As New List(Of PackageNodeModel)
        // '
        // '====================================================================================================
        // ''' <summary>
        // ''' list of name/time used to performance analysis
        // ''' </summary>
        // ''' <returns></returns>
        // Public Property packageProfileList As New List(Of PackageProfileModel)
        // 
        // ====================================================================================================
        /// <summary>
        /// status message displayed on get-form
        /// </summary>
        /// <returns></returns>
        public string statusMessage {
            get {
                return _StatusMessage;
            }
            set {
                _StatusMessage = value;
            }
        }
        private string _StatusMessage = "";
        // 
        // ====================================================================================================
        /// <summary>
        /// application's user properties
        /// </summary>
        /// <returns></returns>
        public Models.Domain.UserPropertiesModel userProperties {
            get {
                if (_userProperties is null) {
                    _userProperties = new Models.Domain.UserPropertiesModel(cp);
                }
                return _userProperties;
            }
        }
        private Models.Domain.UserPropertiesModel _userProperties = null;
        // '
        // '====================================================================================================
        // ''' <summary>
        // ''' get the serialized results
        // ''' </summary>
        // ''' <returns></returns>
        // Public Function getSerializedPackage() As String
        // Try
        // Return cp.JSON.Serialize(New PackageModel With {
        // .success = packageErrorList.Count.Equals(0),
        // .nodeList = packageNodeList,
        // .errorList = packageErrorList,
        // .profileList = packageProfileList
        // })
        // Catch ex As Exception
        // cp.Site.ErrorReport(ex)
        // Throw
        // End Try
        // End Function
        // 
        // ====================================================================================================
        /// <summary>
        /// Constructor
        /// </summary>
        /// <remarks></remarks>
        public ApplicationController(CPBaseClass cp, bool requiresAuthentication = true) {
            this.cp = cp;
            if (requiresAuthentication & !cp.User.IsAuthenticated) {
                packageErrorList.Add(new PackageErrorModel() { number = (int)constants.ResultErrorEnum.errAuthentication, description = "Authorization is required." });
                cp.Response.SetStatus(((int)constants.HttpErrorEnum.forbidden).ToString() + " Forbidden");
            }
        }
        // 
        #region  IDisposable Support 
        protected bool disposed = false;
        // 
        // ==========================================================================================
        /// <summary>
        /// dispose
        /// </summary>
        /// <param name="disposing"></param>
        /// <remarks></remarks>
        protected virtual void Dispose(bool disposing) {
            if (!disposed) {
                if (disposing) {
                    // 
                    // ----- call .dispose for managed objects
                    // 
                }
                // 
                // Add code here to release the unmanaged resource.
                // 
            }
            disposed = true;
        }
        // Do not change or add Overridable to these methods.
        // Put cleanup code in Dispose(ByVal disposing As Boolean).
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        ~ApplicationController() {
            Dispose(false);
        }
        #endregion
    }
}