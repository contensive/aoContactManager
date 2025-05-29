using Contensive.BaseClasses;

namespace Contensive.Addons.ContactManager.Models.Domain {
    /// <summary>
    /// Expose public properties for all user properties created for this application.
    /// </summary>
    public class UserPropertiesModel {
        // 
        private readonly CPBaseClass cp;
        // 
        // 
        // ====================================================================================================
        /// <summary>
        /// constructor
        /// </summary>
        /// <param name="cp"></param>
        public UserPropertiesModel(CPBaseClass cp) {
            this.cp = cp;
        }
        // 
        // ====================================================================================================
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string contactSearchCriteria {
            get {
                return cp.TempFiles.Read(@"contactManager\user-contact-search-criteria.txt");
            }
            set {
                cp.TempFiles.Save(@"contactManager\user-contact-search-criteria.txt", value);
            }
        }
        // 
        // ====================================================================================================
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string contactGroupCriteria {
            get {
                return cp.TempFiles.Read(@"contactManager\user-group-search-criteria.txt");
            }
            set {
                cp.TempFiles.Save(@"contactManager\user-group-search-criteria.txt", value);
            }
        }
        // 
        // ====================================================================================================
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public int selectSubTab {
            get {
                return cp.User.GetInteger(nameSelectSubTab);
            }
            set {
                cp.User.SetProperty(nameSelectSubTab, value);
            }
        }
        private const string nameSelectSubTab = "SelectSubTab";
        // 
        // ====================================================================================================
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public int subTab {
            get {
                return cp.User.GetInteger(nameSubTab, 1);
            }
            set {
                cp.User.SetProperty(nameSubTab, value);
            }
        }
        private const string nameSubTab = "SubTab";
        // 
        // ====================================================================================================
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public int contactContentID {
            get {
                return cp.User.GetInteger(nameContactContentID, cp.Content.GetID("people"));
            }
            set {
                cp.User.SetProperty(nameContactContentID, value);
            }
        }
        private const string nameContactContentID = "ContactContentID";
        // 
        // 
        // 
    }
}