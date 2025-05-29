using System;

namespace Contensive.Addons.ContactManager {

    // 
    // ====================================================================================================
    /// <summary>
/// error list for jsonPackage
/// </summary>
    [Serializable()]
    public class PackageErrorModel {
        public int number { get; set; } = 0;
        public string description { get; set; } = "";
    }
}