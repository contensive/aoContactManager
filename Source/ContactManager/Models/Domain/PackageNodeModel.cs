using System;

namespace Contensive.Addons.ContactManager {

    // 
    // ====================================================================================================
    /// <summary>
/// data store for jsonPackage
/// </summary>
    [Serializable()]
    public class PackageNodeModel {
        public string dataFor { get; set; } = "";
        public object data { get; set; }
    }
}