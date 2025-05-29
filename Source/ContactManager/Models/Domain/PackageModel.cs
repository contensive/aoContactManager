using System;
using System.Collections.Generic;

namespace Contensive.Addons.ContactManager {
    // 
    // ====================================================================================================
    // 
    /// <summary>
/// remote method top level data structure
/// </summary>
    [Serializable()]
    public class PackageModel {
        public bool success { get; set; } = false;
        public List<PackageErrorModel> errorList { get; set; } = new List<PackageErrorModel>();
        public List<PackageNodeModel> nodeList { get; set; } = new List<PackageNodeModel>();
        public List<PackageProfileModel> profileList { get; set; }
    }
}