
namespace Contensive.Addons.ContactManager {

    // 
    // ========================================================================
    /// <summary>
/// meta data for each row of the people filter
/// </summary>
    public class FieldMeta {
        public string fieldName { get; set; }
        public string fieldCaption { get; set; }
        public int fieldId { get; set; }
        public int fieldType { get; set; }
        public string currentValue { get; set; }
        public int fieldOperator { get; set; }
        public string fieldLookupContentName { get; set; }
        public string fieldLookupList { get; set; }
        public string fieldEditTab { get; set; }
    }
}