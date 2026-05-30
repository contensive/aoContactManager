
using Contensive.BaseClasses;

namespace Contensive.Addons.ContactManager {
    /// <summary>
/// requests processed by this form
/// </summary>
    public class RequestModel {
        public int TabNumber { get; set; }
        public string Button { get; set; }
        public int DetailMemberID { get; set; }
        public Constants.FormIdEnum FormID { get; set; }
        public int RowCount { get; set; }
        public int GroupID { get; set; }
        public Constants.GroupToolActionEnum GroupToolAction { get; set; }
        public int GroupToolSelect { get; set; }
        public string SelectionGroupSubTab { get; set; }
        public string SelectionSearchSubTab { get; set; }
        // 
        public RequestModel(CPBaseClass cp) {
            // 
            // todo - convert request names to Constants
            // 
            TabNumber = cp.Doc.GetInteger("tab");
            Button = cp.Doc.GetText("Button");
            DetailMemberID = cp.Doc.GetInteger(Constants.RequestNameMemberID);
            RowCount = cp.Doc.GetInteger("M.Count");
            GroupID = cp.Doc.GetInteger("GroupID");
            GroupToolSelect = cp.Doc.GetInteger("GroupToolSelect");
            SelectionGroupSubTab = cp.Doc.GetText("SelectionGroupSubTab");
            SelectionSearchSubTab = cp.Doc.GetText("SelectionSearchSubTab");
            // 
            int testAction = cp.Doc.GetInteger("GroupToolAction");
            GroupToolAction = testAction > 0 & testAction < 5 ? (Constants.GroupToolActionEnum)testAction : Constants.GroupToolActionEnum.nop;
            // 
            int testFormId = cp.Doc.GetInteger(Constants.RequestNameFormID);
            FormID = testFormId > 0 & testFormId < 4 ? (Constants.FormIdEnum)testFormId : Constants.FormIdEnum.FormUnknown;

        }
    }
}