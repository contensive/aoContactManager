# Plan: Move Contact Manager Addon into aoCRM Collection

## Overview

Move the Contact Manager addon source code out of the standalone `aoContactManager` collection and into the `aoCRM` collection so it ships as a native CRM addon instead of an imported dependency.

## Current State

- **aoContactManager** is a standalone collection (GUID `{51a67cb6-ea94-497c-896a-979a851be3db}`)
- **aoCRM** imports it via `<ImportCollection Name="Contact Manager">` in its collection.xml
- aoCRM already registers a Portal Feature record (`{375f0cd2-1dd7-442f-a8f6-2755d1880dd9}`) that points to the Contact Manager addon GUID (`{57602963-CA46-41EB-8053-8ACB8C47C2F2}`) under Tools

## Inventory of Files to Move

### C# Source (18 files, ~3,800 lines)

| File | Target Location in aoCRM |
|------|--------------------------|
| `Addons/ContactManagerAddon.cs` | `server/aoCrm/Addons/ContactManagerAddon.cs` |
| `Constants.cs` | merge into `server/aoCrm/constants.cs` |
| `Controllers/ApplicationController.cs` | refactor (see collisions below) |
| `Controllers/AdminUIController.cs` | `server/aoCrm/Controllers/AdminUIController.cs` |
| `Controllers/GenericController.cs` | merge into existing `server/aoCrm/Controllers/GenericController.cs` |
| `Controllers/HtmlController.cs` | `server/aoCrm/Controllers/HtmlController.cs` |
| `Controllers/TabController.cs` | `server/aoCrm/Controllers/TabController.cs` |
| `Controllers/DateController.cs` | merge into existing `server/aoCrm/Controllers/DateController.cs` |
| `Views/SearchView.cs` | `server/aoCrm/Views/ContactManager_SearchView.cs` |
| `Views/ListView.cs` | `server/aoCrm/Views/ContactManager_ListView.cs` |
| `Views/DetailView.cs` | `server/aoCrm/Views/ContactManager_DetailView.cs` |
| `Models/Domain/baseDomainModel.cs` | `server/aoCrm/Models/Domain/baseDomainModel.cs` |
| `Models/Domain/RequestModel.cs` | `server/aoCrm/Models/Domain/ContactManager_RequestModel.cs` |
| `Models/Domain/UserPropertiesModel.cs` | `server/aoCrm/Models/Domain/ContactManager_UserPropertiesModel.cs` |
| `Models/Domain/FieldMetaModel.cs` | `server/aoCrm/Models/Domain/FieldMetaModel.cs` |
| `Models/Domain/PackageModel.cs` | `server/aoCrm/Models/Domain/PackageModel.cs` |
| `Models/Domain/PackageErrorModel.cs` | `server/aoCrm/Models/Domain/PackageErrorModel.cs` |
| `Models/Domain/PackageNodeModel.cs` | `server/aoCrm/Models/Domain/PackageNodeModel.cs` |
| `Models/Domain/PackageProfileModel.cs` | `server/aoCrm/Models/Domain/PackageProfileModel.cs` |

### UI Files (2 files)

| File | Target Location in aoCRM |
|------|--------------------------|
| `ui/wwwFiles/ContactManager/ContactManager.css` | `ui/wwwFiles/ContactManager/ContactManager.css` |
| `ui/wwwFiles/ContactManager/ContactManager.js` | `ui/wwwFiles/ContactManager/ContactManager.js` |

### NuGet Dependencies

No new packages needed. Contact Manager only uses `Contensive.CPBaseClass` and `Contensive.DbModels`, both already referenced by aoCRM at equal or newer versions.

## Steps

### Step 1 - Namespace Refactoring

Rename all `Contensive.Addons.ContactManager` references to `Contensive.CRM` across all 18 source files.

- `Contensive.Addons.ContactManager` -> `Contensive.CRM`
- `Contensive.Addons.ContactManager.Views` -> `Contensive.CRM`
- `Contensive.Addons.ContactManager.Controllers` -> `Contensive.CRM.Controllers`
- `Contensive.Addons.ContactManager.Models.Domain` -> `Contensive.CRM.Models`

Update the `DotNetClass` in collection.xml:
- `Contensive.Addons.ContactManager.Views.ContactManagerAddon` -> `Contensive.CRM.ContactManagerAddon`

### Step 2 - Resolve Naming Collisions

Both repos have overlapping class names that must be merged or deduplicated:

#### ApplicationController (Medium)
- **Contact Manager version**: Tracks `packageErrorList`, `statusMessage`, lazy-loads `userProperties`, checks auth
- **aoCRM version**: Wraps `ApplicationModel` as a DI container holding `CPBaseClass` and current `PersonModel`
- **Resolution**: Refactor Contact Manager code to use aoCRM's `ApplicationModel` pattern. Move error/status tracking into the Contact Manager views or a dedicated `ContactManagerContext` helper.

#### GenericController (Low)
- **Contact Manager version**: `modifyLinkQuery`, request encoding, URL helpers
- **aoCRM version**: Different utility methods
- **Resolution**: Merge Contact Manager methods into aoCRM's `GenericController`, prefixing if needed to avoid ambiguity.

#### DateController (Low)
- **Contact Manager version**: Single `IsDate()` method
- **aoCRM version**: Different date utilities
- **Resolution**: Add `IsDate()` to aoCRM's `DateController` if not already present.

#### Constants (Low)
- **Contact Manager version**: Field type constants, form/error/HTTP enums, button names, addon GUIDs
- **aoCRM version**: CRM-specific GUIDs, property names, field names
- **Resolution**: Merge Contact Manager constants into aoCRM's `constants.cs`. Move Contact Manager enums into a nested class or region to keep them organized (e.g., `Constants.ContactManager.FormIdEnum`).

### Step 3 - Copy Source Files into aoCRM

Copy the C# files listed in the inventory above into the aoCRM project folders. Apply namespace changes from Step 1.

### Step 4 - Copy UI Files into aoCRM

Copy `ContactManager.css` and `ContactManager.js` into `aoCRM/ui/wwwFiles/ContactManager/`.

### Step 5 - Update aoCRM collection.xml

1. **Move the addon registration** from `aoContactManager.xml` into `aoCrm.xml`:
   ```xml
   <Addon Name="Contact Manager" Guid="{57602963-CA46-41EB-8053-8ACB8C47C2F2}" Type="Tool">
       <DotNetClass><![CDATA[Contensive.CRM.ContactManagerAddon]]></DotNetClass>
       <BlockEditTools>Yes</BlockEditTools>
       <IncludeAddon Name="Bootstrap" Guid="{EF1FD66C-D62F-4BD2-BF07-38F47996EBB3}"/>
       <JSHeadScriptSrc><![CDATA[/ContactManager/ContactManager.js]]></JSHeadScriptSrc>
       <styleslinkhref><![CDATA[/ContactManager/ContactManager.css]]></styleslinkhref>
   </Addon>
   ```

2. **Add a `<Resource>` entry** for the ContactManager wwwFiles if not already covered by aoCRM's existing wwwFiles.zip packaging.

3. **Remove the import**: Delete `<ImportCollection Name="Contact Manager">{51a67cb6-ea94-497c-896a-979a851be3db}</ImportCollection>` from aoCrm.xml.

4. **Keep the existing Portal Feature record** (`{375f0cd2-1dd7-442f-a8f6-2755d1880dd9}`) - it already correctly references the addon GUID and Tools parent.

### Step 6 - Update aoCRM Build Scripts

Verify that the aoCRM build scripts (`build.ps1`) package the new `ContactManager/` wwwFiles subfolder into the wwwFiles.zip.

### Step 7 - Build and Test

1. Build the aoCRM solution - verify no compile errors
2. Install the updated aoCRM collection on a test site
3. Verify Contact Manager appears in the CRM portal under Tools
4. Test all three views: Search, List, Detail
5. Test CSV export functionality (uses external addon GUID `{5C25F35D-A2A8-4791-B510-B1FFE0645004}`)
6. Test group management operations
7. Verify portal redirect works when addon is accessed directly outside portal

### Step 8 - Retire aoContactManager

1. Remove the `<ImportCollection>` from aoCrm.xml (done in Step 5)
2. Archive the aoContactManager repository (mark as deprecated/archived)
3. Remove aoContactManager from any deployment pipelines
4. Uninstall the standalone Contact Manager collection from sites that have the updated aoCRM

## Risk Assessment

| Risk | Severity | Mitigation |
|------|----------|------------|
| Namespace change breaks DotNetClass lookup | High | Update collection.xml DotNetClass in same deploy |
| ApplicationController refactor introduces bugs | Medium | Test all three views thoroughly after refactor |
| Duplicate utility methods cause ambiguity | Low | Compiler will catch; resolve at build time |
| Sites with standalone Contact Manager get duplicate | Medium | Uninstall standalone collection before or after upgrading aoCRM |
| CSV Export addon dependency not met | Low | CSV Export is a separate collection; ensure it remains installed |
