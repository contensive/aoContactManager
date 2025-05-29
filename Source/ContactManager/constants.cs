
namespace Contensive.Addons.ContactManager {




    public static class constants {
        // 
        // 
        public const int Version = 1;
        // 
        public const string addonGuidExportCSV = "{5C25F35D-A2A8-4791-B510-B1FFE0645004}";
        // 
        // -- sample
        public const string rnInputValue = "inputValue";
        public const string RequestNameMemberID = "memberId";
        public const string RequestNameFormID = "formId";
        public const string RequestNamePageSize = "pageSize";
        public const string RequestNamePageNumber = "pageNumber";
        public const string RequestNameDetailSubtab = "subtab";
        // 
        public const string ButtonSearch = "Search";
        public const string ButtonCancel = "Cancel";
        public const string ButtonCancelAll = "Cancel All";
        public const string ButtonNewSearch = "New Search";
        public const string ButtonApply = "Apply";
        public const string ButtonSave = "Save";
        public const string ButtonOK = "OK";
        // 
        public const int FieldTypeInteger = 1;     // An Long number
        public const int FieldTypeText = 2;     // A text field (up To 255 characters)
        public const int FieldTypeLongText = 3;     // A memo field (up To 8000 characters)
        public const int FieldTypeBoolean = 4;     // A yes/no field
        public const int FieldTypeDate = 5;     // A Date field
        public const int FieldTypeFile = 6;     // A filename Of a file In the files directory.
        public const int FieldTypeLookup = 7;     // A lookup Is a FieldTypeInteger that indexes into another table
        public const int FieldTypeRedirect = 8;     // creates a link To another section
        public const int FieldTypeCurrency = 9;     // A Float that prints In dollars
        public const int FieldTypeFileText = 10;     // Text saved In a file In the files area.
        public const int FieldTypeFileImage = 11;     // A filename Of a file In the files directory.
        public const int FieldTypeFloat = 12;     // A float number
        public const int FieldTypeAutoIdIncrement = 13;     // Long that automatically increments With the New record
        public const int FieldTypeManyToMany = 14;     // no database field - sets up a relationship through a Rule table To another table
        public const int FieldTypeMemberSelect = 15;     // This ID Is a ccMembers record In a group defined by the MemberSelectGroupID field
        public const int FieldTypeFileCSS = 16;     // A filename Of a CSS compatible file
        public const int FieldTypeFileXML = 17;     // the filename Of an XML compatible file
        public const int FieldTypeFileJavascript = 18;     // the filename Of a javascript compatible file
        public const int FieldTypeLink = 19;     // Links used In href tags -- can go To pages Or resources
        public const int FieldTypeResourceLink = 20;     // Links used In resources, link <img Or <Object. Should Not be pages
        public const int FieldTypeHTML = 21;     // LongText field that expects HTML content
        public const int FieldTypeFileHTML = 22;     // TextFile field that expects HTML content
        public const int FieldTypeMax = 22;             // 
                                                        // 
        public const int ReportFormVisitList = 9999;
        // 

        // 
        // 
        // -- errors for resultErrList
        public enum ResultErrorEnum {
            errPermission = 50,
            errDuplicate = 100,
            errVerification = 110,
            errRestriction = 120,
            errInput = 200,
            errAuthentication = 300,
            errAdd = 400,
            errSave = 500,
            errDelete = 600,
            errLookup = 700,
            errLoad = 710,
            errContent = 800,
            errMiscellaneous = 900
        }
        // 
        // -- http errors
        public enum HttpErrorEnum {
            badRequest = 400,
            unauthorized = 401,
            paymentRequired = 402,
            forbidden = 403,
            notFound = 404,
            methodNotAllowed = 405,
            notAcceptable = 406,
            proxyAuthenticationRequired = 407,
            requestTimeout = 408,
            conflict = 409,
            gone = 410,
            lengthRequired = 411,
            preconditionFailed = 412,
            payloadTooLarge = 413,
            urlTooLong = 414,
            unsupportedMediaType = 415,
            rangeNotSatisfiable = 416,
            expectationFailed = 417,
            teapot = 418,
            methodFailure = 420,
            enhanceYourCalm = 420,
            misdirectedRequest = 421,
            unprocessableEntity = 422,
            locked = 423,
            failedDependency = 424,
            upgradeRequired = 426,
            preconditionRequired = 428,
            tooManyRequests = 429,
            requestHeaderFieldsTooLarge = 431,
            loginTimeout = 440,
            noResponse = 444,
            retryWith = 449,
            redirect = 451,
            unavailableForLegalReasons = 451,
            sslCertificateError = 495,
            sslCertificateRequired = 496,
            httpRequestSentToSecurePort = 497,
            invalidToken = 498,
            clientClosedRequest = 499,
            tokenRequired = 499,
            internalServerError = 500
        }
        // 
        // ========================================================================
        /// <summary>
    /// possible formid values
    /// </summary>
        public enum FormIdEnum {
            FormUnknown = 0,
            FormDetails = 1,
            FormList = 2,
            FormSearch = 3
        }
        // 
        // ========================================================================
        /// <summary>
    /// options for the group tool
    /// </summary>
        public enum GroupToolActionEnum {
            nop = 0,
            AddToGroup = 1,
            RemoveFromGroup = 2,
            ExportGroup = 3,
            SetGroupEmail = 4
        }
    }
}