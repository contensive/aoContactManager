
Option Explicit On
Option Strict On

Public Module constants
    '
    ' -- sample
    Public Const Version As Integer = 1
    '
    ' -- sample
    Public Const rnInputValue As String = "inputValue"
    Public Const RequestNameMemberID As String = "memberId"
    Public Const RequestNameFormID As String = "formId"
    Public Const RequestNamePageSize As String = "pageSize"
    Public Const RequestNamePageNumber As String = "pageNumber"
    Public Const RequestNameDetailSubtab As String = "subtab"
    '
    Public Const FormDetails As Integer = 1
    Public Const FormList As Integer = 2
    Public Const FormSearch As Integer = 3
    '
    Public Const ButtonSearch As String = "Search"
    Public Const ButtonCancel As String = "Cancel"
    Public Const ButtonCancelAll As String = "Cancel All"
    Public Const ButtonNewSearch As String = "New Search"
    Public Const ButtonApply As String = "Apply"
    Public Const ButtonSave As String = "Save"
    Public Const ButtonOK As String = "OK"
    '
    Public Const FieldTypeInteger As Integer = 1     ' An Long number
    Public Const FieldTypeText As Integer = 2     ' A text field (up To 255 characters)
    Public Const FieldTypeLongText As Integer = 3     ' A memo field (up To 8000 characters)
    Public Const FieldTypeBoolean As Integer = 4     ' A yes/no field
    Public Const FieldTypeDate As Integer = 5     ' A Date field
    Public Const FieldTypeFile As Integer = 6     ' A filename Of a file In the files directory.
    Public Const FieldTypeLookup As Integer = 7     ' A lookup Is a FieldTypeInteger that indexes into another table
    Public Const FieldTypeRedirect As Integer = 8     ' creates a link To another section
    Public Const FieldTypeCurrency As Integer = 9     ' A Float that prints In dollars
    Public Const FieldTypeFileText As Integer = 10     ' Text saved In a file In the files area.
    Public Const FieldTypeFileImage As Integer = 11     ' A filename Of a file In the files directory.
    Public Const FieldTypeFloat As Integer = 12     ' A float number
    Public Const FieldTypeAutoIdIncrement As Integer = 13     'Long that automatically increments With the New record
    Public Const FieldTypeManyToMany As Integer = 14     ' no database field - sets up a relationship through a Rule table To another table
    Public Const FieldTypeMemberSelect As Integer = 15     ' This ID Is a ccMembers record In a group defined by the MemberSelectGroupID field
    Public Const FieldTypeFileCSS As Integer = 16     ' A filename Of a CSS compatible file
    Public Const FieldTypeFileXML As Integer = 17     ' the filename Of an XML compatible file
    Public Const FieldTypeFileJavascript As Integer = 18     ' the filename Of a javascript compatible file
    Public Const FieldTypeLink As Integer = 19     ' Links used In href tags -- can go To pages Or resources
    Public Const FieldTypeResourceLink As Integer = 20     ' Links used In resources, link <img Or <Object. Should Not be pages
    Public Const FieldTypeHTML As Integer = 21     ' LongText field that expects HTML content
    Public Const FieldTypeFileHTML As Integer = 22     ' TextFile field that expects HTML content
    Public Const FieldTypeMax As Integer = 22             '
    '
    Public Const ReportFormVisitList As Integer = 9999
    '

    '
    '
    ' -- errors for resultErrList
    Public Enum resultErrorEnum
        errPermission = 50
        errDuplicate = 100
        errVerification = 110
        errRestriction = 120
        errInput = 200
        errAuthentication = 300
        errAdd = 400
        errSave = 500
        errDelete = 600
        errLookup = 700
        errLoad = 710
        errContent = 800
        errMiscellaneous = 900
    End Enum
    '
    ' -- http errors
    Public Enum httpErrorEnum
        badRequest = 400
        unauthorized = 401
        paymentRequired = 402
        forbidden = 403
        notFound = 404
        methodNotAllowed = 405
        notAcceptable = 406
        proxyAuthenticationRequired = 407
        requestTimeout = 408
        conflict = 409
        gone = 410
        lengthRequired = 411
        preconditionFailed = 412
        payloadTooLarge = 413
        urlTooLong = 414
        unsupportedMediaType = 415
        rangeNotSatisfiable = 416
        expectationFailed = 417
        teapot = 418
        methodFailure = 420
        enhanceYourCalm = 420
        misdirectedRequest = 421
        unprocessableEntity = 422
        locked = 423
        failedDependency = 424
        upgradeRequired = 426
        preconditionRequired = 428
        tooManyRequests = 429
        requestHeaderFieldsTooLarge = 431
        loginTimeout = 440
        noResponse = 444
        retryWith = 449
        redirect = 451
        unavailableForLegalReasons = 451
        sslCertificateError = 495
        sslCertificateRequired = 496
        httpRequestSentToSecurePort = 497
        invalidToken = 498
        clientClosedRequest = 499
        tokenRequired = 499
        internalServerError = 500
    End Enum
End Module

'
' During VB6 code upgrade, these are typical find-replace strings
'   - convert all routines to try-catch and not on error goto
'
'
'
' -- VB6 typical Replace List
'                   Find                                            Replace                                             Manual
'                   -----------------------------------------       --------------------------------------------------  ------------------------------
'                   a-s long                                        as integer
'                   a-s variant                                     as object
'                   k-maIndent(                                     nop(                                                -
'                   k-maEncodeInteger(Main.GetSiteProperty(         cp.site.getInteger(                                 - leaves double ending "))"
'                   M-ain.GetFormInputText(                         cp.html.inputText(
'                   M-ain.GetStreamText(                            cp.doc.getText(
'                   M-ain.GetStreamNumber(                          cp.doc.getNumber(
'                   M-ain.GetStreamBoolean(                         cp.doc.getBoolean(
'                   M-ain.GetStreamDate(                            cp.doc.getDate(
'                   D-OMDocument60                                  xml.xmlDocument
'                   M-ain.GetFormButton(                            cp.html.button("button",
'                   cp.utils.encodeBoolean                                cp.utils.encodeBoolean
'                   cp.utils.encodeInteger                                cp.utils.encodeInteger
'                   cp.utils.encodeNumber                                 cp.utils.encodeNumber
'                   cp.utils.encodeDate                                   cp.utils.encodeDate
'                   cp.user.isAdmin                                    cp.user.isAdmin
'                   cp.html.adminHint(                        cp.html.adminHint(
'                                                                   
'
'                   Manual issues
'                   ------------------------------------------------------------------------
'                   optionString = optionString & "&name=value"     cp.doc.setProperty( name, value )
'                   Main.GetFormStart() ... Main.GetFormEnd()       cp.html.form( ... )
'                   Err. -> convert to try+catch

