''''''''''''''''''''''''''''
' MAS Sales Orders Select testing.
'
''''''''''''''''''''''''''''

'Create ProvideX COM Object
Set oScript = CreateObject ("ProvideX.Script")

'Get the ODBC path for the last accessed installation of MAS 90/200
Const HKEY_CURRENT_USER = &H80000001
Set oReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
oReg.GetExpandedStringValue HKEY_CURRENT_USER,"Software\ODBC\ODBC.INI\SOTAMAS90","Directory",PathRoot
PathHome = PathRoot & "\Home"
dbg=false

if dbg then MsgBox(PathHome) end if
Set oReg = Nothing

'The Init method must be the first method called
oScript.Init(PathHome)

'The Session object must be the first MAS 90 object created
Set oSS = oScript.NewObject("SY_Session")

' Set the user for the Session
r = oSS.nLogon()
If r=0 Then
	'''''''''''''''''''''''
	' Enter: username, password
	user = Trim(InputBox("Enter User Name", "", "alex"))
	password = Trim(InputBox("Enter Password", "", ""))
	retVAL = oSS.nSetUser(User,Password)

	r = oss.nSetUser(user, password)
End If

' Set the company, module date and module for the Session
company = Trim(InputBox("Enter Company", "", "ABC"))
r = oss.nsetcompany(company)

'''''''''''' TMAS_Session.CreateCommonObjects
'TMAS_CustomerBUS.Create
'TMAS_CustomerContactBUS.Create
'TMAS_SalesPersonSVC.Create
'F_ShipToAddress_BUS := TMAS_ShipToAddressBUS.Create
'F_ShipToAddress_SVC := TMAS_ShipToAddressSVC.Create
'TMAS_CompanySVC.Create
'TMAS_SalesOrderTaxSummaryBUS.Create
'TMAS_InvoiceTaxSummaryBUS.Create
'TMAS_InvoiceSVC.Create
'TMAS_CustomerShipToTaxExemptionsSVC.Create
'TMAS_ItemCostSVC.Create
'TMAS_InvoiceDetailSVC.Create

'TMAS_Session.GetDocumentTypes
'F_SalesOrderDOCUMENT

'TMAS_Session.SetSession
tmpModuleName = ""
sMASDate = ""
r = oSS.nGetModuleInfo("S/O", tmpModuleName, sMASDate)
errchk r, "Error getting module info", oSS

sDate = oSS.sModuleDate
if sDate = "" then
  sDate = Year(Date) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2)      
end if  
r = oSS.nSetDate("S/O",sDate)
'errchk r, "Setting date", oSS

r = oSS.nSetModule("S/O")
errchk r, "Setting the module", oSS

' Instantiate the business object
r = oSS.nSetProgram(oSS.nLookupTask("SO_SalesOrder_UI"))
errchk r, "Setting the program", oSS

'TMAS_SalesOrderBus.Create
Set o = oScript.NewObject("SO_SalesOrder_bus", oSS)

'BuildFieldsList
dataSources = o.sGetDataSources
if dbg then MsgBox("Received data sources " & dataSources) end if
dataSources_ = splitstring_crlf(dataSources)
if dbg then MsgBox("Processed data sources " & dataSources_(0)) end if

fields_ = o.sGetColumns(dataSources_(0))
if dbg then MsgBox("Fields: " & fields_) end if

'DiscoverFieldsDescriptions
'set pPVXDAX = CreateObject("PVXDAX.PvxDictionary")
'bRet = pPVXDAX.SetDatabase(PathHome)
'if (not bRet) then
'	errchk 1, "Error creating PVXDAX object"
'end if

'"SO_SalesOrderHeader"

' FSession.RetrieveAvailableBatches
' FSession.RetrieveAvailableShippers

' FSession.F_Customer._GetKeyedFieldValuesList
' FSession.F_CustomerContact._GetKeyedFieldValuesList
' FSession.F_SalesPerson._GetKeyedFieldValuesList
' FSession.F_ShipToAddress_SVC._GetFieldValuesList
' FSession.F_ShipToAddress_BUS._GetKeyedFieldValuesList
' FSession.F_Company._GetFieldValuesList


'''''''''''''''''''''''
' Enter: Sales order Number
SalesOrderNo = trim(InputBox("Enter Sales Order number", "", "0000182"))
if dbg then MsgBox(SalesOrderNo) end if

r = o.nFind(SalesOrderNo)
errchk r, "Error finding", o

r = o.nSetKey(SalesOrderNo)
errchk r, "Error setting key", o

' MAS:_GetCurrentRecordFieldValuesList
wResult=""
wMASFields=""
r = o.nGetRecord(wResult, wMASFields)
errchk r, "Error getting record", o

if dbg then MsgBox("wMASFields = " & wMASFields) end if
if dbg then MsgBox("wResult = " & wresult) end if

values = splitString(wResult)
fields = splitString(fields_)

output_fields="The number of fields - " & CStr(UBound(fields))
output_values="The number of values - " & CStr(UBound(values))

output_equals="equals"
output_strFields=""
output_strValues=""
output_reference=""
if (UBound(values) <> UBound(fields)) then
	output_equals="does NOT equal"
end if

output_reference="For reference:"
output_strFields="Fields: " & fields_
output_strValues="Values: " & wResult

Wscript.Echo(output_fields & " " & output_equals & " " & output_values & Chr(13) & output_reference & Chr(13) & output_strFields & Chr(13) & output_strValues)

' Done with the object
r = o.DropObject()
o = 0

Wscript.Echo "DONE"

oss.nCleanup()		' Call Cleanup() before dropping the Session Object
oss.DropObject()
Set oss = Nothing
WScript.Quit