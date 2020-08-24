'TODO:  Follow these steps to enable the Ribbon (XML) item:

'1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Amanda()
'End Function

'2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
'   actions, such as clicking a button. Note: if you have exported this Ribbon from the
'   Ribbon designer, move your code from the event handlers to the callback methods and
'   modify the code to work with the Ribbon extensibility (RibbonX) programming model.

'3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

'For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
Imports Shell32
Imports IWshRuntimeLibrary
Imports File = IWshRuntimeLibrary.File
Imports System.Diagnostics
Imports System.Globalization
Imports System.Management.Automation
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Collections.ObjectModel
Imports System.Management.Automation.Runspaces
Imports Application = Microsoft.Office.Interop.Outlook.Application

<Runtime.InteropServices.ComVisible(True)>
Public Class amanda
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("MyOLRibbon.Amanda.xml")
    End Function

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub OnRibbonLoad(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
        stiPressed = False
        tlsPressed = False
    End Sub

    Public Function GetSize(ByVal control As Office.IRibbonControl) As String
        Select Case control.Id
            Case Is = "btnMoveSenderFolder"
                Return "large"
            Case Is = "btnMoveSenderFolder2"
                Return "large"
            Case Is = "btnMeterReq"
                Return "normal"
            Case Else
                Return "large"
        End Select
    End Function
    Sub GetOnAction(ByVal control As Office.IRibbonControl)
        'Dim olApp As Outlook.Application
        'olApp = DirectCast(Marshal.GetActiveObject("Outlook.Application"), Outlook.Application)
        Dim prevYearDt As DateTime
        Dim prevYear As String
        prevYearDt = DateAdd("yyyy", -1, Now())
        prevYear = Format(prevYearDt, "yyyy")
        Select Case control.Id
            Case Is = "btnSaveInvoice"
                SaveInvoice_PDF()
            Case Is = "btnSaveTaskOrder"
                SaveTaskOrder_PDF()
            Case Is = "btnSaveAttOnly"
                Response_SaveAttOnly()
            Case Is = "btnSaveResponse"
                Response_SaveAsPDFwAtt()
            Case Is = "btnSaveTicket"
                Ticket_SaveAsPDFwAtt()
            Case Is = "btnSavePRNotice"
                PRNotice_SaveAsPDF()
            Case Is = "btnArchiveARKSTIprevyear"
                ArchiveCompleteSJFolder("ARK", "STI", prevYear)
            Case Is = "btnArchiveARKSTIcuryear"
                ArchiveCompleteSJFolder("ARK", "STI")
            Case Is = "btnArchiveTULSTIprevyear"
                ArchiveCompleteSJFolder("TUL", "STI", prevYear)
            Case Is = "btnArchiveTULSTIcuryear"
                ArchiveCompleteSJFolder("TUL", "STI")
            Case Is = "btnArchiveTULTLS18"
                ArchiveCompleteSJFolder("TUL", "TLS", "2018")
            Case Is = "btnArchiveTULTLSprevyear"
                ArchiveCompleteSJFolder("TUL", "TLS", prevYear)
            Case Is = "btnArchiveTULTLScuryear"
                ArchiveCompleteSJFolder("TUL", "TLS")
            Case Is = "btnAccountFolders"
                CreateAccountFolders()
            Case Is = "btnASmedley"
                MoveToFolder("Folders", "TLS Employees\TUL\Amanda Kay Smedley")
            Case Is = "btnAddFolder"
                AddNewFolder()
            Case Is = "btnInvalidate"
                ribbon.Invalidate()
            Case Is = "btnMoveSenderFolder"
                MoveToSenderFolder()
            Case Is = "btnASPayStubs"
                MoveToFolder("Folders", "TLS Employees\TUL\Amanda Kay Smedley\PayStubs")
            Case Is = "btnASAccounts"
                MoveToFolder("Folders", "TLS Employees\TUL\Amanda Kay Smedley\Accounts")
            Case Is = "btnASNorth"
                MoveToFolder("Folders", "TLS Employees\TUL\Amanda Kay Smedley\Northstar")
            Case Is = "btnASGraphics"
                MoveToFolder("Folders", "TLS Employees\TUL\Amanda Kay Smedley\Graphics")
            Case Is = "btnMovePayroll"
                CopyToFolder("Folders", "Payroll")
                MoveToSenderFolder()
            Case Is = "btnSpam"
                MoveToFolder("Folders", "Support\Barracuda Spam Emails")
            Case Is = "btnMeterReq"
                MoveToFolder("Folders", "Vendors\RK Black\Meter Requests")
            Case Is = "btnSTUPS"
                MoveToFolder("Folders", "Vendors\UPS\SignalTek Repair UPS Notifications")
            Case Is = "btnRKBlack"
                MoveToFolder("Folders", "Vendors\RK Black")
            Case Is = "btnSendNewestPhoneList"
                SendContactList()
            Case Is = "btnEmailHR"
                SendEmail("tcarothers@tlsokc.com", "TLS2")
            Case Is = "btnEmailDeduction"
                SendDeductionEmail()
            Case Is = "btnRemovePrefixes"
                RemoveSubjectPrefix()
            Case Is = "btnMovePRResponses"
                MovePRResponses()
            Case Is = "btnAddBilledDetails"
                ' Dim myEntity As String = InputBox("Entity:", "Entity", "STI")
                Dim myJob As String = InputBox("Job Number:", "Small Job")
                CreateBilledDetails(myJob)
            Case Is = "btnCreateBillText"
                CreateBilledDetails(myJobNumberText)
                ribbon.Invalidate()
            Case Is = "btnMoveSenderFolder2"
                MoveToSenderFolder()
            Case Is = "btnSaveUSIC"
                SaveUSICResponse()
            Case Else
                MessageBox.Show("Error! No Action Found")
        End Select

    End Sub
    Public Function GetLabel(ByVal control As Office.IRibbonControl) As String
        Dim prevYearDt As DateTime
        Dim prevYear As String
        Dim curYear As String
        prevYearDt = DateAdd("yyyy", -1, Now())
        curYear = Format(Now(), "yyyy")
        prevYear = Format(prevYearDt, "yyyy")
        Select Case control.Id
            Case Is = "btnSaveUSIC"
                Return "Save USIC Response"
            Case Is = "btnArchiveTULTLS18"
                Return "TUL TLS (2018)"
            Case Is = "btnArchiveTULTLSprevyear"
                Return "TUL TLS (" & prevYear & ")"
            Case Is = "btnArchiveTULTLScuryear"
                Return "TUL TLS (" & curYear & ")"
            Case Is = "btnArchiveTULSTIprevyear"
                Return "TUL STI (" & prevYear & ")"
            Case Is = "btnArchiveTULSTIcuryear"
                Return "TUL STI (" & curYear & ")"
            Case Is = "btnArchiveARKSTIprevyear"
                Return "ARK STI (" & prevYear & ")"
            Case Is = "btnArchiveARKSTIcuryear"
                Return "ARK STI (" & curYear & ")"
            Case Is = "btnMoveSenderFolder"
                Return "Move to Sender"
            Case Is = "btnMoveSenderFolder2"
                Return "Move to Sender"
            Case Is = "btnMeterReq"
                Return "Meter Requests"
            Case Is = "btnMovePayroll"
                Return "Payroll"
            Case Is = "btnSpam"
                Return "Spam Blocker Emails"
            Case Is = "btnSTUPS"
                Return "ST UPS Notifs."
            Case Is = "btnRKBlack"
                Return "RK Black"
            Case Is = "btnAddBilledDetails"
                Return "Billed Details Txt"
            Case Else
                Return "Unknown Label"
        End Select

    End Function
    Public Function GetDescription(ByVal control As Office.IRibbonControl) As String
        Dim prevYearDt As DateTime
        Dim prevYear As String
        Dim curYear As String
        prevYearDt = DateAdd("yyyy", -1, Now())
        curYear = Format(Now(), "yyyy")
        prevYear = Format(prevYearDt, "yyyy")
        Select Case control.Id
            Case Is = "btnArchiveTULTLS18"
                Return "Archive Tulsa TLS Small Job Folder for 2018."
            Case Is = "btnArchiveTULTLSprevyear"
                Return "Archive Tulsa TLS Small Job Folder for " & prevYear & "."
            Case Is = "btnArchiveTULTLScuryear"
                Return "Archive Tulsa TLS Small Job Folder for " & curYear & "."
            Case Is = "btnArchiveTULSTIprevyear"
                Return "Archive Tulsa STI Small Job Folder for " & prevYear & "."
            Case Is = "btnArchiveTULSTIcuryear"
                Return "Archive Tulsa STI Small Job Folder for " & curYear & "."
            Case Is = "btnArchiveARKSTIprevyear"
                Return "Archive Arkansas STI Small Job Folder for " & prevYear & "."
            Case Is = "btnArchiveARKSTIcuryear"
                Return "Archive Arkansas STI Small Job Folder for " & curYear & "."
            Case Is = "btnDGill"
                Return "ARK"
            Case Is = "btnVacancySTAsst"
                Return "Vacant Tech's Asst."
            Case Is = "btnARKWHMgr"
                Return "Arkansas"
            Case Is = "btnMoveSenderFolder"
                Return "Sort to Folders"
            Case Is = "btnMoveSenderFolder2"
                Return "Sort to Folders"
            Case Is = "btnJMudge"
                Return "Tulsa"
            Case Is = "btnEFrench"
                Return "Tulsa"
            Case Is = "btnSCruz"
                Return "Tulsa"
            Case Is = "btnTGreen"
                Return "Tulsa"
            Case Is = "btnJGuerrero"
                Return "Tulsa"
            Case Is = "btnALittlefield"
                Return "Tulsa"
            Case Is = "btnSStrode"
                Return "Tulsa"
            Case Is = "btnNVanDalsem"
                Return "Tulsa"
            Case Is = "btnMeterReq"
                Return "RK Black Meter Requests"
            Case Is = "btnMovePayroll"
                Return "Move to Payroll"
            Case Is = "btnSpam"
                Return "Barracuda Spam"
            Case Is = "btnSTUPS"
                Return "ST UPS Notifs."
            Case Is = "btnRKBlack"
                Return "RK Black"
            Case Is = "btnJettMudge"
                Return "Tulsa"
            Case Is = "btnAddBilledDetails"
                Return "Small Jobs Billing"
            Case Is = "btnBJohnson"
                Return "AP Clerk"
            Case Is = "btnTBond"
                Return "Controller"
            Case Is = "btnMovePRResponses"
                Return "Move PR Responses to Archive"
            Case Else
                Return "Unknown Description"
        End Select
    End Function
    Public Function GetScreentip(ByVal control As Office.IRibbonControl) As String
        Select Case control.Id
            Case Is = "btnArchiveTULTLS18"
                Return "Tulsa Traffic & Lighting Systems"
            Case Is = "btnArchiveTULTLSprevyear"
                Return "Tulsa Traffic & Lighting Systems"
            Case Is = "btnArchiveTULTLScuryear"
                Return "Tulsa Traffic & Lighting Systems"
            Case Is = "btnArchiveTULSTIprevyear"
                Return "Tulsa SignalTek"
            Case Is = "btnArchiveTULSTIcuryear"
                Return "Tulsa SignalTek"
            Case Is = "btnArchiveARKSTIprevyear"
                Return "Arkansas SignalTek"
            Case Is = "btnArchiveARKSTIcuryear"
                Return "Arkansas SignalTek"
            Case Is = "btnDGill"
                Return "Cell Number / Extension"
            Case Is = "btnVacancySTAsst"
                Return "Cell Number"
            Case Is = "btnARKWHMgr"
                Return "Cell Number / Ext"
            Case Is = "btnMoveSenderFolder"
                Return "Move Selected Emails to Sender Folder"
            Case Is = "btnMoveSenderFolder2"
                Return "Move Selected Emails to Sender Folder"
            Case Is = "btnJMudge"
                Return "Cell Number / Ext"
            Case Is = "btnEFrench"
                Return "Office Extension"
            Case Is = "btnMeterReq"
                Return "RK Black Meter Requests"
            Case Is = "btnMovePayroll"
                Return "Move to Payroll"
            Case Is = "btnSpam"
                Return "Barracuda Spam"
            Case Is = "btnSTUPS"
                Return "ST UPS Notifs."
            Case Is = "btnRKBlack"
                Return "3 Nines"
            Case Is = "btnSCruz"
                Return "Cell Number"
            Case Is = "btnALittlefield"
                Return "Cell Number"
            Case Is = "btnJGuerrero"
                Return "Cell Number"
            Case Is = "btnSStrode"
                Return "Cell Number"
            Case Is = "btnNVanDalsem"
                Return "Cell Number"
            Case Is = "btnJettMudge"
                Return "Cell Number"
            Case Is = "btnAddBilledDetails"
                Return "Add Billing Info"
            Case Is = "btnBJohnson"
                Return "Office Extension"
            Case Is = "btnTBond"
                Return "Office Extension"
            Case Else
                Return "Unknown Description"
        End Select
    End Function
    Public Function GetSupertip(ByVal control As Office.IRibbonControl) As String
        Dim prevYearDt As DateTime
        Dim prevYear As String
        Dim curYear As String
        prevYearDt = DateAdd("yyyy", -1, Now())
        curYear = Format(Now(), "yyyy")
        prevYear = Format(prevYearDt, "yyyy")
        Select Case control.Id
            Case Is = "btnArchiveTULTLS18"
                Return "2018"
            Case Is = "btnArchiveTULTLSprevyear"
                Return prevYear
            Case Is = "btnArchiveTULTLScuryear"
                Return curYear
            Case Is = "btnArchiveTULSTIprevyear"
                Return prevYear
            Case Is = "btnArchiveTULSTIcuryear"
                Return curYear
            Case Is = "btnArchiveARKSTIprevyear"
                Return prevYear
            Case Is = "btnArchiveARKSTIcuryear"
                Return curYear
            Case Is = "btnMoveSenderFolder"
                Return "For Pre-Set Names"
            Case Is = "btnMoveSenderFolder2"
                Return "For Pre-Set Names"
            Case Is = "btnMeterReq"
                Return "Email is Quarterly"
            Case Is = "btnMovePayroll"
                Return "Move to Payroll"
            Case Is = "btnSpam"
                Return "Barracuda Spam"
            Case Is = "btnSTUPS"
                Return "ST UPS Notifs."
            Case Is = "btnRKBlack"
                Return "RK Black"
            Case Is = "btnAddBilledDetails"
                Return "Add text file to Bid Info"
            Case Else
                Return "Unknown Description"
        End Select
    End Function
    Public Function GetImage(ByVal control As Office.IRibbonControl) As stdole.IPictureDisp
        Dim pictureDisp As stdole.IPictureDisp

        Select Case control.Id
            Case Is = "btnSaveAttOnly"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iEmail)
            Case Is = "btnSaveResponse"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iLocateCompass)
            Case Is = "btnSaveTicket"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iLocateTicket)
            Case Is = "btnSavePRNotice"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iLocateCompass)
            Case Is = "mnuArchiveTLS"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iTULTLS)
            Case Is = "mnuArchiveSTI"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iTULST)
            Case Is = "btnArchiveTULTLS18"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iArchiveFolder)
            Case Is = "btnArchiveTULTLSprevyear"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iArchiveFolder)
            Case Is = "btnArchiveTULTLScuryear"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iTULTLS)
            Case Is = "btnArchiveTULSTIprevyear"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iArchiveFolder)
            Case Is = "btnArchiveTULSTIcuryear"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iTULST)
            Case Is = "btnArchiveARKSTIprevyear"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iArchiveFolder)
            Case Is = "btnArchiveARKSTIcuryear"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iARKST)
            Case Is = "btnAddFolder"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iAddFolder)
            Case Is = "btnInvalidate"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iRefreshControls)
            Case Is = "btnMoveSenderFolder"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iOpenFolder2)
            Case Is = "btnMoveSenderFolder2"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iOpenFolder2)
            Case Is = "grpDept"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iOKCGrp)
            Case Is = "mnuExecs"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iExecGrp)
            Case Is = "mnuSafety"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iSafetyGrp)
            Case Is = "mnuProjMgr"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iProjMgrGrp)
            Case Is = "mnuAccounting"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iAccountingGrp)
            Case Is = "mnuOffice"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iOKCGrp)
            Case Is = "mnuAmanda"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iASmedley)
            Case Is = "btnASmedley"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iASmedley)
            Case Is = "mnuWarehouse"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iShopGrp)
            Case Is = "mnuForemans"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.gear_user_group)
            Case Is = "grpForemans"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.gear_user_group)
            Case Is = "btnASPayStubs"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iMove2Folder)
            Case Is = "btnASAccounts"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iMove2Folder)
            Case Is = "btnASNorth"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iMove2Folder)
            Case Is = "btnASGraphics"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iMove2Folder)
            Case Is = "btnRemovePrefixes"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iEmail)
            Case Is = "btnAddBilledDetails"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.briefcase)
            Case Is = "btnSaveTaskOrder"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.briefcase)
            Case Is = "btnSaveUSIC"
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iMove2Folder)
            Case Else
                pictureDisp = PictureConverter.ImageToPictureDisp(My.Resources.iUser)
        End Select
        Return pictureDisp

    End Function
    Function GetItemCount(control As IRibbonControl) As Integer
        Select Case control.Id
            Case "ddBrowse2QuoteFolders"
                Return GetFolder("\\ASmedley@tlsokc.com\Folders\Small Jobs\Quotes - Pending").Folders.Count
            Case "ddBrowse2ArkEmpFolders"
                Return GetFolder("\\ASmedley@tlsokc.com\Folders\TLS Employees\ARK").Folders.Count
            Case "ddBrowse2OkcEmpFolders"
                Return GetFolder("\\ASmedley@tlsokc.com\Folders\TLS Employees\OKC").Folders.Count
            Case "ddBrowse2TulEmpFolders"
                Return GetFolder("\\ASmedley@tlsokc.com\Folders\TLS Employees\TUL").Folders.Count
            Case "ddBrowse2BArchiveFolders"
                Return GetFolder("\\ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\TUL\TLS\2019").Folders.Count
            Case "ddBrowse2BSTArchiveFolders"
                Return GetFolder("\\ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\TUL\STI\2019").Folders.Count
            Case "ddBSTArchiveFolders"
                Return GetFolder("\\ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\TUL\STI\2019").Folders.Count
            Case "ddBArchiveFolders"
                Return GetFolder("\\ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\TUL\TLS\2019").Folders.Count
            Case "ddBSTArchiveFolders20"
                Return GetFolder("\\ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\TUL\STI\2020").Folders.Count
            Case "ddBArchiveFolders20"
                Return GetFolder("\\ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\TUL\TLS\2020").Folders.Count
            Case "ddASTFolders"
                Return GetFolder("\\ASmedley@tlsokc.com\Folders\Small Jobs\ARK\STI").Folders.Count
            Case "ddQuoteFolders"
                Return GetFolder("\\ASmedley@tlsokc.com\Folders\Small Jobs\Quotes - Pending").Folders.Count
            Case "ddVendorFolders"
                Return GetFolder("\\ASmedley@tlsokc.com\Folders\Vendors").Folders.Count
            Case "ddSTIFolders"
                Return GetFolder("\\ASmedley@tlsokc.com\Folders\Small Jobs\TUL\STI").Folders.Count
            Case "ddTLSFolders"
                Return GetFolder("\\ASmedley@tlsokc.com\Folders\Small Jobs\TUL\TLS").Folders.Count
            Case "ddJobFolders"
                Return GetFolder("\\ASmedley@tlsokc.com\Folders\Jobs").Folders.Count
            Case "ddConCity"
                Return GetFolder("\\ASmedley@tlsokc.com\Folders\SignalTek\Cities\Contract").Folders.Count
            Case "ddNonConCity"
                Return GetFolder("\\ASmedley@tlsokc.com\Folders\SignalTek\Cities\Non-Contract").Folders.Count
            Case Else
                Return MessageBox.Show("Control doesn't exist")
        End Select

    End Function
    Function GetItemLabel(control As IRibbonControl, index As Integer) As String
        Select Case control.Id
            Case "ddBrowse2QuoteFolders"
                GetFolderArray_Quotes()
                Return myQuoteFolderArray(index)
            Case "ddBrowse2ArkEmpFolders"
                GetFolderArray_ARKEmployees()
                Return myARKEmployeesArray(index)
            Case "ddBrowse2OkcEmpFolders"
                GetFolderArray_OKCEmployees()
                Return myOKCEmployeesArray(index)
            Case "ddBrowse2TulEmpFolders"
                GetFolderArray_TULEmployees()
                Return myTULEmployeesArray(index)
            Case "ddBrowse2BArchiveFolders"
                GetFolderArray_ArchiveB2019()
                Return myB2019ArchiveArray(index)
            Case "ddBrowse2BSTArchiveFolders"
                GetFolderArray_ArchiveBST2019()
                Return myBST2019ArchiveArray(index)
            Case "ddBSTArchiveFolders"
                GetFolderArray_ArchiveBST2019()
                Return myBST2019ArchiveArray(index)
            Case "ddBArchiveFolders"
                GetFolderArray_ArchiveB2019()
                Return myB2019ArchiveArray(index)
            Case "ddBSTArchiveFolders20"
                GetFolderArray_ArchiveBST2020()
                Return myBST2020ArchiveArray(index)
            Case "ddBArchiveFolders20"
                GetFolderArray_ArchiveB2020()
                Return myB2020ArchiveArray(index)
            Case "ddASTFolders"
                GetFolderArray_ASTSJ()
                Return myASTSmallJobsArray(index)
            Case "ddQuoteFolders"
                GetFolderArray_Quotes()
                Return myQuoteFolderArray(index)
            Case "ddVendorFolders"
                GetFolderArray_Vendors()
                Return myVendorFolderArray(index)
            Case "ddSTIFolders"
                GetFolderArray_STISJ()
                Return mySTISmallJobsArray(index)
            Case "ddTLSFolders"
                GetFolderArray_TLSSJ()
                Return myTLSSmallJobsArray(index)
            Case "ddJobFolders"
                GetFolderArray_Jobs()
                Return myJobFoldersArray(index)
            Case "ddConCity"
                GetFolderArray_ContractCities()
                Return myContractCitiesArray(index)
            Case "ddNonConCity"
                GetFolderArray_NonContractCities()
                Return myNonContractCitiesArray(index)
            Case Else
                Return MessageBox.Show("Control doesn't exist")
        End Select


    End Function
    Function GetItemID(control As IRibbonControl, index As Integer) As String
        Select Case control.Id
            Case "ddQuoteFolders"
                Return "folder" & index
            Case Else
                Return "folder" & index
        End Select

    End Function
    Sub DropDownAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
        Dim myFolder As String
        Select Case control.Id
            Case "ddBrowse2QuoteFolders"
                myFolder = CStr("\\ASmedley@tlsokc.com\Folders\Small Jobs\Quotes - Pending\" & myQuoteFolderArray(selectedIndex))
                Browse2Folder(myFolder)
            Case "ddBrowse2ArkEmpFolders"
                myFolder = CStr("\\ASmedley@tlsokc.com\Folders\TLS Employees\ARK\" & myARKEmployeesArray(selectedIndex))
                Browse2Folder(myFolder)
            Case "ddBrowse2OkcEmpFolders"
                myFolder = CStr("\\ASmedley@tlsokc.com\Folders\TLS Employees\OKC\" & myOKCEmployeesArray(selectedIndex))
                Browse2Folder(myFolder)
            Case "ddBrowse2TulEmpFolders"
                myFolder = CStr("\\ASmedley@tlsokc.com\TLS Employees\TUL\" & myTULEmployeesArray(selectedIndex))
                Browse2Folder(myFolder)
            Case "ddBrowse2BSTArchiveFolders"
                myFolder = CStr("\\ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\TUL\STI\2019\" & myBST2019ArchiveArray(selectedIndex))
                Browse2Folder(myFolder)
            Case "ddBrowse2BArchiveFolders"
                myFolder = CStr("\\ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\TUL\TLS\2019\" & myB2019ArchiveArray(selectedIndex))
                Browse2Folder(myFolder)
            Case "ddBSTArchiveFolders"
                myFolder = CStr("Small Jobs\Complete & Billed\TUL\STI\2019\" & myBST2019ArchiveArray(selectedIndex))
                MoveToFolder("SJArchive", myFolder)
            Case "ddBSTArchiveFolders20"
                myFolder = CStr("Small Jobs\Complete & Billed\TUL\STI\2020\" & myBST2020ArchiveArray(selectedIndex))
                MoveToFolder("SJArchive", myFolder)
            Case "ddBArchiveFolders"
                myFolder = CStr("Small Jobs\Complete & Billed\TUL\TLS\2019\" & myB2019ArchiveArray(selectedIndex))
                MoveToFolder("SJArchive", myFolder)
            Case "ddBArchiveFolders20"
                myFolder = CStr("Small Jobs\Complete & Billed\TUL\TLS\2020\" & myB2020ArchiveArray(selectedIndex))
                MoveToFolder("SJArchive", myFolder)
            Case "ddASTFolders"
                myFolder = CStr("Small Jobs\ARK\STI\" & myASTSmallJobsArray(selectedIndex))
                MoveToFolder("Folders", myFolder)
            Case "ddQuoteFolders"
                myFolder = CStr("Small Jobs\Quotes - Pending\" & myQuoteFolderArray(selectedIndex))
                MoveToFolder("Folders", myFolder)
            Case "ddVendorFolders"
                myFolder = CStr("Vendors\" & myVendorFolderArray(selectedIndex))
                MoveToFolder("Folders", myFolder)
            Case "ddSTIFolders"
                myFolder = CStr("Small Jobs\TUL\STI\" & mySTISmallJobsArray(selectedIndex))
                MoveToFolder("Folders", myFolder)
            Case "ddTLSFolders"
                myFolder = CStr("Small Jobs\TUL\TLS\" & myTLSSmallJobsArray(selectedIndex))
                MoveToFolder("Folders", myFolder)
            Case "ddJobFolders"
                myFolder = CStr("Jobs\" & myJobFoldersArray(selectedIndex))
                MoveToFolder("Folders", myFolder)
            Case "ddConCity"
                myFolder = CStr("SignalTek\Cities\Contract\" & myContractCitiesArray(selectedIndex))
                MoveToFolder("Folders", myFolder)
            Case "ddNonConCity"
                myFolder = CStr("SignalTek\Cities\Non-Contract\" & myNonContractCitiesArray(selectedIndex))
                MoveToFolder("Folders", myFolder)
            Case Else
                MessageBox.Show("Control doesn't exist")
        End Select
        ribbon.InvalidateControl(control.Id)
    End Sub

    Public Sub tgl_OnAction(ByVal control As Office.IRibbonControl, ByRef pressed As Boolean)
        Select Case control.Id
            Case "tglSTI"
                stiPressed = True
                tlsPressed = False
                myToggleCompany = "STI"
            Case "tglTLS"
                stiPressed = False
                tlsPressed = True
                myToggleCompany = "TLS"
        End Select
    End Sub
    Public Sub GetPressedToggle(ByVal control As Office.IRibbonControl, ByRef val As VariantType)
        Select Case control.Id
            Case "tglSTI"
                val = stiPressed
            Case "tglTLS"
                val = tlsPressed
            Case Else
                MessageBox.Show("Error, no toggle button pressed.")
                val = ""
                Exit Sub
        End Select
    End Sub

    Public Function GetEditBoxText(ByVal control As Office.IRibbonControl) As String
        Select Case control.Id
            Case "txtJobNumber"
                GetEditBoxText = "Job #"
            Case Else
                GetEditBoxText = "Unknown"
        End Select

    End Function

    Public Sub GetOnChange(ByVal control As Office.IRibbonControl, ByVal text As String)
        Select Case control.Id
            Case "txtJobNumber"
                myJobNumberText = text
            Case Else
                myJobNumberText = "Nothing"
        End Select
        Console.Write("myJobNumberText=" & myJobNumberText)
        Console.Write("midJobNumText=" & Mid(myJobNumberText, 1, 4))
    End Sub
#End Region

#Region "Helpers"

    Public myQuoteFolderArray As String()
    Public myVendorFolderArray As String()
    Public mySTISmallJobsArray As String()
    Public myASTSmallJobsArray As String()
    Public myTLSSmallJobsArray As String()
    Public myJobFoldersArray As String()
    Public myContractCitiesArray As String()
    Public myNonContractCitiesArray As String()
    Public myBST2019ArchiveArray As String()
    Public myBST2020ArchiveArray As String()
    Public myB2019ArchiveArray As String()
    Public myB2020ArchiveArray As String()
    Public myARKEmployeesArray As String()
    Public myTULEmployeesArray As String()
    Public myOKCEmployeesArray As String()
    Public myJobNumberText As String
    Public myToggleCompany As String
    Public senderFolder As String
    Public myTicketNumber As String
    Public myMemberCode As String
    Public memCodeRegExPattern As String
    Public tktNumRegExPattern As String
    Public Const defaultStatus As String = "Processing..."
    Public Shared isCancelled As Boolean
    Public Shared strStatus As String
    Public Shared progressValue As Long
    Public stiPressed As Boolean
    Public tlsPressed As Boolean
    Public Const SpeedUp As Boolean = True
    Public Const StopAtFirstMatch As Boolean = True

    'Public Sub FindFolder()
    ' Dim sName As String
    'Dim oFolders As Folders
    '   m_Folder = Nothing
    '  m_Find = ""
    ' m_Wildcard = False
    ' sName = InputBox("Find:", "Search Folder")
    ' If Len(Trim(sName)) = 0 Then Exit Sub
    '    m_Find = sName
    '   m_Find = LCase(m_Find)
    ''  m_Find = Replace(m_Find, "%", "*")
    '  m_Wildcard = (InStr(m_Find, "*"))
    '
    '    oFolders = Application.Session.Folders
    '   LoopFolders oFolders

    'If Not m_Folder Is Nothing Then
    'If MsgBox("Activate Folder: " & vbCrLf & m_Folder.FolderPath, vbQuestion Or vbYesNo) = vbYes Then
    '           Application.ActiveExplorer.CurrentFolder = m_Folder
    'End If
    'Else
    '       MsgBox("Not Found", vbInformation)
    'End If
    'End Sub

    'Public Sub LoopFolders(Folders As Outlook.Folders)
    'Dim oFolder As MAPIFolder
    'Dim bFound As Boolean
    '
    'If SpeedUp = False Then System.Windows.Forms.Application.DoEvents()

    'For Each oFolder In Folders
    ' If m_Wildcard Then
    '            bFound = (LCase(oFolder.Name) Like m_Find)
    'Else
    '           bFound = (LCase(oFolder.Name) = m_Find)
    'End If
    '
    'If bFound Then
    'If StopAtFirstMatch = False Then
    'If MsgBox("Found: " & vbCrLf & oFolder.FolderPath & vbCrLf & vbCrLf & "Continue?", vbQuestion Or vbYesNo) = vbYes Then
    '                   bFound = False
    'End If
    'End If
    'End If
    'If bFound Then
    '           m_Folder = oFolder
    'Exit For
    'ElseIf bFound = False Then
    '           LoopFolders(oFolder.Folders)
    'If Not m_Folder Is Nothing Then Exit For
    'End If
    'Next

    'End Sub
    Private Sub NAR(ByVal o As Object)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
            End While
        Catch
        Finally
            o = Nothing
        End Try
    End Sub
    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function
    Public Function GetFolderArray_Quotes() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\Small Jobs\Quotes - Pending")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing, " & vbNewLine & "Quotes - Pending doesn't exist in the current folder structure.")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        'Dim myFolders = CreateObject("System.Collections.ArrayList")
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        myQuoteFolderArray = myFolders.ToArray()
        'myQuoteFolderArray = myFolders.ToString
        Return True

        NAR(Folder)
        NAR(subFolders)
        NAR(oFolder)
        NAR(oNS)
        NAR(oOut)


    End Function
    Public Function GetFolderArray_Vendors() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\Vendors")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        'Dim myFolders = CreateObject("System.Collections.ArrayList")
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        myVendorFolderArray = myFolders.ToArray()
        'myVendorFolderArray = myFolders
        Return True
    End Function
    Public Function GetFolderArray_STISJ() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\Small Jobs\TUL\STI")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        'Dim myFolders = CreateObject("System.Collections.ArrayList")
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        mySTISmallJobsArray = myFolders.ToArray()
        'mySTISmallJobsArray = myFolders
        Return True
    End Function
    Public Function GetFolderArray_ASTSJ() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\Small Jobs\ARK\STI")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        'Dim myFolders = CreateObject("System.Collections.ArrayList")
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        myASTSmallJobsArray = myFolders.ToArray()
        'mySTISmallJobsArray = myFolders
        Return True
    End Function
    Public Function GetFolderArray_TLSSJ() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\Small Jobs\TUL\TLS")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        'Dim myFolders = CreateObject("System.Collections.ArrayList")
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        myTLSSmallJobsArray = myFolders.ToArray()
        'myTLSSmallJobsArray = myFolders
        Return True
    End Function
    Public Function GetFolderArray_Jobs() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\Jobs")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        myJobFoldersArray = myFolders.ToArray()
        Return True
    End Function
    Public Function GetFolderArray_ArchiveBST2019() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\TUL\STI\2019")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        myBST2019ArchiveArray = myFolders.ToArray()
        Return True
    End Function
    Public Function GetFolderArray_ArchiveBST2020() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\TUL\STI\2020")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        myBST2020ArchiveArray = myFolders.ToArray()
        Return True
    End Function
    Public Function GetFolderArray_ArchiveB2019() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\TUL\TLS\2019")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        myB2019ArchiveArray = myFolders.ToArray()
        Return True
    End Function
    Public Function GetFolderArray_ArchiveB2020() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\TUL\TLS\2020")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        myB2020ArchiveArray = myFolders.ToArray()
        Return True
    End Function
    Public Function GetFolderArray_ContractCities() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\SignalTek\Cities\Contract")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        myContractCitiesArray = myFolders.ToArray()
        Return True
    End Function
    Public Function GetFolderArray_NonContractCities() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\SignalTek\Cities\Non-Contract")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        myNonContractCitiesArray = myFolders.ToArray()
        Return True
    End Function
    Public Function GetFolderArray_ARKEmployees() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\TLS Employees\ARK")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        myARKEmployeesArray = myFolders.ToArray()
        Return True
    End Function
    Public Function GetFolderArray_TULEmployees() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\TLS Employees\TUL")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        myTULEmployeesArray = myFolders.ToArray()
        Return True
    End Function
    Public Function GetFolderArray_OKCEmployees() As Boolean
        Dim oOut As Outlook.Application
        Dim oNS As Outlook.NameSpace
        Dim oFolder As Outlook.MAPIFolder
        Dim subFolders As Object
        Dim Folder As Outlook.MAPIFolder
        On Error Resume Next
        oOut = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If oOut Is Nothing Then oOut = CreateObject("Outlook.Application")
        oNS = oOut.GetNamespace("MAPI")
        oFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\TLS Employees\OKC")
        If oFolder Is Nothing Then
            Return False
            MessageBox.Show("oFolder was Nothing")
            Exit Function
        Else
        End If
        subFolders = oFolder.Folders
        Dim myFolders As New List(Of String)()
        For Each Folder In subFolders
            myFolders.Add(Folder.Name)
        Next
        myFolders.Sort()
        myOKCEmployeesArray = myFolders.ToArray()
        Return True
    End Function
    Public Function GetSenderFolder(ByVal senderName As String) As Boolean
        Dim strFolder As String
        Dim arkPath As String
        Dim tulPath As String
        Dim okcPath As String
        arkPath = "\\ASmedley@tlsokc.com\Folders\TLS Employees\ARK\"
        tulPath = "\\ASmedley@tlsokc.com\Folders\TLS Employees\TUL\"
        okcPath = "\\ASmedley@tlsokc.com\Folders\TLS Employees\OKC\"
        Select Case senderName
            Case Is = "Alejandra Ochoa"
                strFolder = tulPath & "Alejandra Ochoa"
            Case Is = "Sebastian Castaneda"
                strFolder = okcPath & "Sebastian Castaneda"
            Case Is = "Juli Aguiniga"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\Gades Sales"
            Case Is = "Nicholas Carpenter"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\RK Black"
            Case Is = "Trevor Turner"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\RK Black"
            Case Is = "R.K. Black, Inc. Help Desk"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\RK Black"
            Case Is = "R.K. Black Help Desk"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\RK Black"
            Case Is = "Brooke Wincapaw"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\RK Black"
            Case Is = "itservices@rkblack.com"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\RK Black"
            Case Is = "tls.ricoh@gmail.com"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\RK Black"
            Case Is = "Heath Hendrix"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\RK Black"
            Case Is = "RK Black Help Desk"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\RK Black"
            Case Is = "Justin Bloomfield"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\RK Black"
            Case Is = "RK Black, Inc. Help Desk"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\RK Black"
            Case Is = "Robert Taylor"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\RK Black"
            Case Is = "Kelley Deardeuff"
                strFolder = okcPath & "Kelley Deardeuff"
            Case Is = "Josh Fisher"
                strFolder = tulPath & "Josh Fisher"
            Case Is = "Bobby Austin"
                strFolder = tulPath & "Bobby Austin"
            Case Is = "Todd Gowen"
                strFolder = tulPath & "Todd Gowen"
            Case Is = "Justin Bloomfield"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\RK Black"
            Case Is = "Justin Dorsey"
                strFolder = tulPath & "Justin Dorsey"
            Case Is = "tcarothers@tlsokc.com"
                strFolder = okcPath & "Tonya Carothers"
            Case Is = "Human Resources"
                strFolder = okcPath & "Tonya Carothers"
            Case Is = "Abisai Hernandez"
                strFolder = okcPath & "Abisai Hernandez"
            Case Is = "Target"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\Target"
            Case Is = "Terry Krajicek"
                strFolder = tulPath & "Terry Krajicek"
            Case Is = "Smedley, Oakie D"
                strFolder = tulPath & "Amanda Kay Smedley"
            Case Is = "Becky Smedley"
                strFolder = tulPath & "Amanda Kay Smedley"
            Case Is = "Amanda Smedley"
                strFolder = tulPath & "Amanda Kay Smedley"
            Case Is = "Amanda K. Smedley"
                strFolder = tulPath & "Amanda Kay Smedley"
            Case Is = "Amanda Kay Smedley"
                strFolder = tulPath & "Amanda Kay Smedley"
            Case Is = "asmedley@tlsokc.com"
                strFolder = tulPath & "Amanda Kay Smedley"
            Case Is = "Antonio Gonzalez"
                strFolder = tulPath & "Mechanics\Antonio Gonzalez"
            Case Is = "Lea Ann Forbes"
                strFolder = tulPath & "Lea Ann Forbes"
            Case Is = "Gary Forbes"
                strFolder = tulPath & "Gary Forbes"
            Case Is = "Taylor Butler"
                strFolder = tulPath & "Taylor Butler"
            Case Is = "Dusty Scott"
                strFolder = okcPath & "Dusty Scott"
            Case Is = "Josh Hartley"
                strFolder = okcPath & "Josh Hartley"
            Case Is = "Terry Bond"
                strFolder = okcPath & "Terry Bond"
            Case Is = "Derrill West"
                strFolder = tulPath & "Derrill West"
            Case Is = "Justin Mudge"
                strFolder = tulPath & "Mechanics\Justin Mudge"
            Case Is = "Stephen Strode"
                strFolder = tulPath & "Mechanics\Stephen Strode"
            Case Is = "Sam Cruz"
                strFolder = tulPath & "Mechanics\Sam Cruz"
            Case Is = "Anthony Littlefield"
                strFolder = tulPath & "Mechanics\Anthony Littlefield"
            Case Is = "Jett Mudge"
                strFolder = tulPath & "Mechanics\Jett Mudge"
            Case Is = "Guillermo Cortes"
                strFolder = tulPath & "Guillermo Cortes"
            Case Is = "Celerino Del Valle"
                strFolder = tulPath & "Celerino Del Valle"
            Case Is = "Wayne Kennedy"
                strFolder = tulPath & "Wayne Kennedy"
            Case Is = "Brandon Parker"
                strFolder = tulPath & "Brandon Parker"
            Case Is = "Cavin Coffelt"
                strFolder = tulPath & "Cavin Coffelt"
            Case Is = "Erika French"
                strFolder = tulPath & "Mechanics\Erika French"
            Case Is = "Luis Delgado"
                strFolder = tulPath & "Luis Delgado"
            Case Is = "John Fissel"
                strFolder = tulPath & "John Fissel"
            Case Is = "Clint Foreman"
                strFolder = tulPath & "Clint Foreman"
            Case Is = "Mario Gonzalez"
                strFolder = tulPath & "Mario Gonzalez"
            Case Is = "Robert Kent"
                strFolder = tulPath & "Robert Kent"
            Case Is = "Mike McCall"
                strFolder = tulPath & "Mike McCall"
            Case Is = "Raymond Ashlock"
                strFolder = tulPath & "Raymond Ashlock"
            Case Is = "Nathan VanDalsem"
                strFolder = tulPath & "Mechanics\Nathan VanDalsem"
            Case Is = "Larry Butler"
                strFolder = okcPath & "Larry Butler"
            Case Is = "Willie Touchette"
                strFolder = okcPath & "Larry Touchette"
            Case Is = "Tonya Carothers"
                strFolder = okcPath & "Tonya Carothers"
            Case Is = "Walter Floyd"
                strFolder = okcPath & "Walter Floyd"
            Case Is = "Cathy J. Willis"
                strFolder = okcPath & "Cathy J Willis"
            Case Is = "David Willis"
                strFolder = okcPath & "David Willis"
            Case Is = "Matthew Creps"
                strFolder = okcPath & "Matt Creps"
            Case Is = "Tracy Willis"
                strFolder = okcPath & "Tracy Willis"
            Case Is = "Joquita Butler"
                strFolder = okcPath & "Joquita Butler"
            Case Is = "Jay Williams"
                strFolder = okcPath & "Jay Williams"
            Case Is = "Charlie Butler"
                strFolder = okcPath & "Charlie Butler"
            Case Is = "James McGinley"
                strFolder = okcPath & "James McGinley"
            Case Is = "lwillis@tlsokc.com"
                strFolder = okcPath & "Loren Willis"
            Case Is = "Robert Lott"
                strFolder = okcPath & "Robert Lott"
            Case Is = "Red Chambers"
                strFolder = okcPath & "Red Chambers"
            Case Is = "Osman Tavilson"
                strFolder = okcPath & "Osman Tavilson"
            Case Is = "Bill Byers"
                strFolder = arkPath & "Bill H Byers"
            Case Is = "Bill Conroy"
                strFolder = arkPath & "Bill Conroy"
            Case Is = "Renae Findley"
                strFolder = arkPath & "Renae Findley"
            Case Is = "Dennis Gill"
                strFolder = arkPath & "Dennis Gill"
            Case Is = "Shawn Mockeridge"
                strFolder = arkPath & "Shawn Mockeridge"
            Case Is = "Loren Willis"
                strFolder = okcPath & "Loren Willis"
            Case Is = "ATSI Service Dept"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\ATSI"
            Case Is = "mec"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\DPS"
            Case Is = "noreply@dps.ok.gov"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\DPS"
            Case Is = "ProofPoint Essentials"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Support\Barracuda Spam Emails"
            Case Is = "lb53@sbcglobal.net"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\L-Tronics - Larry Brown"
            Case Is = "Rodriguez, Rocio - HOU"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\Peek Traffic - Oriux"
            Case Is = "Shah, Bobby - HOU"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\Peek Traffic - Oriux"
            Case Is = "Mpinkley"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\Pinkley"
            Case Is = "Lisa Pinkley"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\Pinkley"
            Case Is = "Do Not Reply"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\Econolite"
            Case Is = "Econolite"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\Econolite"
            Case Is = "Steve Wampler"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\Econolite"
            Case Is = "Repairs Mailbox"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\Econolite"
            Case Is = "Pizza Hut"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\Pizza Hut"
            Case Is = "Pizza Hut Rewards"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\Pizza Hut"
            Case Is = "IHG Rewards Club eStatement"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\Hotels & Reservations"
            Case Is = "Landon Smith"
                strFolder = okcPath & "Landon Smith"
            Case Is = "NoReplySCL1", "NoReplySCL2", "NoReplySCL3"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Locates\Responses"
            Case Is = "Amazon Business"
                strFolder = "\\ASmedley@tlsokc.com\Folders\Vendors\Amazon"
            Case Else
                strFolder = "\\ASmedley@tlsokc.com\Inbox"
        End Select
        If strFolder = "\\ASmedley@tlsokc.com\Inbox" Then
            MessageBox.Show("Folder does not exist for sender " & senderName & ".")
            senderFolder = strFolder
            Return False
        Else
            senderFolder = strFolder
            Return True
        End If
    End Function
    Public Sub MoveToSenderFolder()

        Dim xOl As Outlook.Application
        Dim NS As Outlook.NameSpace
        Dim MoveToFolder As Outlook.MAPIFolder
        Dim objItem As Object
        On Error Resume Next
        xOl = GetObject(, "Outlook.Application")
        If xOl Is Nothing Then
            xOl = CreateObject("Outlook.Application")
        End If
        On Error GoTo 0
        NS = xOl.GetNamespace("MAPI")

        If xOl.ActiveExplorer.Selection.Count = 0 Then
            MessageBox.Show("No item selected.")
        End If

        For Each objItem In xOl.ActiveExplorer.Selection
            'If objItem.Class = Outlook.OlObjectClass.olMail Then
            If objItem.subject Like "*Payroll Check Print*" Then
                MoveToFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\TLS Employees\TUL\Amanda Kay Smedley\PayStubs")
            ElseIf objItem.subject Like "*Applicant*" Or objItem.Subject Like "*Application*" Then
                MoveToFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\TLS Employees\Applicants")
            ElseIf objItem.subject Like "*Leave Used*" Or objItem.subject Like "*Leave Cancelled*" Then
                MoveToFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\TLS Employees\List Of Emps Out Today")
            ElseIf objItem.subject Like "*List of Employees Out*" Then
                MoveToFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\TLS Employees\List Of Emps Out Today")
            Else
                GetSenderFolder(objItem.SenderName)
                MoveToFolder = GetFolder(senderFolder)
            End If
            If MoveToFolder Is Nothing Then
                MessageBox.Show("Sender folder for " & objItem.SenderName & " not found!")
                releaseObject(objItem)
            Else
                objItem.UnRead = False
                objItem.Move(MoveToFolder)
                releaseObject(objItem)
            End If
            'Else
            'MessageBox.Show("Item is not a mail item.")
            'End If
        Next
        NAR(xOl)
        NAR(NS)
        'NAR(MoveToFolder)
        'NAR(objItem)
    End Sub
    Public Sub RemoveSubjectPrefix()

        Dim xOl As Outlook.Application
        Dim xItem As Object

        On Error Resume Next

        xOl = GetObject(, "Outlook.Application")

        If xOl Is Nothing Then
            xOl = CreateObject("Outlook.Application")
        End If
        On Error GoTo 0
        'xItem = GetCurrentItem()
        If TypeName(xOl.ActiveWindow) = "Explorer" Then
            For Each xItem In xOl.ActiveExplorer.Selection
                xItem.Subject = RemoveUnwantedText(xItem.Subject)
                xItem.Save()
                On Error Resume Next
                releaseObject(xItem)
                NAR(xItem)
                On Error GoTo 0
            Next xItem
        ElseIf TypeName(xOl.ActiveWindow) = "Inspector" Then
            xItem = xOl.ActiveInspector.CurrentItem
            xItem.Subject = RemoveUnwantedText(xItem.Subject)
            xItem.Save()
            On Error Resume Next
            releaseObject(xItem)
            NAR(xItem)
            On Error GoTo 0
        End If
        xItem = Nothing
        xOl = Nothing
    End Sub
    Function RemoveUnwantedText(ByVal xString As String) As String
        Dim arr(13)
        Dim i As Long
        arr(0) = "[External]"
        arr(1) = "RE:"
        arr(2) = "Re:"
        arr(3) = "re:"
        arr(4) = "FW:"
        arr(5) = "Fw:"
        arr(6) = "fw:"
        arr(7) = ".pdf"
        arr(8) = ".PDF"
        arr(9) = "Fwd:"
        arr(10) = "FWD:"
        arr(11) = "fwd:"
        arr(12) = "[External]"

        For i = LBound(arr) To UBound(arr)
            If InStr(xString, arr(i), vbTextCompare) > 0 Then
                xString = Replace(xString, arr(i), "",,, vbTextCompare)
            End If
        Next i

        RemoveUnwantedText = Trim(xString)
    End Function
    Function GetCurrentItem() As Object
        Dim xOl As Outlook.Application
        Dim strStubject As String
        Dim objItem As Object

        On Error Resume Next

        xOl = GetObject(, "Outlook.Application")
        If xOl Is Nothing Then
            xOl = CreateObject("Outlook.Application")
        End If

        Select Case TypeName(xOl.ActiveWindow)
            Case "Explorer"
                GetCurrentItem = xOl.ActiveExplorer.Selection.Item(1)
                releaseObject(xOl)
            Case "Inspector"
                GetCurrentItem = xOl.ActiveInspector.CurrentItem
                releaseObject(xOl)
        End Select
    End Function
    Public Sub Browse2Folder(ByVal folderPath As String)
        On Error Resume Next
        Dim xOl As Outlook.Application
        xOl = GetObject(, "Outlook.Application")
        If xOl Is Nothing Then
            xOl = CreateObject("Outlook.Application")
        End If
        On Error GoTo 0
        Dim myFolder As Outlook.Folder
        myFolder = GetFolder(folderPath)
        If Not (myFolder Is Nothing) Then
            xOl.ActiveExplorer.CurrentFolder = myFolder
        End If
    End Sub
    Public Sub ArchiveCompleteSJFolder(Division As String, JobType As String, Optional sYear As String = "")
        Dim oOut As Outlook.Application
        'Dim oNS As Outlook.NameSpace
        Dim destFolder As Outlook.Folder
        Dim curFolder As Outlook.MAPIFolder
        oOut = GetObject(, "Outlook.Application")
        ' oNS = oOut.GetNamespace("MAPI")
        curFolder = oOut.ActiveExplorer.CurrentFolder
        If sYear = "" Then
            sYear = Year(Now())
        End If

        destFolder = GetPSTFolder("ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\" & Division & "\" & JobType & "\" & sYear)

        If destFolder Is Nothing Then
            MessageBox.Show("The destination folder doesn't exist. You may need to add the year folder.")
            Exit Sub
        End If
        curFolder.MoveTo(destFolder)
ExitHandler:
        destFolder = Nothing
        curFolder = Nothing
        'oNS = Nothing
        oOut = Nothing
    End Sub
    Public Function GetPSTFolder(strFolderPath As String) As Outlook.MAPIFolder
        Dim objApp As Outlook.Application
        Dim objNS As Outlook.NameSpace
        Dim colFolders As Outlook.Folders
        Dim objFolder As Outlook.MAPIFolder
        Dim arrFolders() As String
        Dim i As Long
        On Error Resume Next
        strFolderPath = Replace(strFolderPath, "/", "\")
        arrFolders = Split(strFolderPath, "\")
        objApp = GetObject(, "Outlook.Application")
        objNS = objApp.GetNamespace("MAPI")
        objFolder = objNS.Folders.Item(arrFolders(0))
        If Not (objFolder Is Nothing) Then
            For i = 1 To UBound(arrFolders)
                colFolders = objFolder.Folders
                objFolder = Nothing
                objFolder = colFolders.Item(arrFolders(i))
                If objFolder Is Nothing Then
                    Exit For
                End If
            Next
        End If
        Debug.Print("objFolder: " & objFolder.Name)
        GetPSTFolder = objFolder
        colFolders = Nothing
        objNS = Nothing
        objApp = Nothing
    End Function
    Public Sub MoveToFolder(ByVal FoldersOrArchive As String, ByVal targetFolder As String)
        On Error Resume Next
        Dim xOl As Outlook.Application
        xOl = GetObject(, "Outlook.Application")
        If xOl Is Nothing Then
            xOl = CreateObject("Outlook.Application")
        End If
        On Error GoTo 0
        Dim NS As Outlook.NameSpace
        Dim MoveToFolder As Outlook.MAPIFolder
        Dim objItem As Outlook.MailItem
        Dim currentExplorer As Outlook.Explorer
        Dim Selection As Outlook.Selection
        Dim lngMovedItems As Long
        Dim objMessage As Object
        NS = xOl.GetNamespace("MAPI")
        currentExplorer = xOl.ActiveExplorer
Retry_Action:
        Selection = currentExplorer.Selection
        If FoldersOrArchive = "Folders" Then
            MoveToFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\" & targetFolder)
        ElseIf FoldersOrArchive = "Archive" Then
            MoveToFolder = GetFolder("\\Archive\" & targetFolder)
        ElseIf FoldersOrArchive = "SJArchive" Then
            MoveToFolder = GetFolder("\\ASmedley@tlsokc.com\Archive\" & targetFolder)
        ElseIf FoldersOrArchive = "LocatesArchive" Then
            MoveToFolder = GetFolder("\\Locates\" & targetFolder)
        Else
            MoveToFolder = Nothing
            Dim msgResult = MessageBox.Show("You did not specify the 'FoldersOrArchive' parameter. Please check your code.", "Parameter Error!", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2)
            If (msgResult = DialogResult.Cancel) Then
                Exit Sub
            End If
        End If

        If Selection.Count = 0 Then
            Dim msg2Result = MessageBox.Show("Select a message first.", "No Items Selected", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2)
            If (msg2Result = DialogResult.Cancel) Then
                Exit Sub
            Else
                GoTo Retry_Action
            End If
        End If

        If MoveToFolder Is Nothing Then
            MessageBox.Show("Target folder " & targetFolder & " not found!" & vbNewLine & "Please check the code.", "Target Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        For Each objMessage In Selection
            With objMessage
                On Error Resume Next
                If MoveToFolder.DefaultItemType = Outlook.OlItemType.olMailItem Then
                    Select Case .Class
                        Case OlObjectClass.olMail
                            .UnRead = False
                            .Move(MoveToFolder)
                            lngMovedItems += 1
                        Case OlObjectClass.olReport
                            .UnRead = False
                            .Move(MoveToFolder)
                            lngMovedItems += 1
                        Case OlObjectClass.olMeetingCancellation,
                         OlObjectClass.olMeetingResponseTentative,
                         OlObjectClass.olMeetingResponsePositive,
                         OlObjectClass.olMeetingResponseNegative,
                         OlObjectClass.olMeetingRequest,
                         OlObjectClass.olMeetingForwardNotification,
                        .UnRead = False
                            .Move(MoveToFolder)
                            lngMovedItems += 1
                        Case Else
                            Continue For
                    End Select
                End If
            End With
        Next

ExitHandler:
        NAR(Selection)
        NAR(currentExplorer)
        NAR(NS)
        NAR(xOl)
        'MessageBox.Show("Moved " & lngMovedItems & " message(s).")
        Exit Sub

    End Sub
    Public Sub MoveToFolder_FullPath(ByVal targetFolder As String)
        On Error Resume Next
        Dim xOl As Outlook.Application
        xOl = GetObject(, "Outlook.Application")
        If xOl Is Nothing Then
            xOl = CreateObject("Outlook.Application")
        End If
        On Error GoTo 0
        Dim NS As Outlook.NameSpace
        Dim MoveToFolder As Outlook.MAPIFolder
        Dim objItem As Outlook.MailItem
        Dim currentExplorer As Outlook.Explorer
        Dim Selection As Outlook.Selection
        Dim lngMovedItems As Long
        Dim objMessage As Object
        NS = xOl.GetNamespace("MAPI")
        currentExplorer = xOl.ActiveExplorer
Retry_Action:
        Selection = currentExplorer.Selection
        MoveToFolder = GetFolder(targetFolder)
        If Selection.Count = 0 Then
            Dim msg2Result = MessageBox.Show("Select a message first.", "No Items Selected", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2)
            If (msg2Result = DialogResult.Cancel) Then
                Exit Sub
            Else
                GoTo Retry_Action
            End If
        End If
        If MoveToFolder Is Nothing Then
            MessageBox.Show("Target folder " & targetFolder & " not found!" & vbNewLine & "Please check the code.", "Target Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If
        For Each objMessage In Selection
            With objMessage
                On Error Resume Next
                If MoveToFolder.DefaultItemType = Outlook.OlItemType.olMailItem Then
                    Select Case .Class
                        Case OlObjectClass.olMail
                            .UnRead = False
                            .Move(MoveToFolder)
                            lngMovedItems += 1
                        Case OlObjectClass.olReport
                            .UnRead = False
                            .Move(MoveToFolder)
                            lngMovedItems += 1
                        Case OlObjectClass.olMeetingCancellation,
                         OlObjectClass.olMeetingResponseTentative,
                         OlObjectClass.olMeetingResponsePositive,
                         OlObjectClass.olMeetingResponseNegative,
                         OlObjectClass.olMeetingRequest,
                         OlObjectClass.olMeetingForwardNotification,
                        .UnRead = False
                            .Move(MoveToFolder)
                            lngMovedItems += 1
                        Case Else
                            Continue For
                    End Select
                End If
            End With
        Next

ExitHandler:
        NAR(Selection)
        NAR(currentExplorer)
        NAR(NS)
        NAR(xOl)
        'MessageBox.Show("Moved " & lngMovedItems & " message(s).")
        Exit Sub

    End Sub
    Public Sub CopyToFolder(ByVal FoldersOrArchive As String, ByVal targetFolder As String)
        On Error Resume Next
        Dim xOl As Outlook.Application
        xOl = GetObject(, "Outlook.Application")
        If xOl Is Nothing Then
            xOl = CreateObject("Outlook.Application")
        End If
        On Error GoTo 0
        Dim NS As Outlook.NameSpace
        Dim CopyFolder As Outlook.MAPIFolder
        Dim objItem As Outlook.MailItem
        Dim currentExplorer As Outlook.Explorer
        Dim Selection As Outlook.Selection
        Dim lngCopiedItems As Long
        Dim objMessage As Object
        Dim copiedMessage As Object
        NS = xOl.GetNamespace("MAPI")
        currentExplorer = xOl.ActiveExplorer
Retry_Action:
        Selection = currentExplorer.Selection
        If FoldersOrArchive = "Folders" Then
            CopyFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\" & targetFolder)
        ElseIf FoldersOrArchive = "Archive" Then
            CopyFolder = GetFolder("\\ASmedley@tlsokc.com\Archive\" & targetFolder)
        Else
            CopyFolder = Nothing
            Dim msgResult = MessageBox.Show("You did not specify the 'FoldersOrArchive' parameter. Please check your code.", "Parameter Error!", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2)
            If (msgResult = DialogResult.Cancel) Then
                Exit Sub
            End If
        End If

        If Selection.Count = 0 Then
            Dim msg2Result = MessageBox.Show("Select a message first.", "No Items Selected", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2)
            If (msg2Result = DialogResult.Cancel) Then
                Exit Sub
            Else
                GoTo Retry_Action
            End If
        End If

        If CopyFolder Is Nothing Then
            MessageBox.Show("Target folder " & targetFolder & " not found!" & vbNewLine & "Please check the code.", "Target Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        For Each objMessage In Selection
            With objMessage
                On Error Resume Next
                If CopyFolder.DefaultItemType = Outlook.OlItemType.olMailItem Then
                    Select Case .Class
                        Case OlObjectClass.olMail
                            .UnRead = False
                            copiedMessage = .Copy
                            copiedMessage.Move(CopyFolder)
                            lngCopiedItems += 1
                        Case OlObjectClass.olReport
                            .UnRead = False
                            copiedMessage = .Copy
                            copiedMessage.Move(CopyFolder)
                            lngCopiedItems += 1
                        Case OlObjectClass.olMeetingCancellation,
                         OlObjectClass.olMeetingResponseTentative,
                         OlObjectClass.olMeetingResponsePositive,
                         OlObjectClass.olMeetingResponseNegative,
                         OlObjectClass.olMeetingRequest,
                         OlObjectClass.olMeetingForwardNotification,
                            .UnRead = False
                            copiedMessage = .Copy
                            copiedMessage.Move(CopyFolder)
                            lngCopiedItems += 1
                        Case Else
                            Continue For
                    End Select
                End If
            End With
        Next

ExitHandler:
        'MessageBox.Show("Moved " & lngMovedItems & " message(s).")
        Exit Sub

    End Sub
    Public Sub SendEmail(ByVal toWho As String, mySignature As String, Optional ByRef attPath As String = vbNullString, Optional ByRef ccWho As String = vbNullString)
        Dim oApp As Outlook.Application
        Dim oMail As Outlook.MailItem
        Dim strSignature, strSignatureFile As String
        Dim objTextStream, objFileSystem As Object

        strSignatureFile = CStr(Environ("USERPROFILE")) & "\AppData\Roaming\Microsoft\Signatures\My\" & mySignature & ".htm"
        objFileSystem = CreateObject("Scripting.FileSystemObject")
        objTextStream = objFileSystem.OpenTextFile(strSignatureFile)
        strSignature = objTextStream.ReadAll
        oApp = GetObject(, "Outlook.Application")
        If oApp Is Nothing Then
            oApp = CreateObject("Outlook.Application")
        End If

        oMail = oApp.CreateItem(OlItemType.olMailItem)
        With oMail
            .To = toWho
            If Not IsNothing(ccWho) Then
                .CC = ccWho
            End If
            If Not IsNothing(attPath) Then
                .Attachments.Add(attPath)
            Else
            End If
            .Recipients.ResolveAll()
            .HTMLBody = ""
            .HTMLBody = .HTMLBody & "<HTML><BODY><br>" & strSignature & "</br></BODY></HTML>"
            .Display()
        End With
    End Sub
    Public Sub SendContactList()
        Dim oApp As Outlook.Application
        Dim oMail As Outlook.MailItem
        Dim strBody, strSubject, strAtt, strAdded, strRemoved, strUpdated As String
        Dim strAdd() As String
        Dim strUp() As String
        Dim strRem() As String

        strAdded = InputBox("Name of Added:", "Newly Added", "N/A")
        If strAdded <> "N/A" Then
            strAdd = Split(strAdded, ",")
            strAdded = vbNullString
            For a = LBound(strAdd) To UBound(strAdd)
                strAdded = strAdded & strAdd(a) & "<br>"
            Next
        End If
        strRemoved = InputBox("Name of Removed:", "Newly Removed", "N/A")
        If strRemoved <> "N/A" Then
            strRem = Split(strRemoved, ",")
            strRemoved = vbNullString
            For r = LBound(strRem) To UBound(strRem)
                strRemoved = strRemoved & strRem(r) & "<br>"
            Next
        End If
        strUpdated = InputBox("Name of Updated:", "Newly Updated", "N/A")
        If strUpdated <> "N/A" Then
            strUp = Split(strUpdated, ",")
            strUpdated = vbNullString
            For u = LBound(strUp) To UBound(strUp)
                strUpdated = strUpdated & strUp(u) & "<br>"
            Next
        End If
        oApp = GetObject(, "Outlook.Application")
        If oApp Is Nothing Then
            oApp = CreateObject("Outlook.Application")
        End If
        oMail = oApp.CreateItem(OlItemType.olMailItem)
        strAtt = GetMostRecentPDF("\\TLS-FILE\TUL Administrative\Phone Listings\")
        strSubject = "Tulsa Employee Phone List Updated " & Format(Now(), "MM/dd/yyyy")
        strBody = "<HTML><BODY><p><b><u>Added:</u></b><br>" & strAdded & "</p><p><b><u>Updated:</u></b><br>" & strUpdated & "</p><p><b><u>Removed:</u></b><br>" & strRemoved & "</p>"

        With oMail
            .Display()
            .To = "TLS Cookout; cwillis@tlsokc.com; jbutler@tlsokc.com; tbond@tlsokc.com; dwillis@tlsokc.com; lbutler@tlsokc.com"
            .Recipients.ResolveAll()
            .Subject = strSubject
            .HTMLBody = strBody & .HTMLBody
            .Attachments.Add(strAtt)
            .Display()
        End With

        releaseObject(oMail)
        releaseObject(oApp)

    End Sub
    Public Sub SendDeductionEmail()
        Dim oOut As Outlook.Application
        Dim objMail As Outlook.MailItem
        Dim eTo As String
        Dim strSignature As String
        Dim objTextStream As Object
        Dim objFileSystem As Object
        Dim strSignatureFile As String
        Dim strAtt As String
        oOut = GetObject(, "Outlook.Application")
        If oOut Is Nothing Then
            oOut = CreateObject("Outlook.Application")
        End If
        strSignatureFile = CStr(Environ("USERPROFILE")) & "\AppData\Roaming\Microsoft\Signatures\My\TLS2.htm"
        objFileSystem = CreateObject("Scripting.FileSystemObject")
        objTextStream = objFileSystem.OpenTextFile(strSignatureFile)
        strSignature = objTextStream.ReadAll
        strAtt = DirSearchPDF("X:\Tulsa 2019\Employees\Current Employees\")
        MessageBox.Show(strAtt)
        objMail = oOut.CreateItem(OlItemType.olMailItem)
        eTo = "tcarothers@tlsokc.com"
        If System.IO.File.Exists(strAtt) Then
            objMail.Attachments.Add(strAtt)
        End If
        objMail.To = eTo
        objMail.Recipients.ResolveAll()
        objMail.Subject = "Deduction Request: " & InputBox("Emp Name & Date", "Ded. Req. for?")
        objMail.HTMLBody = "Deduction Request"
        objMail.HTMLBody = objMail.HTMLBody & "<HTML><BODY><br>" & strSignature & "</br></BODY></HTML>"
        objMail.Display()

    End Sub
    Public Shared Sub OpenTextFile(ByVal filePath As String)
        'verify that the file exists
        If System.IO.File.Exists(filePath) = False Then
            Console.Write("File Not Found: " & filePath)
        Else
            'Open the text file and display it's contents
            Dim sr As System.IO.StreamReader = System.IO.File.OpenText(filePath)
            Console.Write(sr.ReadToEnd)
            sr.Close()
        End If
    End Sub
    'Public Shared Sub CreateBilledDetails(ByVal entityAbbrv As String, ByVal jobNum As String)
    Public Shared Sub CreateBilledDetails(ByVal jobNum As String)
        Dim BilledTo As String
        Dim BilledDate As Date = Date.Today
        Dim BilledYear As String = Date.Today.ToString("yyyy")
        Dim myInitals As String = "AKS"
        Dim Biller As String
        Dim BillingConfirmation As String
        Dim BilledOut As Boolean
        Dim entityAbbrv As String
        Console.Write("jobNum: " & jobNum)
        Console.Write("midJobNum " & Mid(jobNum, 1, 4))
        If Mid(jobNum, 4, 1) = "9" Then
            entityAbbrv = "STI"
        Else
            entityAbbrv = "TLS"
        End If
        Console.Write("entityAbbrv=" & entityAbbrv)
        If entityAbbrv = "STI" Then
            Biller = "JOQUITA BUTLER"
        Else
            Biller = "TRACY WILLIS"
        End If

        If Left(jobNum, 1) = "B" Then
            BilledYear = Date.Today.ToString("yyyy")
        Else
            BilledYear = "20" & Mid(jobNum, 2, 2)
        End If

        BilledDate = BilledDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture)
        Dim docPath As String
        Dim fullBilledFileName As String
        Dim fullSubmittedFileName As String
        docPath = "\\TLS-FILE\TUL TLS Data\Job Folders\Tulsa\Small Jobs\" & entityAbbrv & "\" & BilledYear & "\" & jobNum & "\Bid-Billing Info\"
        If System.IO.Directory.Exists(docPath) = False Then
            docPath = "\\TLS-FILE\TUL TLS Data\Job Folders\Tulsa\Small Jobs\" & entityAbbrv & "\" & BilledYear & "\COMPLETE & BILLED\" & jobNum & "\Bid-Billing Info\"
            fullBilledFileName = Path.Combine(docPath, Convert.ToString("BILLED DETAILS - " & BilledDate.ToString("MM-dd-yyyy") & " - " & myInitals & ".txt"))
            fullSubmittedFileName = ""
            BilledOut = True
        Else
            docPath = "\\TLS-FILE\TUL TLS Data\Job Folders\Tulsa\Small Jobs\" & entityAbbrv & "\" & BilledYear & "\" & jobNum & "\Bid-Billing Info\"
            fullSubmittedFileName = Path.Combine(docPath, Convert.ToString("BILLING SUBMITTED - " & BilledDate.ToString("MM-dd-yyyy") & " - " & myInitals & ".txt"))
            fullBilledFileName = ""
            BilledOut = False
        End If

        If BilledOut = False Then
            Using billedDetails As New StreamWriter(fullSubmittedFileName)
                BilledTo = Biller & " TO BILL OUT"
                Dim lines() As String = {"EMAILED TO:", BilledTo, BilledDate, myInitals}
                For Each line As String In lines
                    billedDetails.WriteLine(line)
                Next
            End Using
        Else
            Using billedDetails As New StreamWriter(fullBilledFileName)
                BillingConfirmation = "EMAIL RECEIVED FROM: " & Biller & " STATING THAT JOB WAS BILLED TO CUSTOMER."
                Dim lines() As String = {BillingConfirmation, BilledDate, myInitals}
                For Each line As String In lines
                    billedDetails.WriteLine(line)
                Next
            End Using
        End If
    End Sub
    Public Shared Sub CreateBillingSubmittedDetails(ByVal entityAbbrv As String, ByVal jobNum As String)
        Dim BilledTo As String
        Dim BilledDate As Date = Date.Today
        Dim BilledYear As String = Date.Today.ToString("yyyy")
        Dim myInitals As String = "AKS"
        Dim docPath As String = "\\TLS-FILE\TUL TLS Data\Job Folders\Tulsa\Small Jobs\" & entityAbbrv & "\" & BilledYear & "\" & jobNum & "\Bid Info\"
        If entityAbbrv = "STI" Then
            BilledTo = "JOQUITA BUTLER TO BILL OUT"
        Else
            BilledTo = "TRACY WILLIS TO BILL OUT"
        End If
        BilledDate = BilledDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture)
        Dim lines() As String = {"EMAILED TO:", BilledTo, BilledDate, myInitals}
        Using billedDetails As New StreamWriter(Path.Combine(docPath, Convert.ToString("BILLED DETAILS - " & BilledDate.ToString("MM-dd-yyyy") & " - " & myInitals & ".txt")))
            For Each line As String In lines
                billedDetails.WriteLine(line)
            Next
        End Using
    End Sub
    Public Function GetMostRecentFile(ByVal dirPath As String) As String
        Dim fso, file, recentFile
        fso = CreateObject("Scripting.FileSystemObject")
        recentFile = Nothing
        For Each file In fso.GetFolder(dirPath).Files
            If (recentFile Is Nothing) Then
                recentFile = file
            ElseIf (file.datelastmodified > recentFile.DateLastModified) Then
                recentFile = file
            End If
        Next

        If recentFile Is Nothing Then
            Return MessageBox.Show("NO Recent Files")
        Else
            Return recentFile.Path
        End If
    End Function
    Public Function GetMostRecentPDF(ByVal dirPath As String) As String
        Dim recentFile, file, myFoundFile
        Dim dir As DirectoryInfo = New DirectoryInfo(dirPath)

        recentFile = Nothing
        myFoundFile = Nothing
        file = Nothing

        For Each file In dir.GetFiles("*.pdf")
            If (recentFile Is Nothing) Then
                recentFile = file
            ElseIf (file.LastAccessTime > recentFile.LastAccessTime) Then
                recentFile = file
            End If
        Next

        myFoundFile = CStr(recentFile.FullName)
        Return myFoundFile
    End Function
    Public Function DirSearchPDF(ByVal sDir As String) As String
        Dim d As String
        Dim f
        Dim recentFile
        recentFile = Nothing
        Try
            For Each d In Directory.GetDirectories(sDir)
                For Each f In Directory.GetFiles(d, ".pdf")
                    If (recentFile Is Nothing) Then
                        recentFile = f
                    ElseIf (f.LastAcessTime > recentFile.LastAccessTime) Then
                        recentFile = f
                    End If
                    Debug.WriteLine(recentFile.FullName)
                Next
                'DirSearchPDF(d)
            Next
            Return recentFile.FullName
            Debug.WriteLine(recentFile.FullName)
        Catch ex As System.Exception
            Debug.WriteLine(ex.Message)
            Return String.Empty
        End Try

    End Function
    Public Shared Function GetShortcutTargetFile(ByVal shortcutFilename As String) As String
        Dim myPath As String
        Dim pathOnly As String = Path.GetDirectoryName(shortcutFilename)
        Dim filenameOnly As String = Path.GetFileName(shortcutFilename)
        Dim shell As Shell = New Shell()
        Dim folder As Shell32.Folder = shell.[NameSpace](pathOnly)
        Dim folderItem As FolderItem = folder.ParseName(filenameOnly)
        If folderItem IsNot Nothing Then
            Dim link As ShellLinkObject = CType(folderItem.GetLink, ShellLinkObject)
            myPath = link.Path
        Else
            myPath = String.Empty
        End If

        Return myPath
    End Function
    Public Function GetFolder(ByVal folderPath As String) As Outlook.MAPIFolder
        Dim xOl As Outlook.Application = New Outlook.Application
        On Error Resume Next
        xOl = GetObject(, "Outlook.Application")
        On Error GoTo 0
        If xOl Is Nothing Then
            xOl = CreateObject("Outlook.Application")
        End If

        Dim NS As Outlook.NameSpace
        NS = xOl.GetNamespace("MAPI")
        Dim myFolder As Outlook.MAPIFolder
        Dim FoldersArray As Array
        Dim i As Integer

        On Error GoTo GetFolder_Error
        If Left(folderPath, 2) = "\\" Then
            folderPath = Right(folderPath, Len(folderPath) - 2)
        End If

        FoldersArray = Split(folderPath, "\")
        myFolder = NS.Folders.Item(FoldersArray(0))
        If Not myFolder Is Nothing Then
            For i = 1 To UBound(FoldersArray, 1)
                Dim subFolders As Outlook.Folders
                subFolders = myFolder.Folders
                myFolder = subFolders.Item(FoldersArray(i))
                If myFolder Is Nothing Then
                    GetFolder = Nothing
                End If
            Next
        End If
        GetFolder = myFolder
        Exit Function

GetFolder_Error:
        GetFolder = Nothing
        Exit Function

    End Function
    Public Sub AddNewFolder()
        Dim myPFName As String
        Dim myNFName As String
        Dim mySFName As String
        Dim mySF2Name As String
        Dim mySF3Name As String
        myPFName = InputBox("1st Parent (Required):", "Starting Point Folders")
        mySFName = InputBox("1st Subfolder (Optional):", "Starting Point Folders\" & myPFName, "")
        mySF2Name = InputBox("2nd Subfolder (Optional):", "Starting Point Folders\" & myPFName & "\" & mySFName, "")
        mySF3Name = InputBox("3rd Subfolder (Optional):", "Starting Point Folders\" & myPFName & "\" & mySFName & "\" & mySF2Name, "")
        If mySF3Name = "" Then
            If mySF2Name = "" Then
                If mySFName = "" Then
                    myNFName = InputBox("New Folder Name (Required):", "Under Folders\" & myPFName)
                    CreateAFolder(myPFName, myNFName)
                Else
                    myNFName = InputBox("New Folder Name (Required):", "Under Folders\" & myPFName & "\" & mySFName)
                    CreateAFolder(myPFName, myNFName, mySFName)
                End If
            Else
                myNFName = InputBox("New Folder Name (Required):", "Under Folders\" & myPFName & "\" & mySFName & "\" & mySF2Name)
                CreateAFolder(myPFName, myNFName, mySFName, mySF2Name)
            End If
        Else
            myNFName = InputBox("New Folder Name (Required):", "Under Folders\" & myPFName & "\" & mySFName & "\" & mySF2Name & "\" & mySF3Name)
            CreateAFolder(myPFName, myNFName, mySFName, mySF2Name, mySF3Name)
        End If
    End Sub
    Public Sub CreateAccountFolders()
        Dim myParentFolder As Outlook.MAPIFolder
        Dim myCurrentFolder As Outlook.Folders

        myParentFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\TLS Employees\TUL\Amanda Kay Smedley\Accounts")
        myCurrentFolder = myParentFolder.Folders
        myCurrentFolder.Add("ONG", OlDefaultFolders.olFolderInbox)
        myCurrentFolder.Add("PSOklahoma", OlDefaultFolders.olFolderInbox)
        myCurrentFolder.Add("Northstar", OlDefaultFolders.olFolderInbox)
        myCurrentFolder.Add("Cox", OlDefaultFolders.olFolderInbox)
        myCurrentFolder.Add("USPS", OlDefaultFolders.olFolderInbox)
        myCurrentFolder.Add("PikePass", OlDefaultFolders.olFolderInbox)
        myCurrentFolder.Add("Retirement", OlDefaultFolders.olFolderInbox)
        myCurrentFolder.Add("My Chart - St Francis", OlDefaultFolders.olFolderInbox)


        myParentFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\TLS Employees\TUL\Amanda Kay Smedley\Accounts\Progressive")
        myCurrentFolder = myParentFolder.Folders
        myCurrentFolder.Add("Renters", OlDefaultFolders.olFolderInbox)
        myCurrentFolder.Add("Auto", OlDefaultFolders.olFolderInbox)

        myParentFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\TLS Employees\TUL\Amanda Kay Smedley\Accounts\Banking")
        myCurrentFolder = myParentFolder.Folders
        myCurrentFolder.Add("Arvest", OlDefaultFolders.olFolderInbox)
        myCurrentFolder.Add("BarclayCard", OlDefaultFolders.olFolderInbox)
        myCurrentFolder.Add("Synchrony", OlDefaultFolders.olFolderInbox)
        myCurrentFolder.Add("Chase", OlDefaultFolders.olFolderInbox)
        myCurrentFolder.Add("CapitalOne", OlDefaultFolders.olFolderInbox)
        myCurrentFolder.Add("US Bank", OlDefaultFolders.olFolderInbox)
        myCurrentFolder.Add("Navient", OlDefaultFolders.olFolderInbox)

    End Sub
    Public Sub CreateAFolder(ByVal parentFolderName As String, ByVal newFolderName As String, Optional subFolderName As String = "", Optional subFolderName2 As String = "", Optional subFolderName3 As String = "")
        Dim myParentFolder As Outlook.MAPIFolder
        Dim myCurrentFolder As Outlook.Folders
        If subFolderName3 = "" Then
            If subFolderName2 = "" Then
                If subFolderName = "" Then
                    myParentFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\" & parentFolderName)
                    myCurrentFolder = myParentFolder.Folders
                    myCurrentFolder.Add(newFolderName, OlDefaultFolders.olFolderInbox)
                Else
                    myParentFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\" & parentFolderName & "\" & subFolderName)
                    myCurrentFolder = myParentFolder.Folders
                    myCurrentFolder.Add(newFolderName, OlDefaultFolders.olFolderInbox)
                End If
            Else
                myParentFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\" & parentFolderName & "\" & subFolderName & "\" & subFolderName2)
                myCurrentFolder = myParentFolder.Folders
                myCurrentFolder.Add(newFolderName, OlDefaultFolders.olFolderInbox)
            End If
        Else
            myParentFolder = GetFolder("\\ASmedley@tlsokc.com\Folders\" & parentFolderName & "\" & subFolderName & "\" & subFolderName2 & "\" & subFolderName3)
            myCurrentFolder = myParentFolder.Folders
            myCurrentFolder.Add(newFolderName, OlDefaultFolders.olFolderInbox)
        End If
        ribbon.Invalidate()
    End Sub
    Public Sub Response_SaveAttOnly()
        Dim olOut As Outlook.Application
        Dim fso As FileSystemObject
        Dim blnOverwrite As Boolean
        Dim sendEmailAddr As String
        Dim senderName As String
        Dim rcvdTime As String
        Dim pubTime As String
        Dim looper As Integer
        Dim plooper As Integer
        Dim oMail As Outlook.MailItem
        Dim obj As Object
        Dim mySelection As Selection
        Dim bPath As String
        Dim EmailSubject As String
        Dim saveName As String
        Dim objCount As Long
        Dim objCount2 As Long
        Dim progressForm As frmProgressBar
        Dim atmt As Attachment
        Dim atmtName As String
        Dim atmtSave As String
        Dim iForLoop As Long
        Dim iForLoop2 As Long
        Dim aForLoop As Long
        Dim objItem As Outlook.MailItem

        On Error Resume Next
        olOut = GetObject(, "Outlook.Application")
        If olOut Is Nothing Then
            olOut = CreateObject("Outlook.Application")
        End If
        On Error GoTo Err_Handler

        progressForm = New frmProgressBar()
        With progressForm
            .Show()
            .lblStatus.Text = defaultStatus
            .myProgressBar.Value = 0
        End With
        isCancelled = False

        mySelection = olOut.ActiveExplorer.Selection
        objCount2 = mySelection.Count
        objCount = 0
        progressValue = 0
        progressForm.myProgressBar.Value = progressValue
        progressForm.myProgressBar.Update()

        For Each obj In mySelection
            System.Windows.Forms.Application.DoEvents()
            If isCancelled Then
                MessageBox.Show("User cancelled at " & CStr(objCount) & " of " & CStr(objCount2) & " emails.")
                Exit Sub
            End If

            objCount = objCount + 1

            strStatus = "Processing attachment from email " & objCount & " of " & objCount2 & "..."
            progressForm.lblStatus.Text = strStatus
            progressForm.Update()

            oMail = obj

            rcvdTime = "_Rcvd" & Format(oMail.ReceivedTime, "yyMMddhhmmss")
            pubTime = "_Pub" & Format(Now(), "yyMMddhhmmss")

            strStatus = "Finding ticket number from attachment from email " & objCount & " of " & objCount2 & "..."
            progressForm.lblStatus.Text = strStatus
            progressForm.Update()

            myTicketNumber = GetTicketNumber(oMail)
            myMemberCode = GetMemberCode(oMail)

            'User Options
            blnOverwrite = False 'False = don't overwrite existing pdf, true = do overwrite
            'Path to save directory
            bPath = Path.Combine(My.Settings.LocateLandingPath & myTicketNumber & "\")
            Debug.Print("bPath: " & bPath)
            'Create directory if it does't already exist
            If Dir(bPath, vbDirectory) = vbNullString Then
                MkDir(bPath)
            End If

            If oMail.Attachments.Count > 0 Then
                strStatus = "Saving attachment(s) from email " & objCount & " of " & objCount2 & "..."
                progressForm.lblStatus.Text = strStatus
                progressForm.Update()
                For Each atmt In oMail.Attachments
                    atmtName = CleanFileName(atmt.FileName)
                    atmtSave = bPath & 2 & myTicketNumber & "_" & myMemberCode & "_" & atmtName
                    atmt.SaveAsFile(atmtSave)
                Next atmt
            End If
            oMail.Close(OlInspectorClose.olDiscard)
            progressValue = (objCount / objCount2) * 100
            progressForm.myProgressBar.Value = progressValue
            progressForm.myProgressBar.Update()
        Next obj

        progressForm.lblStatus.Text = defaultStatus
        progressForm.myProgressBar.Value = 0
        progressForm.myProgressBar.Update()
        progressForm.Close()

        If mySelection.Count > 0 Then
            For Each objItem In mySelection
                If isCancelled = False Then
                    objItem.UnRead = False
                Else
                    GoTo Exit_Handler
                End If
            Next objItem
            MoveToFolder("LocatesArchive", "Locates\Responses")
        ElseIf mySelection.Count = 0 Then
            MessageBox.Show("No items selected.")
            GoTo Exit_Handler
        End If

Exit_Handler:
        isCancelled = False
        NAR(olOut)
        NAR(mySelection)
        NAR(progressForm)

        Exit Sub
Err_Handler:
        MessageBox.Show("Err: " & Err.Number & vbNewLine & "Desc: " & Err.Description)
        GoTo Exit_Handler
    End Sub
    Public Sub Response_SaveAsPDFwAtt()
        Dim olOut As Outlook.Application
        Dim olNS As Outlook.NameSpace
        Dim oMail As Outlook.MailItem
        Dim mySelection As Selection
        Dim moveTo As Outlook.MAPIFolder
        Dim fso As New FileSystemObject
        Dim strSubject, strSaveName, sendEmailAddr, senderName, rcvdTime, pubTime, strID, bPath, emailSubject, saveName, pdfSave As String
        Dim blnOverwrite As Boolean
        Dim obj, objMailDocument, objHyperlink As Object
        Dim looper, pLooper, objCount, objCount2 As Long
        Dim progressForm As frmProgressBar
        Dim wordApp As Word.Application
        Dim wordDocs As Word.Documents
        Dim wordDoc As Word.Document
        Dim wordOpen As Boolean
        Dim atmt As Attachment
        Dim atmtName As String
        Dim atmtSave As String
        Dim objItem As Outlook.MailItem
        Dim mySelectionCount As Long
        Dim iForLoop As Long
        Dim iForLoop2 As Long
        Dim ActExp As Explorer
        On Error Resume Next
        olOut = GetObject(, "Outlook.Application")
        If olOut Is Nothing Then
            olOut = CreateObject("Outlook.Application")
        End If
        On Error GoTo 0

        progressForm = New frmProgressBar()
        With progressForm
            .Show()
            .lblStatus.Text = defaultStatus
            .myProgressBar.Value = 0
        End With
        isCancelled = False
        ActExp = olOut.ActiveExplorer
        mySelection = ActExp.Selection
        objCount2 = mySelection.Count
        objCount = 0
        progressValue = 0
        progressForm.myProgressBar.Value = progressValue
        progressForm.myProgressBar.Update()

        For Each obj In mySelection
            System.Windows.Forms.Application.DoEvents()
            If isCancelled Then
                MessageBox.Show("User cancelled at " & CStr(objCount) & " of " & CStr(objCount2) & " emails.")
                Exit Sub
            End If

            objCount = objCount + 1

            strStatus = "Processing email " & objCount & " of " & objCount2 & "..."
            progressForm.lblStatus.Text = strStatus
            progressForm.Update()

            oMail = obj

            objMailDocument = oMail.GetInspector.WordEditor
            If objMailDocument.Hyperlinks.Count > 0 Then
                strStatus = "Deleting Hyperlinks from email " & objCount & " of " & objCount2 & "..."
                progressForm.lblStatus.Text = strStatus
                For Each objHyperlink In objMailDocument.Hyperlinks
                    On Error Resume Next
                    objHyperlink.Delete
                    On Error GoTo 0
                Next
            End If

            'Get username portion of sender email address
            'sendEmailAddr = oMail.SenderEmailAddress
            'senderName = Left(sendEmailAddr, InStr(sendEmailAddr, "@") - 1)
            'Get time email was received
            rcvdTime = "_Rcvd-" & Format(oMail.ReceivedTime, "yyMMddhhmmss")
            'Get time this code was run
            pubTime = "_Pub-" & Format(Now(), "yyMMddhhmmss")
            'Get ticket number from email
            strStatus = "Finding ticket number & generating the filename from email " & objCount & " of " & objCount2 & "..."
            progressForm.lblStatus.Text = strStatus
            myTicketNumber = GetTicketNumber(oMail)
            myMemberCode = GetMemberCode(oMail)

            'User Options
            blnOverwrite = False 'False = don't overwrite existing pdf, true = do overwrite
            'Path to save directory
            bPath = Path.Combine(My.Settings.LocateLandingPath & myTicketNumber & "\")
            Debug.Print("bPath: " & bPath)
            'Create directory if it does't already exist
            If Dir(bPath, vbDirectory) = vbNullString Then
                MkDir(bPath)
            End If
            'Set save name


            saveName = 2 & myTicketNumber & "_" & myMemberCode & ".mht"

            fso = New FileSystemObject

            'increment filename if it already exists
            If blnOverwrite = False Then
                looper = 0
                Do While fso.FileExists(bPath & saveName)
                    looper = looper + 1
                    saveName = 2 & myTicketNumber & "_" & myMemberCode & rcvdTime & pubTime & "_" & Format(looper, "0000") & ".mht"

                Loop
            Else
            End If

            'save .mht file to create the pdf from word
            strStatus = "Saving response for ticket " & myTicketNumber & " from email " & objCount & " of " & objCount2 & " as .mht file..."
            progressForm.lblStatus.Text = strStatus
            oMail.SaveAs(bPath & saveName, OlSaveAsType.olMHTML)


            pdfSave = bPath & 2 & myTicketNumber & "_" & myMemberCode & ".pdf"


            If fso.FileExists(pdfSave) Then
                pLooper = 0
                Do While fso.FileExists(pdfSave)
                    pLooper = pLooper + 1
                    pdfSave = bPath & 2 & myTicketNumber & "_" & myMemberCode & rcvdTime & pubTime & "_" & Format(pLooper, "0000") & ".pdf"
                Loop
            Else
            End If

            'open word to convert the .mht to .pdf
            strStatus = "Converting response for ticket " & myTicketNumber & " from email " & objCount & " of " & objCount2 & " from .mht file to .pdf file..."
            progressForm.lblStatus.Text = strStatus

            On Error Resume Next
            wordApp = GetObject(, "Word.Application")
            On Error GoTo 0
            If wordApp Is Nothing Then
                wordApp = CreateObject("Word.Application")
                wordOpen = True
                wordApp.ScreenUpdating = False
                wordApp.DisplayAlerts = False
            End If
            'open .mht file and export to pdf
            wordDocs = wordApp.Documents
            wordDoc = wordDocs.Open(FileName:=bPath & saveName, Visible:=True)
            'wordApp.ActiveDocument.ExportAsFixedFormat(OutputFileName:=pdfSave, ExportFormat:=Word.WdExportFormat.wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=Word.WdExportOptimizeFor.wdExportOptimizeForPrint, Range:=Word.WdExportRange.wdExportAllDocument, From:=0, To:=0, Item:=Word.WdExportItem.wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:=Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:=True, UseISO19005_1:=False)
            wordDoc.ExportAsFixedFormat(OutputFileName:=pdfSave, ExportFormat:=Word.WdExportFormat.wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=Word.WdExportOptimizeFor.wdExportOptimizeForPrint, Range:=Word.WdExportRange.wdExportAllDocument, From:=0, To:=0, Item:=Word.WdExportItem.wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:=Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:=True, UseISO19005_1:=False)
            wordDoc.Close()
            NAR(wordApp)
            wordOpen = False

            'delete the .mht file
            strStatus = "Deleting the .mht file of email " & objCount & " of " & objCount2 & "..."
            progressForm.lblStatus.Text = strStatus
            Kill(bPath & saveName)

            'save attachements
            'If oMail.Attachments.Count > 0 Then
            'strStatus = "Saving attachment(s) from email " & objCount & " of " & objCount2 & "..."
            'progressForm.lblStatus.Text = strStatus
            '
            'For Each atmt In oMail.Attachments
            'atmtName = CleanFileName(atmt.FileName)
            'atmtSave = bPath & 2 & myTicketNumber & "_" & myMemberCode & rcvdTime & pubTime & "_" & atmtName
            'atmt.SaveAsFile(atmtSave)
            'Next
            'End If

            oMail.Close(OlInspectorClose.olDiscard)
            progressValue = (objCount / objCount2) * 100
            progressForm.myProgressBar.Value = progressValue
            progressForm.myProgressBar.Update()
        Next obj

        progressForm.lblStatus.Text = defaultStatus
        progressForm.myProgressBar.Value = 0
        progressForm.myProgressBar.Update()
        progressForm.Close()

        If mySelection.Count > 0 Then

            For Each objItem In mySelection
                If isCancelled = False Then
                    objItem.UnRead = False
                Else
                    GoTo Exit_Handler
                End If
            Next objItem

            MoveToFolder("LocatesArchive", "Locates\Responses")
        ElseIf mySelection.Count = 0 Then
            MessageBox.Show("No items selected.")
            GoTo Exit_Handler
        End If

Exit_Handler:
        isCancelled = False
        On Error Resume Next
        NAR(fso)
        NAR(progressForm)
        NAR(mySelection)
        NAR(olOut)
        NAR(ActExp)
        Exit Sub

Err_Handler:
        MessageBox.Show(text:="Err: " & Err.Number & vbNewLine & "Desc: " & Err.Description & vbNewLine & "Src: Response_SaveAsPDFwAtt", caption:="ERROR", buttons:=+vbOKOnly)
        Resume Exit_Handler
    End Sub
    Public Sub PRNotice_SaveAsPDF()
        Dim olOut As Outlook.Application
        Dim olNS As Outlook.NameSpace
        Dim oMail As Outlook.MailItem
        Dim mySelection As Selection
        Dim moveTo As Outlook.MAPIFolder
        Dim fso As New FileSystemObject
        Dim strSubject, strSaveName, sendEmailAddr, senderName, rcvdTime, pubTime, strID, bPath, emailSubject, saveName, pdfSave As String
        Dim blnOverwrite As Boolean
        Dim obj, objMailDocument, objHyperlink As Object
        Dim looper, pLooper, objCount, objCount2 As Long
        Dim progressForm As frmProgressBar
        Dim wordApp As Word.Application
        Dim wordDoc As Word.Document
        Dim wordOpen As Boolean
        Dim atmt As Attachment
        Dim atmtName As String
        Dim atmtSave As String
        Dim objItem As Outlook.MailItem
        Dim mySelectionCount As Long
        Dim iForLoop As Long
        Dim iForLoop2 As Long

        On Error Resume Next
        olOut = GetObject(, "Outlook.Application")
        If olOut Is Nothing Then
            olOut = CreateObject("Outlook.Application")
        End If
        On Error GoTo 0

        progressForm = New frmProgressBar()
        With progressForm
            .Show()
            .lblStatus.Text = defaultStatus
            .myProgressBar.Value = 0
        End With
        isCancelled = False

        mySelection = olOut.ActiveExplorer.Selection
        objCount2 = mySelection.Count
        objCount = 0
        progressValue = 0
        progressForm.myProgressBar.Value = progressValue
        progressForm.myProgressBar.Update()

        For Each obj In mySelection
            System.Windows.Forms.Application.DoEvents()
            If isCancelled Then
                MessageBox.Show("User cancelled at " & CStr(objCount) & " of " & CStr(objCount2) & " emails.")
                Exit Sub
            End If

            objCount = objCount + 1

            strStatus = "Processing email " & objCount & " of " & objCount2 & "..."
            progressForm.lblStatus.Text = strStatus
            progressForm.Update()

            oMail = obj

            objMailDocument = oMail.GetInspector.WordEditor
            If objMailDocument.Hyperlinks.Count > 0 Then
                strStatus = "Deleting Hyperlinks from email " & objCount & " of " & objCount2 & "..."
                progressForm.lblStatus.Text = strStatus
                For Each objHyperlink In objMailDocument.Hyperlinks
                    On Error Resume Next
                    objHyperlink.Delete
                    On Error GoTo 0
                Next
            End If

            'Get username portion of sender email address
            'sendEmailAddr = oMail.SenderEmailAddress
            'senderName = Left(sendEmailAddr, InStr(sendEmailAddr, "@") - 1)
            'Get time email was received
            rcvdTime = "_Rcvd-" & Format(oMail.ReceivedTime, "yyMMddhhmmss")
            'Get time this code was run
            pubTime = "_Pub-" & Format(Now(), "yyMMddhhmmss")
            'Get ticket number from email
            strStatus = "Finding ticket number & generating the filename from email " & objCount & " of " & objCount2 & "..."
            progressForm.lblStatus.Text = strStatus
            myTicketNumber = GetTicketNumber(oMail)
            'myMemberCode = GetMemberCode(oMail)

            'User Options
            blnOverwrite = False 'False = don't overwrite existing pdf, true = do overwrite
            'Path to save directory
            bPath = Path.Combine(My.Settings.LocateLandingPath & myTicketNumber & "\")
            Debug.Print("bPath: " & bPath)
            'Create directory if it does't already exist
            If Dir(bPath, vbDirectory) = vbNullString Then
                MkDir(bPath)
            End If
            'Set save name


            saveName = 2 & myTicketNumber & "_" & "OKIE811PRNotice" & ".mht"

            fso = New FileSystemObject

            'increment filename if it already exists
            If blnOverwrite = False Then
                looper = 0
                Do While fso.FileExists(bPath & saveName)
                    looper = looper + 1
                    saveName = 2 & myTicketNumber & "_" & "OKIE811PRNotice" & rcvdTime & pubTime & "_" & Format(looper, "0000") & ".mht"

                Loop
            Else
            End If

            'save .mht file to create the pdf from word
            strStatus = "Saving response for ticket " & myTicketNumber & " from email " & objCount & " of " & objCount2 & " as .mht file..."
            progressForm.lblStatus.Text = strStatus
            oMail.SaveAs(bPath & saveName, OlSaveAsType.olMHTML)


            pdfSave = bPath & 2 & myTicketNumber & "_" & "OKIE811PRNotice" & ".pdf"


            If fso.FileExists(pdfSave) Then
                pLooper = 0
                Do While fso.FileExists(pdfSave)
                    pLooper = pLooper + 1
                    pdfSave = bPath & 2 & myTicketNumber & "_" & "OKIE811PRNotice" & rcvdTime & pubTime & "_" & Format(pLooper, "0000") & ".pdf"
                Loop
            Else
            End If

            'open word to convert the .mht to .pdf
            strStatus = "Converting response for ticket " & myTicketNumber & " from email " & objCount & " of " & objCount2 & " from .mht file to .pdf file..."
            progressForm.lblStatus.Text = strStatus

            On Error Resume Next
            wordApp = GetObject(, "Word.Application")
            On Error GoTo 0
            If wordApp Is Nothing Then
                wordApp = CreateObject("Word.Application")
                wordOpen = True
                wordApp.ScreenUpdating = False
                wordApp.DisplayAlerts = False
            End If
            'open .mht file and export to pdf
            wordDoc = wordApp.Documents.Open(FileName:=bPath & saveName, Visible:=True)
            wordApp.ActiveDocument.ExportAsFixedFormat(OutputFileName:=pdfSave, ExportFormat:=Word.WdExportFormat.wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=Word.WdExportOptimizeFor.wdExportOptimizeForPrint, Range:=Word.WdExportRange.wdExportAllDocument, From:=0, To:=0, Item:=Word.WdExportItem.wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:=Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:=True, UseISO19005_1:=False)
            wordDoc.Close()
            releaseObject(wordApp)
            wordOpen = False

            'delete the .mht file
            strStatus = "Deleting the .mht file of email " & objCount & " of " & objCount2 & "..."
            progressForm.lblStatus.Text = strStatus
            Kill(bPath & saveName)

            'save attachements
            'If oMail.Attachments.Count > 0 Then
            'strStatus = "Saving attachment(s) from email " & objCount & " of " & objCount2 & "..."
            'progressForm.lblStatus.Text = strStatus
            '
            'For Each atmt In oMail.Attachments
            'atmtName = CleanFileName(atmt.FileName)
            'atmtSave = bPath & 2 & myTicketNumber & "_" & myMemberCode & rcvdTime & pubTime & "_" & atmtName
            'atmt.SaveAsFile(atmtSave)
            'Next
            'End If

            oMail.Close(OlInspectorClose.olDiscard)
            progressValue = (objCount / objCount2) * 100
            progressForm.myProgressBar.Value = progressValue
            progressForm.myProgressBar.Update()
        Next obj

        progressForm.lblStatus.Text = defaultStatus
        progressForm.myProgressBar.Value = 0
        progressForm.myProgressBar.Update()
        progressForm.Close()

        If mySelection.Count > 0 Then

            For Each objItem In mySelection
                If isCancelled = False Then
                    objItem.UnRead = False
                Else
                    GoTo Exit_Handler
                End If
            Next objItem

            MoveToFolder("LocatesArchive", "Locates\Responses\OKIE811 PR Notices")
        ElseIf mySelection.Count = 0 Then
            MessageBox.Show("No items selected.")
            GoTo Exit_Handler
        End If

Exit_Handler:
        isCancelled = False
        On Error Resume Next
        NAR(olOut)
        NAR(mySelection)
        NAR(fso)
        NAR(progressForm)
        Exit Sub

Err_Handler:
        MessageBox.Show(text:="Err: " & Err.Number & vbNewLine & "Desc: " & Err.Description & vbNewLine & "Src: PRNotice_SaveAsPDF", caption:="ERROR", buttons:=+vbOKOnly)
        Resume Exit_Handler
    End Sub
    Public Sub MovePRResponses()
        Dim olOut As Outlook.Application
        Dim mySelection As Selection

        On Error Resume Next
        olOut = GetObject(, "Outlook.Application")
        If olOut Is Nothing Then
            olOut = CreateObject("Outlook.Application")
        End If
        On Error GoTo 0
        mySelection = olOut.ActiveExplorer.Selection

        If mySelection.Count > 0 Then
            MoveToFolder("LocatesArchive", "Locates\Responses\OKIE811 PR Notices")
        ElseIf mySelection.Count = 0 Then
            MessageBox.Show("No items selected.")
        End If

        NAR(mySelection)
        NAR(olOut)
    End Sub
    Public Sub Ticket_SaveAsPDFwAtt()
        Dim olOut As Outlook.Application
        Dim olNS As Outlook.NameSpace
        Dim oMail As Outlook.MailItem
        Dim mySelection As Selection
        Dim fso As New FileSystemObject
        Dim strSubject, strSaveName, sendEmailAddr, senderName, rcvdTime, pubTime, strID, bPath, emailSubject, saveName, pdfSave As String
        Dim blnOverwrite As Boolean
        Dim obj, objMailDocument, objHyperlink As Object
        Dim looper, pLooper, objCount, objCount2 As Long
        Dim progressForm As frmProgressBar
        Dim wordApp As Word.Application
        Dim wordDoc As Word.Document
        Dim wordOpen As Boolean
        Dim atmt As Attachment
        Dim atmtName As String
        Dim atmtSave As String
        Dim objItem As Outlook.MailItem


        On Error Resume Next
        olOut = GetObject(, "Outlook.Application")
        If olOut Is Nothing Then
            olOut = CreateObject("Outlook.Application")
        End If
        On Error GoTo 0

        progressForm = New frmProgressBar()
        With progressForm
            .Show()
            .lblStatus.Text = defaultStatus
            .myProgressBar.Value = 0
            .Update()
        End With

        isCancelled = False

        mySelection = olOut.ActiveExplorer.Selection
        objCount2 = mySelection.Count
        objCount = 0
        progressValue = 0
        progressForm.myProgressBar.Value = progressValue
        progressForm.myProgressBar.Update()


        For Each obj In mySelection
            System.Windows.Forms.Application.DoEvents()
            If isCancelled Then
                MessageBox.Show("User cancelled at " & CStr(objCount) & " of " & CStr(objCount2) & " emails.")
                Exit Sub
            End If

            objCount = objCount + 1

            strStatus = "Processing email " & objCount & " of " & objCount2 & "..."
            progressForm.lblStatus.Text = strStatus
            progressForm.Update()

            oMail = obj
            objMailDocument = oMail.GetInspector.WordEditor
            If objMailDocument.Hyperlinks.Count > 0 Then
                strStatus = "Deleting Hyperlinks from email " & objCount & " of " & objCount2 & "..."
                progressForm.lblStatus.Text = strStatus
                progressForm.Update()
                For Each objHyperlink In objMailDocument.Hyperlinks
                    On Error Resume Next
                    objHyperlink.Delete
                    On Error GoTo 0
                Next
            End If

            'Get username portion of sender email address
            'sendEmailAddr = oMail.SenderEmailAddress
            'senderName = Left(sendEmailAddr, InStr(sendEmailAddr, "@") - 1)
            'Get time email was received
            rcvdTime = "_Rcvd-" & Format(oMail.ReceivedTime, "yyMMddhhmmss")
            'Get time this code was run
            pubTime = "_Pub-" & Format(Now(), "yyMMddhhmmss")
            'Get ticket number from email
            strStatus = "Finding ticket number & generating the filename from email " & objCount & " of " & objCount2 & "..."
            progressForm.lblStatus.Text = strStatus
            progressForm.Update()
            myTicketNumber = GetTicketNumber(oMail)

            'User Options
            blnOverwrite = False 'False = don't overwrite existing pdf, true = do overwrite
            'Path to save directory
            bPath = Path.Combine(My.Settings.LocateLandingPath & myTicketNumber & "\")
            Debug.Print("bPath: " & bPath)

            'Create directory if it does't already exist
            If Dir(bPath, vbDirectory) = vbNullString Then
                MkDir(bPath)
            End If
            'Set save name
            saveName = myTicketNumber & "_TKT_" & rcvdTime & pubTime & ".mht"
            fso = New FileSystemObject

            'increment filename if it already exists
            If blnOverwrite = False Then
                looper = 0
                Do While fso.FileExists(bPath & saveName)
                    looper = looper + 1
                    saveName = myTicketNumber & "_TKT_" & rcvdTime & pubTime & "_" & Format(looper, "0000") & ".mht"
                Loop
            Else
            End If

            'save .mht file to create the pdf from word
            strStatus = "Saving ticket " & myTicketNumber & " from email " & objCount & " of " & objCount2 & " as .mht file..."
            progressForm.lblStatus.Text = strStatus
            progressForm.Update()
            oMail.SaveAs(bPath & saveName, OlSaveAsType.olMHTML)

            pdfSave = bPath & myTicketNumber & "_TKT_" & rcvdTime & pubTime & ".pdf"

            If fso.FileExists(pdfSave) Then
                pLooper = 0
                Do While fso.FileExists(pdfSave)
                    pLooper = pLooper + 1
                    pdfSave = bPath & myTicketNumber & "_TKT_" & rcvdTime & pubTime & "_" & Format(pLooper, "0000") & ".pdf"
                Loop
            Else
            End If

            'open word to convert the .mht to .pdf
            strStatus = "Converting ticket " & myTicketNumber & " from email " & objCount & " of " & objCount2 & " from .mht file to .pdf file..."
            progressForm.lblStatus.Text = strStatus
            progressForm.Update()
            On Error Resume Next
            wordApp = GetObject(, "Word.Application")
            On Error GoTo 0
            If wordApp Is Nothing Then
                wordApp = CreateObject("Word.Application")
                wordOpen = True
                wordApp.ScreenUpdating = False
                wordApp.DisplayAlerts = False
            End If

            'open .mht file and export to pdf
            wordDoc = wordApp.Documents.Open(FileName:=bPath & saveName, Visible:=True)
            wordApp.ActiveDocument.ExportAsFixedFormat(OutputFileName:=pdfSave, ExportFormat:=Word.WdExportFormat.wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=Word.WdExportOptimizeFor.wdExportOptimizeForPrint, Range:=Word.WdExportRange.wdExportAllDocument, From:=0, To:=0, Item:=Word.WdExportItem.wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:=Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:=True, UseISO19005_1:=False)
            wordDoc.Close()
            releaseObject(wordApp)
            wordOpen = False

            'delete the .mht file
            strStatus = "Deleting the .mht file of email " & objCount & " of " & objCount2 & "..."
            progressForm.lblStatus.Text = strStatus
            progressForm.Update()
            Kill(bPath & saveName)

            'save attachements
            'If oMail.Attachments.Count > 0 Then
            'strStatus = "Saving attachment(s) from email " & objCount & " of " & objCount2 & "..."
            'progressForm.lblStatus.Text = strStatus
            'progressForm.Update()
            'For Each atmt In oMail.Attachments
            'atmtName = CleanFileName(atmt.FileName)
            'atmtSave = bPath & myTicketNumber & "_TKT_" & rcvdTime & pubTime & "_" & atmtName
            'atmt.SaveAsFile(atmtSave)
            'Next
            'End If

            oMail.Close(OlInspectorClose.olDiscard)

            progressValue = (objCount / objCount2) * 100
            progressForm.myProgressBar.Value = progressValue
            progressForm.myProgressBar.Update()
            progressForm.Update()
        Next obj

        progressForm.lblStatus.Text = defaultStatus
        progressForm.myProgressBar.Value = 0
        progressForm.myProgressBar.Update()
        progressForm.Close()

        If mySelection.Count > 0 Then

            For Each objItem In mySelection
                If isCancelled = False Then
                    objItem.UnRead = False
                Else
                    GoTo Exit_Handler
                End If
            Next objItem

            MoveToFolder("LocatesArchive", "Locates\Tickets")
        ElseIf mySelection.Count = 0 Then
            MessageBox.Show("No items selected.")
            GoTo Exit_Handler
        End If

Exit_Handler:
        isCancelled = False
        If wordOpen = True Then
            MessageBox.Show(text:="Err: Word still open in background after Ticket saves.", caption:="Non-Fatal Error", buttons:=+vbOKOnly)
        End If

        Exit Sub

Err_Handler:
        MessageBox.Show(text:="Err: " & Err.Number & vbNewLine & "Desc: " & Err.Description & vbNewLine & "Src: Ticket_SaveAsPDFwAtt", caption:="ERROR", buttons:=+vbOKOnly)
        Resume Exit_Handler
    End Sub

    Public Sub SaveUSICResponse()
        Dim olOut As Outlook.Application
        Dim fso As FileSystemObject
        Dim blnOverwrite As Boolean
        Dim sendEmailAddr As String
        Dim senderName As String
        Dim rcvdTime As String
        Dim pubTime As String
        Dim looper As Integer
        Dim plooper As Integer
        Dim oMail As Outlook.MailItem
        Dim obj As Object
        Dim mySelection As Selection
        Dim bPath As String
        Dim EmailSubject As String
        Dim saveName As String
        Dim objCount As Long
        Dim objCount2 As Long
        Dim progressForm As frmProgressBar
        Dim atmt As Attachment
        Dim atmtName As String
        Dim atmtSave As String
        Dim iForLoop As Long
        Dim aForLoop As Long
        Dim objItem As Outlook.MailItem
        Dim fileName As String
        Dim ticketNumber As String
        Dim memberCode As String
        Dim sFolders As Outlook.Folders
        Dim oFolder As Outlook.Folder
        Dim bWild As Boolean
        Dim bFound As Boolean
        Dim sFind As String
        bWild = True

        On Error Resume Next
        olOut = GetObject(, "Outlook.Application")
        If olOut Is Nothing Then
            olOut = CreateObject("Outlook.Application")
        End If
        On Error GoTo Err_Handler

        progressForm = New frmProgressBar()
        With progressForm
            .Show()
            .lblStatus.Text = defaultStatus
            .myProgressBar.Value = 0
        End With
        isCancelled = False

        mySelection = olOut.ActiveExplorer.Selection
        objCount2 = mySelection.Count
        objCount = 0
        progressValue = 0
        progressForm.myProgressBar.Value = progressValue
        progressForm.myProgressBar.Update()

        For Each obj In mySelection
            System.Windows.Forms.Application.DoEvents()
            If isCancelled Then
                MessageBox.Show("User cancelled at " & CStr(objCount) & " of " & CStr(objCount2) & " emails.")
                Exit Sub
            End If

            objCount = objCount + 1

            strStatus = "Processing attachment from email " & objCount & " of " & objCount2 & "..."
            progressForm.lblStatus.Text = strStatus
            progressForm.Update()

            oMail = obj

            memberCode = "USIC"

            'User Options
            blnOverwrite = True 'False = don't overwrite existing pdf, true = do overwrite

            If oMail.Attachments.Count > 0 Then

                strStatus = "Saving attachment(s) from email " & objCount & " of " & objCount2 & "..."
                    progressForm.lblStatus.Text = strStatus
                    progressForm.Update()
                    For Each atmt In oMail.Attachments
                        If Right(atmt.FileName, 3) = "pdf" Then
                            ticketNumber = ExtractTicketNumber(atmt.FileName)
                            fileName = "2" & ticketNumber & "_" & memberCode & ".pdf"

                            'Path to save directory
                            bPath = Path.Combine(My.Settings.LocateLandingPath & ticketNumber & "\")
                            Debug.Print("bPath: " & bPath)
                            'Create directory if it does't already exist
                            If Dir(bPath, vbDirectory) = vbNullString Then
                                MkDir(bPath)
                            End If

                            atmtName = bPath & fileName
                            atmtSave = atmtName
                            atmt.SaveAsFile(atmtSave)
                        End If
                    Next atmt

            End If
            oMail.Close(OlInspectorClose.olDiscard)
            progressValue = (objCount / objCount2) * 100
            progressForm.myProgressBar.Value = progressValue
            progressForm.myProgressBar.Update()

            progressForm.lblStatus.Text = defaultStatus
            progressForm.myProgressBar.Value = 0
            progressForm.myProgressBar.Update()
            progressForm.Close()

            If mySelection.Count > 0 Then
                If mySelection.Count > 1 Then
                    MessageBox.Show("Too many items selected.")
                    GoTo Exit_Handler
                Else
                    For Each objItem In mySelection
                        If isCancelled = False Then
                            objItem.UnRead = False
                        Else
                            GoTo Exit_Handler
                        End If
                    Next objItem

                    If SpeedUp = False Then System.Windows.Forms.Application.DoEvents()

                    MoveToFolder_FullPath("\\Locates\Locates\Responses\USIC")

                End If
            ElseIf mySelection.Count = 0 Then
                MessageBox.Show("No items selected.")
                GoTo Exit_Handler
            End If
        Next obj
Exit_Handler:
        isCancelled = False
        NAR(olOut)
        NAR(mySelection)
        NAR(progressForm)
        Exit Sub
Err_Handler:
        MessageBox.Show("Err: " & Err.Number & vbNewLine & "Desc: " & Err.Description & vbNewLine & Err.Source)
        GoTo Exit_Handler
    End Sub
    Public Sub SaveInvoice_PDF()
        Dim olOut As Outlook.Application
        Dim fso As FileSystemObject
        Dim blnOverwrite As Boolean
        Dim sendEmailAddr As String
        Dim senderName As String
        Dim rcvdTime As String
        Dim pubTime As String
        Dim looper As Integer
        Dim plooper As Integer
        Dim oMail As Outlook.MailItem
        Dim obj As Object
        Dim mySelection As Selection
        Dim bPath As String
        Dim EmailSubject As String
        Dim saveName As String
        Dim objCount As Long
        Dim objCount2 As Long
        Dim progressForm As frmProgressBar
        Dim atmt As Attachment
        Dim atmtName1 As String
        Dim atmtName2 As String
        Dim atmtSave1 As String
        Dim atmtSave2 As String
        Dim iForLoop As Long
        Dim iForLoop2 As Long
        Dim aForLoop As Long
        Dim objItem As Outlook.MailItem
        Dim fileName As String
        Dim jobNumber As String
        Dim invoiceDate As String
        Dim sFolders As Outlook.Folders
        Dim oFolder As Outlook.Folder
        Dim bWild As Boolean
        Dim bFound As Boolean
        Dim sFind As String
        bWild = True

        On Error Resume Next
        olOut = GetObject(, "Outlook.Application")
        If olOut Is Nothing Then
            olOut = CreateObject("Outlook.Application")
        End If
        On Error GoTo Err_Handler

        progressForm = New frmProgressBar()
        With progressForm
            .Show()
            .lblStatus.Text = defaultStatus
            .myProgressBar.Value = 0
        End With
        isCancelled = False

        mySelection = olOut.ActiveExplorer.Selection
        objCount2 = mySelection.Count
        objCount = 0
        progressValue = 0
        progressForm.myProgressBar.Value = progressValue
        progressForm.myProgressBar.Update()

        For Each obj In mySelection
            System.Windows.Forms.Application.DoEvents()
            If isCancelled Then
                MessageBox.Show("User cancelled at " & CStr(objCount) & " of " & CStr(objCount2) & " emails.")
                Exit Sub
            End If

            objCount = objCount + 1

            strStatus = "Processing attachment from email " & objCount & " of " & objCount2 & "..."
            progressForm.lblStatus.Text = strStatus
            progressForm.Update()

            oMail = obj

            rcvdTime = "_Rcvd" & Format(oMail.ReceivedTime, "yyMMddhhmmss")
            pubTime = "_Pub" & Format(Now(), "yyMMddhhmmss")

            jobNumber = InputBox("Job Number?")
            invoiceDate = CDate(InputBox("Invoice Date?", "Date", Date.Today)).ToString("yyyy-MM-dd")

            fileName = "Invoice " & jobNumber & ".01 - " & invoiceDate & ".pdf"
            'User Options
            blnOverwrite = True 'False = don't overwrite existing pdf, true = do overwrite

            'Path to save directory
            If Mid(jobNumber, 4, 1) = 7 Then
                bPath = Path.Combine(My.Settings.TLSJobPath & "20" & Mid(jobNumber, 2, 2) & "\Complete & Billed\" & jobNumber & "\Bid-Billing Info\")
            Else
                bPath = Path.Combine(My.Settings.STIJobPath & "20" & Mid(jobNumber, 2, 2) & "\Complete & Billed\" & jobNumber & "\Bid-Billing Info\")
            End If
            sFind = jobNumber & "*"
            Debug.Print("bPath: " & bPath)
            'Create directory if it does't already exist
            If Dir(bPath, vbDirectory) = vbNullString Then
                MkDir(bPath)
            End If

            If oMail.Attachments.Count > 0 Then
                strStatus = "Saving attachment(s) from email " & objCount & " of " & objCount2 & "..."
                progressForm.lblStatus.Text = strStatus
                progressForm.Update()
                For Each atmt In oMail.Attachments
                    If Right(atmt.FileName, 3) = "pdf" Then
                        atmtName1 = bPath & fileName
                        atmtSave1 = atmtName1
                        atmt.SaveAsFile(atmtSave1)
                    End If
                Next atmt
            End If
            oMail.Close(OlInspectorClose.olDiscard)
            progressValue = (objCount / objCount2) * 100
            progressForm.myProgressBar.Value = progressValue
            progressForm.myProgressBar.Update()

            progressForm.lblStatus.Text = defaultStatus
            progressForm.myProgressBar.Value = 0
            progressForm.myProgressBar.Update()
            progressForm.Close()

            If mySelection.Count > 0 Then
                If mySelection.Count > 1 Then
                    MessageBox.Show("Too many items selected.")
                    GoTo Exit_Handler
                Else
                    For Each objItem In mySelection
                        If isCancelled = False Then
                            objItem.UnRead = False
                        Else
                            GoTo Exit_Handler
                        End If
                    Next objItem
                    If Mid(jobNumber, 4, 1) = 7 Then
                        sFolders = GetFolder("\\ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\TUL\TLS\20" & Mid(jobNumber, 2, 2)).Folders
                    Else
                        sFolders = GetFolder("\\ASmedley@tlsokc.com\Archive\Small Jobs\Complete & Billed\TUL\STI\20" & Mid(jobNumber, 2, 2)).Folders
                    End If
                    If SpeedUp = False Then System.Windows.Forms.Application.DoEvents()
                    For Each oFolder In sFolders
                        Debug.Print(oFolder.Name)
                        bFound = (LCase(oFolder.Name) Like sFind)
                        If bFound Then
                            Debug.Print(oFolder.FolderPath)
                            MoveToFolder_FullPath(oFolder.FolderPath)
                        End If
                    Next
                End If
            ElseIf mySelection.Count = 0 Then
                MessageBox.Show("No items selected.")
                GoTo Exit_Handler
            End If
        Next obj
Exit_Handler:
        isCancelled = False
        NAR(olOut)
        NAR(mySelection)
        NAR(progressForm)
        Exit Sub
Err_Handler:
        MessageBox.Show("Err: " & Err.Number & vbNewLine & "Desc: " & Err.Description & vbNewLine & Err.Source)
        GoTo Exit_Handler
    End Sub
    Public Sub SaveTaskOrder_PDF()
        Dim olOut As Outlook.Application
        Dim fso As FileSystemObject
        Dim blnOverwrite As Boolean
        Dim sendEmailAddr As String
        Dim senderName As String
        Dim rcvdTime As String
        Dim pubTime As String
        Dim looper As Integer
        Dim plooper As Integer
        Dim oMail As Outlook.MailItem
        Dim obj As Object
        Dim mySelection As Selection
        Dim bPath1 As String
        Dim bPath2 As String
        Dim EmailSubject As String
        Dim saveName As String
        Dim objCount As Long
        Dim objCount2 As Long
        Dim progressForm As frmProgressBar
        Dim atmt As Attachment
        Dim atmtName1 As String
        Dim atmtName2 As String
        Dim atmtSave1 As String
        Dim atmtSave2 As String

        Dim iForLoop As Long
        Dim iForLoop2 As Long
        Dim aForLoop As Long
        Dim objItem As Outlook.MailItem
        Dim fileName As String
        Dim jobNumber As String
        Dim contractNumber As String
        Dim taskOrderNumber As String

        On Error Resume Next
        olOut = GetObject(, "Outlook.Application")
        If olOut Is Nothing Then
            olOut = CreateObject("Outlook.Application")
        End If
        On Error GoTo Err_Handler

        progressForm = New frmProgressBar()
        With progressForm
            .Show()
            .lblStatus.Text = defaultStatus
            .myProgressBar.Value = 0
        End With
        isCancelled = False

        mySelection = olOut.ActiveExplorer.Selection
        objCount2 = mySelection.Count
        objCount = 0
        progressValue = 0
        progressForm.myProgressBar.Value = progressValue
        progressForm.myProgressBar.Update()

        For Each obj In mySelection
            System.Windows.Forms.Application.DoEvents()
            If isCancelled Then
                MessageBox.Show("User cancelled at " & CStr(objCount) & " of " & CStr(objCount2) & " emails.")
                Exit Sub
            End If

            objCount = objCount + 1

            strStatus = "Processing attachment from email " & objCount & " of " & objCount2 & "..."
            progressForm.lblStatus.Text = strStatus
            progressForm.Update()

            oMail = obj

            rcvdTime = "_Rcvd" & Format(oMail.ReceivedTime, "yyMMddhhmmss")
            pubTime = "_Pub" & Format(Now(), "yyMMddhhmmss")

            jobNumber = InputBox("Job Number?")
            taskOrderNumber = InputBox("Task Order Number?")
            contractNumber = "BH831846"

            fileName = "Task Order " & taskOrderNumber & ", " & jobNumber & " - " & contractNumber & ".pdf"


            'User Options
            blnOverwrite = True 'False = don't overwrite existing pdf, true = do overwrite
            'Path to save directory
            bPath1 = Path.Combine(My.Settings.TaskOrderPath)
            bPath2 = Path.Combine(My.Settings.TLSJobPath & "20" & Mid(jobNumber, 2, 2) & "\" & jobNumber & "\" & "Bid-Billing Info\")
            Debug.Print("bPath1: " & bPath1)
            Debug.Print("bPath2: " & bPath2)
            'Create directory if it does't already exist
            If Dir(bPath2, vbDirectory) = vbNullString Then
                MkDir(bPath2)
            End If

            If oMail.Attachments.Count > 0 Then
                strStatus = "Saving attachment(s) from email " & objCount & " of " & objCount2 & "..."
                progressForm.lblStatus.Text = strStatus
                progressForm.Update()
                For Each atmt In oMail.Attachments
                    If Right(atmt.FileName, 3) = "pdf" Then
                        atmtName1 = bPath1 & fileName
                        atmtName2 = bPath2 & fileName
                        atmtSave1 = atmtName1
                        atmtSave2 = atmtName2
                        atmt.SaveAsFile(atmtSave1)
                        atmt.SaveAsFile(atmtSave2)
                    End If
                Next atmt
            End If
            oMail.Close(OlInspectorClose.olDiscard)
            progressValue = (objCount / objCount2) * 100
            progressForm.myProgressBar.Value = progressValue
            progressForm.myProgressBar.Update()


            progressForm.lblStatus.Text = defaultStatus
            progressForm.myProgressBar.Value = 0
            progressForm.myProgressBar.Update()
            progressForm.Close()

            If mySelection.Count > 0 Then
                If mySelection.Count > 1 Then
                    MessageBox.Show("Too many items selected.")
                    GoTo Exit_Handler
                Else
                    For Each objItem In mySelection
                        If isCancelled = False Then
                            objItem.UnRead = False
                        Else
                            GoTo Exit_Handler
                        End If
                    Next objItem
                    CopyToFolder("Folders", "Small Jobs\Transcore TO Contracts")
                    MoveToFolder("Folders", "Small Jobs\TUL\TLS\" & jobNumber & " - TransCore, LP.")
                End If
            ElseIf mySelection.Count = 0 Then
                MessageBox.Show("No items selected.")
                GoTo Exit_Handler
            End If
        Next obj
Exit_Handler:
        isCancelled = False
        NAR(olOut)
        NAR(mySelection)
        NAR(progressForm)
        Exit Sub
Err_Handler:
        MessageBox.Show("Err: " & Err.Number & vbNewLine & "Desc: " & Err.Description)
        GoTo Exit_Handler
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As System.Exception
            obj = Nothing
        End Try
    End Sub
    Public Function CleanFileName(strText As String) As String
        Dim xStripChars As String
        Dim xLen As Long
        Dim i As Long
        xStripChars = "/\[]:=+%@^*~?!," & Chr(34)
        xLen = Len(xStripChars)
        strText = Trim(strText)
        For i = 1 To xLen
            strText = Replace(strText, Mid(xStripChars, i, 1), "")
        Next
        CleanFileName = strText
    End Function
    Public Function GetMemberCode(Item As Outlook.MailItem) As String
        memCodeRegExPattern = "(OGTEAST|P66OK03|P66OK06|T0955D|T0955A|T0955B|EOIT07|((S|T){1}\d{5}))"
        ExtractMemberCode(Item.Subject)
        If Not myMemberCode = "Not Found" Then
            GetMemberCode = myMemberCode
        Else
            ExtractMemberCode(Item.Body)
            If Not myMemberCode = "Not Found" Then
                GetMemberCode = myMemberCode
            Else
                If Item.Subject Like "*for WINOK*" Then
                    GetMemberCode = "WINOK"
                Else
                    Select Case Item.SenderEmailAddress
                        Case "alcs@zlp26512.vci.att.com"
                            GetMemberCode = "T11158"
                        Case "steven.fogle@ciglocating.com"
                            GetMemberCode = "S00219"
                        Case "Dustin.Amey@nglep.com"
                            GetMemberCode = "S00624"
                        Case "KeliGreer@usicllc.com"
                            GetMemberCode = "S00445"
                        Case "waterdistrict_2@yahoo.com"
                            GetMemberCode = "T09696"
                        Case "matthew.perry@ciglocating.com"
                            GetMemberCode = "S00219"
                        Case "centralpark@collins-associates.net"
                            GetMemberCode = "T09907"
                        Case "enablemidstreamprs@korweb.com"
                            GetMemberCode = "EOIT07"
                        Case Else
                            GetMemberCode = myMemberCode
                    End Select
                End If
            End If
        End If
    End Function
    Public Function ExtractTicketNumber(strTkt As String) As String
        tktNumRegExPattern = "(\d{14})"
        Dim tktNumRegEx As New Regex(tktNumRegExPattern)
        Dim matchCol As MatchCollection = tktNumRegEx.Matches(strTkt)
        Dim myMatch As Match
        If matchCol.Count = 0 Then
            ExtractTicketNumber = "Not Found"
        Else
            myMatch = matchCol(0)
            ExtractTicketNumber = myMatch.ToString
        End If
        myTicketNumber = ExtractTicketNumber
    End Function
    Public Function ExtractMemberCode(strTkt As String) As String
        memCodeRegExPattern = "(OGTEAST|P66OK03|((S|T){1}\d{5}))"
        Dim memCodeRegEx As New Regex(memCodeRegExPattern)
        Dim matchCol As MatchCollection = memCodeRegEx.Matches(strTkt)
        Dim myMatch As Match
        If matchCol.Count = 0 Then
            ExtractMemberCode = "Not Found"
        Else
            myMatch = matchCol(0)
            ExtractMemberCode = myMatch.ToString
        End If
        myMemberCode = ExtractMemberCode
    End Function
    Public Function GetTicketNumber(Item As Outlook.MailItem) As String
        tktNumRegExPattern = "(/d{14})"
        ExtractTicketNumber(Item.Subject)
        If Not myTicketNumber = "Not Found" Then
            GetTicketNumber = myTicketNumber
        Else
            ExtractTicketNumber(Item.Body)
            If Not myTicketNumber = "Not Found" Then
                GetTicketNumber = myTicketNumber
            Else
                GetTicketNumber = "Unknown"
            End If
        End If
        Debug.Print("GetTicketNumber: " & GetTicketNumber)
    End Function
    Public Sub MyPowershell(myScriptPath As String, Optional myParameterName1 As String = vbNullString, Optional MyParameter1 As String = vbNullString, Optional MyParameterName2 As String = vbNullString, Optional MyParameter2 As String = vbNullString)
        Dim myPath As String
        If (myParameterName1 = vbNull) Or myParameterName1 = String.Empty Then
            myPath = "C:\Windows\System32\WindowsPowershell\v1.0\powershell.exe -ExecutionPolicy bypass -command " & myScriptPath
        Else
            If (MyParameterName2 = vbNull) Or MyParameterName2 = String.Empty Then
                myPath = "C:\Windows\System32\WindowsPowershell\v1.0\powershell.exe -ExecutionPolicy bypass -command " & myScriptPath & " " & myParameterName1 & " " & MyParameter1
            Else
                myPath = "C:\Windows\System32\WindowsPowershell\v1.0\powershell.exe -ExecutionPolicy bypass -command " & myScriptPath & " " & myParameterName1 & " " & MyParameter1 & " " & MyParameterName2 & " " & MyParameter2
            End If
        End If
        Process.Start(myPath, vbNormalFocus)
    End Sub

    Public oSendTo As String
    Public oSubject As String
    Public oBody As String
    Public oLocation As String
    Public oStartDate As String
    Public oStartTime As String
    Public oEndDate As String
    Public oEndTime As String
    Public oFullDay As Boolean
    Public oAtt As String

    Public Sub GeneratePTOMeeting()
        Dim oOut As Outlook.Application
        Dim oApp As Outlook.AppointmentItem
        oOut = GetObject(, "Outlook.Application")
        If oOut Is Nothing Then
            oOut = CreateObject("Outlook.Application")
        End If
        oApp = oOut.CreateItem(OlItemType.olAppointmentItem)
        With oApp
            .OptionalAttendees = "Cal Invites;" & oSendTo
            .Recipients.ResolveAll()
            .Subject = oSubject
            .Location = oLocation
            .Start = DateValue(oStartDate) & " " & TimeValue(oStartTime)
            .End = DateValue(oEndDate) & " " & TimeValue(oEndTime)
            .ReminderSet = False
            .BusyStatus = OlBusyStatus.olFree
            .Body = oBody
            .AllDayEvent = oFullDay
            .MeetingStatus = OlMeetingStatus.olMeeting
            .ResponseRequested = True
            If Len(oAtt) > 0 Then
                .Attachments.Add(oAtt)
            End If
            .Display()
        End With
    End Sub
    Public Sub GeneratePTOEmail()
        Dim oMail As Outlook.MailItem
        Dim mySignature As String
        Dim oOut As Outlook.Application
        oOut = GetObject(, "Outlook.Application")
        If oOut Is Nothing Then
            oOut = CreateObject("Outlook.Application")
        End If
        oMail = oOut.CreateItem(OlItemType.olMailItem)
        With oMail
            .Display()
        End With
        mySignature = oMail.HTMLBody
        With oMail
            .To = oSendTo
            .Recipients.ResolveAll()
            .Subject = oSubject
            .BodyFormat = OlBodyFormat.olFormatHTML
            .HTMLBody = vbNewLine & vbNewLine & mySignature
            If Len(oAtt) > 0 Then
                .Attachments.Add(oAtt)
            End If
            .Save()
            .Display()
        End With
    End Sub
    Public Function CheckPTOExists(ByVal CheckData As VariantType) As Boolean
        Dim oOut As Outlook.Application
        oOut = GetObject(, "Outlook.Application")
        If oOut Is Nothing Then
            oOut = CreateObject("Outlook.Application")
        End If
        Dim oCal As Outlook.MAPIFolder
        oCal = oOut.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar)
        Dim oApp As Outlook.AppointmentItem
        oApp = oOut.CreateItem(OlItemType.olAppointmentItem)
        For Each oApp In oCal.Items
            If oApp.Class = OlObjectClass.olAppointment Then
                If oApp.Subject = CheckData Then
                    CheckPTOExists = True
                    oApp.Display()
                    Exit For
                Else
                    CheckPTOExists = False
                End If
            Else
                CheckPTOExists = False
            End If
        Next
        Return CheckPTOExists
    End Function
    Public Function fOSUserName() As String
        Dim wshNet As Object
        wshNet = CreateObject("WScript.Network")
        fOSUserName = UCase(Left(wshNet.UserName, 2)) & Mid(wshNet.UserName, 3, Len(wshNet.UserName))
    End Function
#End Region

End Class

Friend Class PictureConverter
    Inherits AxHost

    Private Sub New()
        MyBase.New(String.Empty)
    End Sub

    Public Shared Function ImageToPictureDisp(ByVal image As Image) As stdole.IPictureDisp
        Return CType(GetIPictureDispFromPicture(image), stdole.IPictureDisp)
    End Function

    Public Shared Function IconToPictureDisp(ByVal icon As Icon) As stdole.IPictureDisp
        Return ImageToPictureDisp(icon.ToBitmap())
    End Function

    Public Shared Function PictureDispToImage(ByVal picture As stdole.IPictureDisp) As Image
        Return GetPictureFromIPicture(picture)
    End Function
End Class