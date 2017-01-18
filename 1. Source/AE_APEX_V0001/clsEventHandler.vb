Option Explicit On
'Imports SAPbouiCOM.Framework
Imports System.Windows.Forms


Public Class clsEventHandler
    Dim WithEvents SBO_Application As SAPbouiCOM.Application ' holds connection with SBO
    Dim p_oDICompany As New SAPbobsCOM.Company

    Public Sub New(ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company)
        Dim sFuncName As String = String.Empty
        Try
            sFuncName = "Class_Initialize()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO Application handle", sFuncName)
            SBO_Application = oApplication
            p_oDICompany = oCompany

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Call WriteToLogFile(exc.Message, sFuncName)
        End Try
    End Sub

    Public Function SetApplication(ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function   :    SetApplication()
        '   Purpose    :    This function will be calling to initialize the default settings
        '                   such as Retrieving the Company Default settings, Creating Menus, and
        '                   Initialize the Event Filters
        '               
        '   Parameters :    ByRef sErrDesc AS string
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "SetApplication()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetMenus()", sFuncName)
            If SetMenus(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetFilters()", sFuncName)
            If SetFilters(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SetApplication = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(exc.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            SetApplication = RTN_ERROR
        End Try
    End Function

    Private Function SetMenus(ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function   :    SetMenus()
        '   Purpose    :    This function will be gathering to create the customized menu
        '               
        '   Parameters :    ByRef sErrDesc AS string
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sFuncName As String = String.Empty
        ' Dim oMenuItem As SAPbouiCOM.MenuItem
        Try
            sFuncName = "SetMenus()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SetMenus = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            SetMenus = RTN_ERROR
        End Try
    End Function

    Private Function SetFilters(ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function   :    SetFilters()
        '   Purpose    :    This function will be gathering to declare the event filter 
        '                   before starting the AddOn Application
        '               
        '   Parameters :    ByRef sErrDesc AS string
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************

        Dim oFilters As SAPbouiCOM.EventFilters
        Dim oFilter As SAPbouiCOM.EventFilter
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "SetFilters()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing EventFilters object", sFuncName)
            oFilters = New SAPbouiCOM.EventFilters

           

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding filters", sFuncName)
            SBO_Application.SetFilter(oFilters)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            SetFilters = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            SetFilters = RTN_ERROR
        End Try
    End Function

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
        ' **********************************************************************************
        '   Function   :    SBO_Application_AppEvent()
        '   Purpose    :    This function will be handling the SAP Application Event
        '               
        '   Parameters :    ByVal EventType As SAPbouiCOM.BoAppEventTypes
        '                       EventType = set the SAP UI Application Eveny Object        
        ' **********************************************************************************
        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty
        Dim sMessage As String = String.Empty

        Try
            sFuncName = "SBO_Application_AppEvent()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                    sMessage = String.Format("Please wait for a while to disconnect the AddOn {0} ....", System.Windows.Forms.Application.ProductName)
                    p_oSBOApplication.SetStatusBarMessage(sMessage, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    End
            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ShowErr(sErrDesc)
        Finally
            GC.Collect()  'Forces garbage collection of all generations.
        End Try
    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        ' **********************************************************************************
        '   Function   :    SBO_Application_MenuEvent()
        '   Purpose    :    This function will be handling the SAP Menu Event
        '               
        '   Parameters :    ByRef pVal As SAPbouiCOM.MenuEvent
        '                       pVal = set the SAP UI MenuEvent Object
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set the True/False        
        ' **********************************************************************************
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sErrDesc As String = String.Empty
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "SBO_Application_MenuEvent()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If Not p_oDICompany.Connected Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case "JE"
                        Try
                            LoadFromXML("JournalEntry.srf", SBO_Application)
                            oForm = SBO_Application.Forms.Item("JournalE")

                            oForm.Visible = True
                            Exit Try

                        Catch ex As Exception
                            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                        End Try
                        Exit Sub

                End Select
            End If
         
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            BubbleEvent = False
            ShowErr(exc.Message)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            WriteToLogFile(Err.Description, sFuncName)
        End Try
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, _
            ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        ' **********************************************************************************
        '   Function   :    SBO_Application_ItemEvent()
        '   Purpose    :    This function will be handling the SAP Menu Event
        '               
        '   Parameters :    ByVal FormUID As String
        '                       FormUID = set the FormUID
        '                   ByRef pVal As SAPbouiCOM.ItemEvent
        '                       pVal = set the SAP UI ItemEvent Object
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set the True/False        
        ' **********************************************************************************

        Dim sErrDesc As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim p_oDVJE As DataView = Nothing
        Dim oDTDistinct As DataTable = Nothing
        Dim oDTRowFilter As DataTable = Nothing

        Try
            sFuncName = "SBO_Application_ItemEvent()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If Not IsNothing(p_oDICompany) Then
                If Not p_oDICompany.Connected Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                    If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If
            End If

            If pVal.BeforeAction = False Then

                Select Case pVal.FormUID
                    Case "JournalE"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "Item_8" Then
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                sFuncName = "'Browse' Button Click - ID 'Item_8'"
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling File Open Function", sFuncName)
                                oForm.Items.Item("Item_5").Specific.string = fillopen()
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With Success File Open Function", sFuncName)
                                ' oForm.Items.Item("Item_5").Specific.string = p_sSelectedFilepath
                                Exit Sub
                            End If

                            If pVal.ItemUID = "Item_6" Then
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                Dim oDataview_Tmp As DataView = Nothing
                                sFuncName = "'Create JE' Button Click - ID 'Item_6'"

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetDataViewFromCSV Function", sFuncName)
                                p_oDVJE = GetDataViewFromCSV(oForm.Items.Item("Item_5").Specific.string, p_sSelectedFileName)

                                If p_oDVJE Is Nothing Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Datas in the CSV file", sFuncName)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Success GetDataViewFromCSV Function", sFuncName)
                                oDTDistinct = p_oDVJE.Table.DefaultView.ToTable(True, "F7")

                                ReDim p_sRefNuber(oDTDistinct.Rows.Count, 4)
                                ReDim p_sAccountCodes(p_oDVJE.Count)
                                ReDim p_sAccountCodes_ActiveAccount(p_oDVJE.Count)
                                p_iArrayAcctCount = 0
                                p_iArrayCount = 0
                                p_iArrayAcctActiveCount = 0

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Validation Function", sFuncName)

                                For IntRow As Integer = 1 To oDTDistinct.Rows.Count - 1
                                    p_oDVJE.RowFilter = "F7 = '" & oDTDistinct.Rows(IntRow).Item(0).ToString & "'"
                                    Validation(p_oDVJE)
                                Next
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With Success Validation Function", sFuncName)

                                If Not String.IsNullOrEmpty(p_sAccountCodes(0)) Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AccNumbers do not have a corresponding SAP G/L Account in the mapping table", sFuncName)
                                    Write_TextFile_Account(p_sAccountCodes)
                                    BubbleEvent = False
                                    ' sCursor.Current = Windows.Forms.Cursors.Default
                                    Exit Sub
                                End If

                                If Not String.IsNullOrEmpty(p_sAccountCodes_ActiveAccount(0)) Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AccNumbers do not exist in SAP G/L Account table or inactive accounts", sFuncName)
                                    Write_TextFile_ActiveAccount(p_sAccountCodes_ActiveAccount)
                                    BubbleEvent = False
                                    ' sCursor.Current = Windows.Forms.Cursors.Default
                                    Exit Sub
                                End If

                                If Not String.IsNullOrEmpty(p_sRefNuber(0, 0)) Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("The Total Debit is not equal to the Total Credit", sFuncName)
                                    Write_TextFile_Amount(p_sRefNuber)
                                    BubbleEvent = False
                                    'sCursor.Current = Windows.Forms.Cursors.Default
                                    Exit Sub
                                Else
                                    'sCursor.Current = Windows.Forms.Cursors.Default
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Start the SAP Transaction", sFuncName)
                                    If Not p_oDICompany.InTransaction = True Then p_oDICompany.StartTransaction()

                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling JournalEntryPosting Function", sFuncName)

                                    For IntRow As Integer = 1 To oDTDistinct.Rows.Count - 1
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Filtering data with respective reference -  " & oDTDistinct.Rows(IntRow).Item(0).ToString, sFuncName)

                                        p_oDVJE.RowFilter = "F7 = '" & oDTDistinct.Rows(IntRow).Item(0).ToString & "'"
                                        If JournalEntryPosting(p_oDVJE, oForm.Items.Item("Item_3").Specific.string, oForm.Items.Item("Item_4").Specific.string, sErrDesc) = RTN_ERROR Then
                                            If p_oDICompany.InTransaction Then p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With Error JournalEntryPosting Function", sFuncName)
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback transaction on company database ", sFuncName)
                                            WriteToLogFile("Completed With Error JournalEntry Posting Function", sFuncName)
                                            WriteToLogFile("Rollback transaction on company database", sFuncName)
                                            Exit Sub
                                        End If
                                    Next
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Success JournalEntryPosting Function", sFuncName)

                                    If p_oDICompany.InTransaction Then p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                    p_oSBOApplication.MessageBox("Journal Entry has been successfully generated ....... !", 1, "Ok")
                                    oForm.Items.Item("Item_3").Specific.String = String.Empty
                                    oForm.Items.Item("Item_4").Specific.String = String.Empty
                                    oForm.Items.Item("Item_5").Specific.String = String.Empty
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Transaction has been committed successfully ", sFuncName)
                                End If
                                Exit Sub
                            End If
                        End If
                End Select
            Else
                Select Case pVal.FormUID
                    Case "JournalE"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "Item_6" Then
                                Dim oForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Head Validation Function", sFuncName)

                                If HeaderValidation(oForm, sErrDesc) = 0 Then
                                    BubbleEvent = False
                                    Exit Sub
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Header Validation Function", sFuncName)
                                End If
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Header Validation Function", sFuncName)
                                Exit Sub
                            End If
                        End If
                End Select
            End If


            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            BubbleEvent = False
            sErrDesc = exc.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            WriteToLogFile(Err.Description, sFuncName)
            ShowErr(sErrDesc)
        End Try

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
