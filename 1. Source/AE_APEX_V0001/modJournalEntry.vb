Module modJournalEntry

    Sub Validation(DataView_ As DataView)
        Dim p_dDebitAmount As Decimal = 0.0
        Dim p_dCreditAmount As Decimal = 0.0
        Dim sFuncName As String = String.Empty

        Try

            sFuncName = "Validation"
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            For IntK As Integer = 0 To DataView_.Count - 1

                AccountCodeChecking(DataView_.Item(IntK).Row(1).ToString.Trim, DataView_.Item(IntK).Row(0))

                If DataView_.Item(IntK).Row(5) >= 0 Then
                    p_dDebitAmount += CDbl(DataView_.Item(IntK).Row(5))
                Else
                    p_dCreditAmount -= CDbl(DataView_.Item(IntK).Row(5))
                End If
                p_oSBOApplication.StatusBar.SetText("Validating the Account Code, Credit & Debit Amount Check ...... !" & " Account Code :  " & DataView_.Item(IntK).Row(1) & "    Reference No " & DataView_.Item(0).Row(6) & "  ---  " & IntK, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Next

            If (p_dDebitAmount - p_dCreditAmount) <> 0 Then
                p_oSBOApplication.StatusBar.SetText(p_dDebitAmount & "  " & p_dCreditAmount & "   " & p_dDebitAmount - p_dCreditAmount, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                p_sRefNuber(p_iArrayCount, 0) = p_dDebitAmount
                p_sRefNuber(p_iArrayCount, 1) = p_dCreditAmount
                p_sRefNuber(p_iArrayCount, 2) = CDbl(p_dDebitAmount) - CDbl(p_dCreditAmount)
                p_sRefNuber(p_iArrayCount, 3) = DataView_.Item(0).Row(6)
                p_iArrayCount = +1
            End If

            ' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR ", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub

    Function JournalEntryPosting(DGV As DataView, JEDate As String, Memo As String, ByRef sErrDesc As String) As Long
        Dim lRetCode As Long
        Dim sJE As String = String.Empty
        Dim sGetAccount As String = String.Empty
        Dim dTaxDate As Date
        Dim oBP As SAPbobsCOM.BusinessPartners = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        Dim oProgress As SAPbouiCOM.ProgressBar = Nothing
        oProgress = p_oSBOApplication.StatusBar.CreateProgressBar("Fetching the information for Reference  ", 1000, True)


        Dim sJEDate As String = GetDate(JEDate, p_oDICompany)
        Dim dJEDocDate As Date = Left(sJEDate, 4) & "/" & Mid(sJEDate, 5, 2) & "/" & Right(sJEDate, 2)

        Dim sFuncName As String = String.Empty
        Try
            Dim oJournalEntry As SAPbobsCOM.JournalEntries = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            sFuncName = "JournalEntryPosting"
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            oJournalEntry.ReferenceDate = dJEDocDate
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Document - Posting Date " & dJEDocDate, sFuncName)
            oJournalEntry.DueDate = dJEDocDate
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Document - Due Date " & dJEDocDate, sFuncName)
            oJournalEntry.TaxDate = dJEDocDate
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Document - Tax Date " & dJEDocDate, sFuncName)
            oJournalEntry.Memo = Memo
            oJournalEntry.Reference = DGV.Item(0).Row(2).ToString
            oJournalEntry.Reference2 = DGV.Item(0).Row(6).ToString

            For IntRow As Integer = 0 To DGV.Count - 1

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSingleValue Function ", sFuncName)

                sGetAccount = GetSingleValue(DGV.Item(IntRow).Row(1).ToString, DGV.Item(IntRow).Row(0).ToString)

                oJournalEntry.Lines.ShortName = sGetAccount
                oJournalEntry.Lines.AccountCode = sGetAccount

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With Success GetSingleValue Function ", sFuncName)
                dTaxDate = ConvertStringToDate(DGV.Item(IntRow).Row(3))
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Line - Tax Date " & dTaxDate, sFuncName)
                oJournalEntry.Lines.TaxDate = dTaxDate

                oJournalEntry.Lines.Reference1 = DGV.Item(IntRow).Row(4).ToString

                oJournalEntry.Lines.DueDate = dJEDocDate
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Line - Due Date " & dJEDocDate, sFuncName)
                oJournalEntry.Lines.ReferenceDate1 = dJEDocDate
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Line - Posting Date " & dJEDocDate, sFuncName)

                If CDbl(DGV.Item(IntRow).Row(5)) > 0 Then
                    oJournalEntry.Lines.Debit = CDbl(DGV.Item(IntRow).Row(5))
                Else
                    oJournalEntry.Lines.Credit = Replace(CDbl(DGV.Item(IntRow).Row(5)), "-", "")
                End If

                oJournalEntry.Lines.Add()
                '  p_oSBOApplication.StatusBar.SetText("Fetching the information for Reference  " & DGV.Item(IntRow).Row(6) & " ---  " & IntRow, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                ' p_oSBOApplication.StatusBar.CreateProgressBar("Fetching the information for Reference  " & DGV.Item(IntRow).Row(6) & " ---  " & IntRow, 10, True)
                oProgress.Value = IntRow + 20
                oProgress.Text = "Fetching the information for Reference " & DGV.Item(IntRow).Row(6).ToString & " Line " & IntRow
            Next

            lRetCode = oJournalEntry.Add()
           
            If lRetCode <> 0 Then
                oProgress.Text = "Error while generating Journal Entry for the reference no. " & DGV.Item(0).Row(6) & " Error is " & p_oDICompany.GetLastErrorDescription
                p_oSBOApplication.MessageBox("Error while generating Journal Entry for the reference no. " & DGV.Item(0).Row(6) & " -- Error is " & p_oDICompany.GetLastErrorDescription, 1, "Ok")
                Call WriteToLogFile("Completed with ERROR ---" & p_oDICompany.GetLastErrorDescription, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & p_oDICompany.GetLastErrorDescription, sFuncName)
                JournalEntryPosting = RTN_ERROR
            Else
                p_oDICompany.GetNewObjectCode(sJE)
                oProgress.Text = "Journal Entry has been successfully generated for the reference no. " & DGV.Item(0).Row(6)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Journal Entry DocEntry " & sJE, sFuncName)
                JournalEntryPosting = RTN_SUCCESS

            End If

        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR " & ex.Message, sFuncName)
            JournalEntryPosting = RTN_ERROR
            Exit Function
        Finally
            oProgress.Stop()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgress)
            oProgress = Nothing
        End Try

    End Function

    Sub AccountCodeChecking(ByVal AcctCode As String, ByVal GDC As String)
        Try
            Dim sFuncName As String = String.Empty
            Dim sSqlString As String = String.Empty
            Dim sGLAccount As String = String.Empty

            sFuncName = "AccountCodeChecking"

            Dim oRS As SAPbobsCOM.Recordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Select Case GDC
                Case "G"
                    sSqlString = "SELECT T0.U_BibbySGCode [Name] FROM [dbo].[@BIBBY_ACCT_MAPPING]  T0 WHERE T0.U_BibbyAFCode ='" & AcctCode & "'"
                Case Else
                    sSqlString = "SELECT T0.U_BibbySGCode [Name] FROM [dbo].[@BIBBY_ACCT_MAPPING]  T0 WHERE T0.U_BibbyAFCode ='" & GDC & "'"
            End Select

            oRS.DoQuery(sSqlString)
            ' Checking for - CSV Account Account code are not associate with the mapping table this if condition will triggers and execute the true part
            If oRS.RecordCount = 0 Then

                If GDC = "C" Then
                    AcctCode = "Creditor Account Code"
                ElseIf GDC = "D" Then
                    AcctCode = "Debitor Account Code"
                End If

                If p_sAccountCodes.Contains(AcctCode) = False Then
                    p_sAccountCodes(p_iArrayAcctCount) = AcctCode
                    p_iArrayAcctCount += 1
                End If
            Else
                ' Checking for - Mapping table G/L Accounts are exist in the SAP Chart of Accounts or the account is not an Active Account
                sGLAccount = oRS.Fields.Item("Name").Value
                oRS.DoQuery("SELECT T0.[AcctCode] FROM OACT T0 WHERE T0.[AcctCode] = '" & sGLAccount & "' and T0.[Postable] = 'Y'")

                If oRS.RecordCount = 0 Then
                    If p_sAccountCodes_ActiveAccount.Contains(sGLAccount) = False Then
                        p_sAccountCodes_ActiveAccount(p_iArrayAcctActiveCount) = sGLAccount
                        p_iArrayAcctActiveCount += 1
                    End If
                End If
            End If

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR ", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub


    

End Module
