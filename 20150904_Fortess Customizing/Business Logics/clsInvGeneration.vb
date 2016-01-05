Public Class clsInvGeneration
    Inherits clsBase
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox, oCombobox1, oCombobox2, oCombobox3 As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private ocombo As SAPbouiCOM.ComboBoxColumn
    Private oCheckbox1 As SAPbouiCOM.CheckBox
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private oTemp As SAPbobsCOM.Recordset
    Private InvBaseDocNo, strname As String
    Private InvForConsumedItems As Integer
    Private oMenuobject As Object
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_InvGeneration, frm_InvGeneration)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        FillItemGroup(oForm)
        AddChooseFromList(oForm)

        oForm.DataSources.UserDataSources.Add("Reqno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oApplication.Utilities.setUserDatabind(oForm, "14", "Reqno")
        oEditText = oForm.Items.Item("14").Specific
        oEditText.ChooseFromListUID = "CFL1"
        oEditText.ChooseFromListAlias = "CardCode"

        oForm.DataSources.UserDataSources.Add("Reqno1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oApplication.Utilities.setUserDatabind(oForm, "16", "Reqno1")
        oEditText = oForm.Items.Item("16").Specific
        oEditText.ChooseFromListUID = "CFL2"
        oEditText.ChooseFromListAlias = "CardCode"

        oForm.DataSources.UserDataSources.Add("EndFrom", SAPbouiCOM.BoDataType.dt_DATE)
        oApplication.Utilities.setUserDatabind(oForm, "10", "EndFrom")
        oForm.DataSources.UserDataSources.Add("EndTo", SAPbouiCOM.BoDataType.dt_DATE)
        oApplication.Utilities.setUserDatabind(oForm, "1000002", "EndTo")

        oForm.DataSources.UserDataSources.Add("WaType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oCombobox = oForm.Items.Item("18").Specific
        oCombobox.DataBind.SetBound(True, "", "WaType")

        For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
            oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        oCombobox.ValidValues.Add("", "")
        oCombobox.ValidValues.Add("F", "Free")
        oCombobox.ValidValues.Add("W", "Warranty")
        oCombobox.ValidValues.Add("R", "Regular")
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        ' oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly


        oForm.DataSources.UserDataSources.Add("Status", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oCombobox = oForm.Items.Item("22").Specific
        oCombobox.DataBind.SetBound(True, "", "Status")

        For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
            oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        oCombobox.ValidValues.Add("", "")
        oCombobox.ValidValues.Add("D", "Draft")
        oCombobox.ValidValues.Add("F", "On Hold")
        oCombobox.ValidValues.Add("T", "Terminated")
        oCombobox.ValidValues.Add("A", "Approved")
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        'oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly


        oForm.DataSources.UserDataSources.Add("Issued", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oCombobox = oForm.Items.Item("24").Specific
        oCombobox.DataBind.SetBound(True, "", "Issued")

        For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
            oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        oCombobox.ValidValues.Add("", "")
        oCombobox.ValidValues.Add("Y", "Yes")
        oCombobox.ValidValues.Add("N", "No")
       
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        ' oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

        oForm.PaneLevel = 1
        oForm.Freeze(False)
    End Sub
    Private Sub FillItemGroup(ByVal aForm As SAPbouiCOM.Form)
        oCombobox = aForm.Items.Item("20").Specific
        Dim oSlpRS As SAPbobsCOM.Recordset
        oSlpRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oSlpRS.DoQuery("select ItmsGrpCod,ItmsGrpNam  from OITB order by ItmsGrpCod")
        oCombobox.ValidValues.Add("", "")
        For intRow As Integer = 0 To oSlpRS.RecordCount - 1
            oCombobox.ValidValues.Add(oSlpRS.Fields.Item(0).Value, oSlpRS.Fields.Item(1).Value)
            oSlpRS.MoveNext()
        Next
        aForm.Items.Item("20").DisplayDesc = True
    End Sub
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"

            oCFL.SetConditions(oCons)


            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL2"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"

            oCFL.SetConditions(oCons)


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try

            Dim EndFrom, EndTo, custFrom, custTo As String
            oCombobox = oForm.Items.Item("18").Specific
            oCombobox1 = oForm.Items.Item("20").Specific
            oCombobox2 = oForm.Items.Item("22").Specific
            oCombobox3 = oForm.Items.Item("24").Specific

            EndFrom = oApplication.Utilities.getEdittextvalue(aForm, "10")
            EndTo = oApplication.Utilities.getEdittextvalue(aForm, "1000002")
            custFrom = oApplication.Utilities.getEdittextvalue(aForm, "14")
            custTo = oApplication.Utilities.getEdittextvalue(aForm, "16")
            'If EndFrom = "" And EndTo = "" And custFrom = "" And custTo = "" And oCombobox.Selected.Value = "" And oCombobox1.Selected.Value = "" And oCombobox2.Selected.Value = "" And oCombobox3.Selected.Value = "" Then
            '    ' oApplication.Utilities.Message("Select any one...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    ' Return False
            'End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

#Region "Databind"
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim strstring, EndFrom, EndTo, CustFrom, custTo, Warranty, Itemgroup, ConStatus, Invoice, WarrantyCondition, ItemGroupCondition, StatusCondition, strCondition, InvoiceCondition As String
            Dim Endfrdt, EndTodt As Date
            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCombobox = aform.Items.Item("18").Specific
            oCombobox1 = aform.Items.Item("20").Specific
            oCombobox2 = aform.Items.Item("22").Specific
            oCombobox3 = aform.Items.Item("24").Specific
            EndFrom = oApplication.Utilities.getEdittextvalue(aform, "10")
            EndTo = oApplication.Utilities.getEdittextvalue(aform, "1000002")
            CustFrom = oApplication.Utilities.getEdittextvalue(aform, "14")
            custTo = oApplication.Utilities.getEdittextvalue(aform, "16")
            Warranty = oCombobox.Selected.Value
            Itemgroup = oCombobox1.Selected.Value
            ConStatus = oCombobox2.Selected.Value
            Invoice = oCombobox3.Selected.Value
            If oCombobox.Selected.Value <> "" Then
                WarrantyCondition = "T0.U_Z_WarType='" & oCombobox.Selected.Value & "'"
            Else
                WarrantyCondition = "1=1"
            End If
            If oCombobox1.Selected.Value <> "" Then
                ItemGroupCondition = "T1.ItemGroup='" & oCombobox1.Selected.Value & "'"
            Else
                ItemGroupCondition = "1=1"
            End If
            If oCombobox2.Selected.Value <> "" Then
                StatusCondition = "T0.Status='" & oCombobox2.Selected.Value & "'"
            Else
                StatusCondition = "1=1"
            End If
            If oCombobox3.Selected.Value <> "" Then
                InvoiceCondition = "isnull(T0.U_Z_Invoice,'N')='" & oCombobox3.Selected.Value & "'"
            Else
                InvoiceCondition = "1=1"
            End If
            If EndFrom <> "" Then
                Endfrdt = oApplication.Utilities.GetDateTimeValue(EndFrom)
            End If
            If EndTo <> "" Then
                EndTodt = oApplication.Utilities.GetDateTimeValue(EndTo)
            End If

            Dim strDateCondition1 As String = ""
            Dim strCustCondition As String = ""

            If EndFrom <> "" And EndTo <> "" Then
                strDateCondition1 = "T0.EndDate > = '" & Endfrdt.ToString("yyyy-MM-dd") & "' and T0.EndDate <= '" & EndTodt.ToString("yyyy-MM-dd") & "'"
            ElseIf EndFrom <> "" And EndTo = "" Then
                strDateCondition1 = "T0.EndDate >= '" & Endfrdt.ToString("yyyy-MM-dd") & "'"
            ElseIf EndFrom = "" And EndTo <> "" Then
                strDateCondition1 = "T0.EndDate <= '" & EndTodt.ToString("yyyy-MM-dd") & "'"
            Else
                strDateCondition1 = " 1=1"
            End If

            If CustFrom <> "" And custTo <> "" Then
                strCustCondition = " CstmrCode between '" & CustFrom & "' and '" & custTo & "'"
            ElseIf CustFrom <> "" And custTo = "" Then
                strCustCondition = " CstmrCode = '" & CustFrom & "'"
            ElseIf CustFrom = "" And custTo <> "" Then
                strCustCondition = " CstmrCode = '" & custTo & "'"
            Else
                strCustCondition = " 1=1"
            End If
            strCondition = strDateCondition1 & " and " & strCustCondition & " and " & WarrantyCondition & " and " & ItemGroupCondition & " and " & StatusCondition & " and " & InvoiceCondition
            Dim strGroupby As String = "Group  by T0.[ContractID],T0.[CstmrCode],T0.[CstmrName],T0.[StartDate],T0.[EndDate],T0.[Status],T0.U_Z_Currency,T0.[U_Z_ContAmt],T0.[U_Z_WarType],T0.[CntrcType],isnull(T0.U_Z_Invoice,'N') , T0.U_Z_GLAcc,T3.AcctCode,T0.U_Z_InvNo "
            strstring = "SELECT '' as 'Select', T0.[ContractID],T0.[CstmrCode],T0.[CstmrName],T0.[StartDate],T0.[EndDate],T0.[Status],T0.U_Z_Currency,T0.[U_Z_ContAmt],T0.[U_Z_WarType],T0.[CntrcType],isnull(T0.U_Z_Invoice,'N') 'Invoiced', T0.U_Z_GLAcc 'GL',T3.AcctCode 'AcctCode' ,T0.U_Z_InvNo 'Invoice Number',Count(*) 'Count' FROM OCTR T0  INNER JOIN CTR1 T1 ON T0.ContractID = T1.ContractID Left outer Join OACT T3 on T3.FormatCode = T0.U_Z_GLAcc where  " & strCondition & strGroupby
            oGrid = aform.Items.Item("6").Specific
            oGrid.DataTable.ExecuteQuery(strstring)
            Formatgrid(oGrid)
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
#End Region

#Region "FormatGrid"
    Private Sub Formatgrid(ByVal agrid As SAPbouiCOM.Grid)
        agrid.Columns.Item(0).TitleObject.Caption = "Generate"
        agrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        agrid.Columns.Item(0).Editable = True
        agrid.Columns.Item("ContractID").TitleObject.Caption = "Contract No"
        agrid.Columns.Item("ContractID").Editable = False
        agrid.Columns.Item("U_Z_Currency").TitleObject.Caption = "Currency"
        agrid.Columns.Item("U_Z_Currency").Editable = False
        oEditTextColumn = agrid.Columns.Item("ContractID")
        oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_ServiceContract
        agrid.Columns.Item("CstmrCode").TitleObject.Caption = "Customer Code"
        agrid.Columns.Item("CstmrCode").Editable = False
        oEditTextColumn = agrid.Columns.Item("CstmrCode")
        oEditTextColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
        agrid.Columns.Item("CstmrName").TitleObject.Caption = "Customer Name"
        agrid.Columns.Item("CstmrName").Editable = False
        agrid.Columns.Item("StartDate").TitleObject.Caption = "Start Date"
        agrid.Columns.Item("StartDate").Editable = False
        agrid.Columns.Item("EndDate").TitleObject.Caption = "End Date"
        agrid.Columns.Item("EndDate").Editable = False
        oGrid.Columns.Item("Status").TitleObject.Caption = "Contract Status"
        oGrid.Columns.Item("Status").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        ocombo = oGrid.Columns.Item("Status")
        ocombo.ValidValues.Add("", "")
        ocombo.ValidValues.Add("A", "Approved")
        ocombo.ValidValues.Add("D", "Draft")
        ocombo.ValidValues.Add("F", "Frozen")
        ocombo.ValidValues.Add("T", "Terminated")
        ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
        oGrid.Columns.Item("Status").Editable = False

        agrid.Columns.Item("U_Z_ContAmt").TitleObject.Caption = "Contract Amount"
        agrid.Columns.Item("U_Z_ContAmt").Editable = False
        agrid.Columns.Item("U_Z_WarType").TitleObject.Caption = "Warranty Type"
        agrid.Columns.Item("U_Z_WarType").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        ocombo = oGrid.Columns.Item("U_Z_WarType")
        ocombo.ValidValues.Add("", "")
        ocombo.ValidValues.Add("F", "Free")
        ocombo.ValidValues.Add("W", "Warranty")
        ocombo.ValidValues.Add("R", "Regular")
        ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
        agrid.Columns.Item("U_Z_WarType").Editable = False
        agrid.Columns.Item("CntrcType").TitleObject.Caption = "Contract Type"
        agrid.Columns.Item("CntrcType").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        ocombo = oGrid.Columns.Item("CntrcType")
        ocombo.ValidValues.Add("", "")
        ocombo.ValidValues.Add("S", "Serial Number")
        ocombo.ValidValues.Add("G", "Item Group")
        ocombo.ValidValues.Add("C", "Customer")
        ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
        agrid.Columns.Item("CntrcType").Editable = False
        agrid.Columns.Item("GL").TitleObject.Caption = "G/L Account"
        agrid.Columns.Item("GL").Editable = False
        agrid.Columns.Item("Invoiced").TitleObject.Caption = "Invoiced"
        agrid.Columns.Item("Invoiced").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        ocombo = oGrid.Columns.Item("Invoiced")
        ocombo.ValidValues.Add("", "")
        ocombo.ValidValues.Add("Y", "Yes")
        ocombo.ValidValues.Add("N", "No")
        ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
        agrid.Columns.Item("Invoiced").Editable = False
        agrid.Columns.Item("Count").Visible = False
        agrid.Columns.Item("AcctCode").TitleObject.Caption = "SAP Account Code"
        agrid.Columns.Item("AcctCode").Editable = False
        agrid.Columns.Item("Invoice Number").Editable = False
        agrid.AutoResizeColumns()
        agrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
        For intRow As Integer = 0 To agrid.DataTable.Rows.Count - 1
            agrid.RowHeaders.SetText(intRow, intRow + 1)
        Next
    End Sub
#End Region
    Private Sub SelectAll(ByVal aform As SAPbouiCOM.Form, ByVal aValue As Boolean)
        aform.Freeze(True)
        oGrid = aform.Items.Item("6").Specific
        Dim oCheckbox As SAPbouiCOM.CheckBoxColumn
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oCheckbox = oGrid.Columns.Item(0)
            oCheckbox.Check(intRow, aValue)
        Next
        aform.Freeze(False)
    End Sub
#Region "AddtoUDT"
    Private Function AddtoUDT1(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim Startdate, EndDate As Date
        Dim strString As String
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oCheckbox As SAPbouiCOM.CheckBoxColumn
        Dim oServiceCon As SAPbobsCOM.Documents

        oGrid = aform.Items.Item("6").Specific
        Try
            aform.Freeze(True)
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oCheckbox = oGrid.Columns.Item(0)
                If oCheckbox.IsChecked(intRow) Then
                    oApplication.Utilities.Message("Processing. Contract ID : " & oGrid.DataTable.GetValue("ContractID", intRow), SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    If oGrid.DataTable.GetValue("U_Z_ContAmt", intRow) > 0 Then
                        oServiceCon = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                        oServiceCon.CardCode = oGrid.DataTable.GetValue("CstmrCode", intRow)
                        oServiceCon.CardName = oGrid.DataTable.GetValue("CstmrName", intRow)
                        oServiceCon.DocDueDate = Now.Date
                        oServiceCon.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                        If oGrid.DataTable.GetValue("U_Z_Currency", intRow) <> "" Then
                            oServiceCon.DocCurrency = oGrid.DataTable.GetValue("U_Z_Currency", intRow)
                        End If
                        Dim dtstartdate, dtenddate As Date
                        dtstartdate = oGrid.DataTable.GetValue("StartDate", intRow)
                        dtenddate = oGrid.DataTable.GetValue("EndDate", intRow)

                        oServiceCon.Comments = "Created based of Contract ID : " & oGrid.DataTable.GetValue("ContractID", intRow) & " Start Date " & dtstartdate.ToString("dd.MM.yyyy") & " and end date  " & dtenddate.ToString("dd.MM.yyyy")

                        oServiceCon.Lines.AccountCode = oGrid.DataTable.GetValue("AcctCode", intRow)
                        oServiceCon.Lines.LineTotal = oGrid.DataTable.GetValue("U_Z_ContAmt", intRow)
                        oServiceCon.Lines.ItemDescription = "After Sales Service"
                        If oServiceCon.Add() <> 0 Then
                            '  oApplication.Utilities.Message("Error in Creating Invoice . Contract ID" & oGrid.DataTable.GetValue("ContractID", intRow) & " Error Message : " & ex.m
                            Dim strMessage As String = "Error while Creating Invoice . Contract ID : " & oGrid.DataTable.GetValue("ContractID", intRow)
                            strMessage = strMessage & " Error Details : " & oApplication.Company.GetLastErrorDescription
                            oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            ' aform.Freeze(False)
                            ' Return False
                        Else
                            Dim str As String
                            oApplication.Company.GetNewObjectCode(str)
                            oServiceCon.GetByKey(CInt(str))

                            Dim strMessage As String = "Invoice Created for Contract ID : " & oGrid.DataTable.GetValue("ContractID", intRow)
                            strMessage = strMessage & " Invoice Number : " & oServiceCon.DocNum.ToString
                            oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            strString = "Update OCTR set U_Z_Invoice='Y',U_Z_InvNo='" & oServiceCon.DocNum.ToString & "' where ContractID=" & oGrid.DataTable.GetValue("ContractID", intRow) & ""
                            oRec.DoQuery(strString)
                        End If
                    End If
                End If
            Next
            Databind(aform)
            aform.Freeze(False)
            Return True

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Function


    Private Function Validation_Grid(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim Startdate, EndDate As Date
        Dim strString As String
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oCheckbox As SAPbouiCOM.CheckBoxColumn
        Dim oServiceCon As SAPbobsCOM.Documents
        oServiceCon = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        oGrid = aform.Items.Item("6").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oCheckbox = oGrid.Columns.Item(0)
            If oCheckbox.IsChecked(intRow) Then
                If oGrid.DataTable.GetValue("AcctCode", intRow) = "" Then
                    oApplication.Utilities.Message("Account code not defined for this contract : Contract ID : " & oGrid.DataTable.GetValue("ContractID", intRow), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
        Next
        Return True
    End Function
#End Region




#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_InvGeneration Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And oForm.PaneLevel = 2 Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "6" And pVal.ColUID = "Select" Then
                                    oGrid = oForm.Items.Item("6").Specific
                                    If oGrid.DataTable.GetValue("Invoiced", pVal.Row) = "Y" Then
                                        '  oGrid.Columns.Item("Select").Click(row, False)
                                        BubbleEvent = False
                                        Exit Sub
                                        'Else
                                        '    oGrid.Columns.Item("Select").Click(row, True)
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "6" And pVal.ColUID = "Select" Then
                                    oGrid = oForm.Items.Item("6").Specific
                                    If oGrid.DataTable.GetValue("Invoiced", pVal.Row) = "Y" Then
                                        '  oGrid.Columns.Item("Select").Click(row, False)
                                        BubbleEvent = False
                                        Exit Sub
                                        'Else
                                        '    oGrid.Columns.Item("Select").Click(row, True)
                                    End If

                                End If

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "6" And pVal.ColUID = "Select" Then
                                    oGrid = oForm.Items.Item("6").Specific
                                    If oGrid.DataTable.GetValue("Invoiced", pVal.Row) = "Y" Then
                                        '  oGrid.Columns.Item("Select").Click(row, False)
                                        BubbleEvent = False
                                        Exit Sub
                                        'Else
                                        '    oGrid.Columns.Item("Select").Click(row, True)
                                    End If

                                End If

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "32"
                                        SelectAll(oForm, False)
                                    Case "31"
                                        SelectAll(oForm, True)
                                    Case "3"
                                        oForm.Freeze(True)
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                        If oForm.PaneLevel = 3 Then
                                            Databind(oForm)
                                        End If
                                        oForm.Freeze(False)
                                    Case "9"
                                        oForm.Freeze(True)
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                        oForm.Freeze(False)
                                    Case "12"
                                        If Validation_Grid(oForm) = False Then
                                            Exit Sub
                                        End If

                                        If oApplication.SBO_Application.MessageBox("Do you want to Create the Invoices for  selected contracts?", , "Continue", "Cancle") = 2 Then
                                            Exit Sub
                                        End If
                                        If AddtoUDT1(oForm) = True Then
                                            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            '   oForm.Close()
                                        Else
                                            Databind(oForm)
                                        End If
                                End Select

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oItm As SAPbobsCOM.Items
                                Dim sCHFL_ID, val As String
                                Dim intChoice, introw As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        oForm.Freeze(True)
                                        oForm.Update()
                                        If pVal.ItemUID = "14" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "14", val)
                                            Catch ex As Exception
                                                oForm.Freeze(False)
                                            End Try
                                        End If
                                        If pVal.ItemUID = "16" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "16", val)
                                            Catch ex As Exception
                                                oForm.Freeze(False)
                                            End Try
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    'MsgBox(ex.Message)
                                End Try
                        End Select
                End Select
            End If

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_InvGeneration
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
              
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class
