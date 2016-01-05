Public Class clsContRenewal
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
        oForm = oApplication.Utilities.LoadForm(xml_ContRenewal, frm_ContRenewal)
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
        ' oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly


        'oForm.DataSources.UserDataSources.Add("Issued", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        'oApplication.Utilities.setUserDatabind(oForm, "24", "Issued")
        'ocombo = oForm.Items.Item("22").Specific
        'For intRow As Integer = ocombo.ValidValues.Count - 1 To 0 Step -1
        '    ocombo.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
        'Next
        'oCombobox.ValidValues.Add("", "")
        'oCombobox.ValidValues.Add("Y", "Yes")
        'oCombobox.ValidValues.Add("N", "No")

        'oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        'ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
        'ocombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

        oForm.PaneLevel = 1
        oForm.Freeze(False)
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

#Region "Databind"
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim strstring, EndFrom, EndTo, CustFrom, custTo, Warranty, Itemgroup, ConStatus, WarrantyCondition, ItemGroupCondition, StatusCondition, strCondition As String
            Dim Endfrdt, EndTodt As Date
            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oCombobox = aform.Items.Item("18").Specific
            oCombobox1 = aform.Items.Item("20").Specific
            oCombobox2 = aform.Items.Item("22").Specific

            EndFrom = oApplication.Utilities.getEdittextvalue(aform, "10")
            EndTo = oApplication.Utilities.getEdittextvalue(aform, "1000002")
            CustFrom = oApplication.Utilities.getEdittextvalue(aform, "14")
            custTo = oApplication.Utilities.getEdittextvalue(aform, "16")

            Warranty = oCombobox.Selected.Value
            Itemgroup = oCombobox1.Selected.Value
            ConStatus = oCombobox2.Selected.Value

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

            strCondition = strDateCondition1 & " and " & strCustCondition & " and " & WarrantyCondition & " and " & ItemGroupCondition & " and " & StatusCondition

            'strstring = "SELECT '' as Renewal, T0.[ContractID],T0.[CstmrCode],T0.[CstmrName],T0.[StartDate],T0.[EndDate],T0.[U_Z_ContAmt],T0.[U_Z_WarType],T0.[CntrcType],T1.[ItemCode],T1.[ItemName],T1.[ItemGroup],T1.[ManufSN],T1.[InternalSN] FROM OCTR T0  INNER JOIN CTR1 T1 ON T0.ContractID = T1.ContractID where isnull(U_Z_Renewal,'N') <>'Y' and " & strCondition
            Dim strGroupby As String = "group by T0.[ContractID],T0.[CstmrCode],T0.[CstmrName],T0.[StartDate],T0.[EndDate],T0.[Status],T0.[U_Z_ContAmt],T0.[U_Z_WarType],T0.[CntrcType],T1.[ItemCode],T1.[ItemName],T1.[ItemGroup]"
            strstring = "SELECT '' as Renewal, T0.[ContractID],T0.[CstmrCode],T0.[CstmrName],T0.[StartDate],T0.[EndDate],T0.[Status],T0.[U_Z_ContAmt],T0.[U_Z_WarType],T0.[CntrcType],T1.[ItemCode],T1.[ItemName],T1.[ItemGroup],Count(*) 'Count' FROM OCTR T0  INNER JOIN CTR1 T1 ON T0.ContractID = T1.ContractID where isnull(U_Z_Renewal,'N') <>'Y' and " & strCondition & strGroupby
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
        agrid.Columns.Item(0).TitleObject.Caption = "Renew"
        agrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        agrid.Columns.Item(0).Editable = True
        agrid.Columns.Item("ContractID").TitleObject.Caption = "Contract No"
        agrid.Columns.Item("ContractID").Editable = False
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
        oGrid.Columns.Item("U_Z_WarType").TitleObject.Caption = "Warranty Type"
        oGrid.Columns.Item("U_Z_WarType").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        ocombo = oGrid.Columns.Item("U_Z_WarType")
        ocombo.ValidValues.Add("", "")
        ocombo.ValidValues.Add("F", "Free")
        ocombo.ValidValues.Add("W", "Warranty")
        ocombo.ValidValues.Add("R", "Regular")
        ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
        oGrid.Columns.Item("U_Z_WarType").Editable = False
        agrid.Columns.Item("CntrcType").TitleObject.Caption = "Contract Type"
        oGrid.Columns.Item("CntrcType").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        ocombo = oGrid.Columns.Item("CntrcType")
        ocombo.ValidValues.Add("", "")
        ocombo.ValidValues.Add("S", "Serial Number")
        ocombo.ValidValues.Add("G", "Item Group")
        ocombo.ValidValues.Add("C", "Customer")
        ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
        agrid.Columns.Item("CntrcType").Editable = False
        agrid.Columns.Item("ItemCode").TitleObject.Caption = "Item Code"
        agrid.Columns.Item("ItemCode").Editable = False
        oEditTextColumn = agrid.Columns.Item("ItemCode")
        oEditTextColumn.LinkedObjectType = "4"
        agrid.Columns.Item("ItemName").TitleObject.Caption = "Item Name"
        agrid.Columns.Item("ItemName").Editable = False
        agrid.Columns.Item("ItemGroup").TitleObject.Caption = "Item Group"
        agrid.Columns.Item("ItemGroup").Editable = False
        'agrid.Columns.Item("ManufSN").TitleObject.Caption = "Mfr Serial No."
        'agrid.Columns.Item("ManufSN").Editable = False
        'agrid.Columns.Item("InternalSN").TitleObject.Caption = "Serial Number"
        'agrid.Columns.Item("InternalSN").Editable = False
        agrid.Columns.Item("Count").Visible = False
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
        Dim oServiceCon As SAPbobsCOM.ServiceContracts
        oServiceCon = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceContracts)
        oGrid = aform.Items.Item("6").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oCheckbox = oGrid.Columns.Item(0)
            If oCheckbox.IsChecked(intRow) Then
                oServiceCon.CustomerCode = oGrid.DataTable.GetValue("CstmrCode", intRow)
                oServiceCon.CustomerName = oGrid.DataTable.GetValue("CstmrName", intRow)
                Startdate = oGrid.DataTable.GetValue("EndDate", intRow)
                oServiceCon.StartDate = Startdate.AddDays(1)
                EndDate = Startdate.AddDays(1)
                oServiceCon.EndDate = EndDate.AddYears(1).AddDays(-1)
                Dim otemp1 As SAPbobsCOM.Recordset
                otemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strString = "SELECT ItemCode, ItemName, ItemGroup, InternalSN, ManufSN FROM CTR1 WHERE  ContractID='" & oGrid.DataTable.GetValue("ContractID", intRow) & "'"
                otemp1.DoQuery(strString)
                For introw1 As Integer = 0 To otemp1.RecordCount - 1
                    oServiceCon.Lines.Add()
                    oServiceCon.Lines.SetCurrentLine(introw1)
                    oServiceCon.Lines.ItemCode = otemp1.Fields.Item(0).Value ' oGrid.DataTable.GetValue("ItemCode", introw1)
                    oServiceCon.Lines.ItemName = otemp1.Fields.Item(1).Value 'oGrid.DataTable.GetValue("ItemName", introw1)
                    oServiceCon.Lines.ItemGroup = otemp1.Fields.Item(2).Value ' oGrid.DataTable.GetValue("ItemGroup", introw1)
                    oServiceCon.Lines.InternalSerialNum = otemp1.Fields.Item(3).Value ' oGrid.DataTable.GetValue("InternalSN", introw1)
                    oServiceCon.Lines.ManufacturerSerialNum = otemp1.Fields.Item(4).Value ' oGrid.DataTable.GetValue("ManufSN", introw1)
                    otemp1.MoveNext()
                Next
                oServiceCon.UserFields.Fields.Item("U_Z_BaseCoNo").Value = oGrid.DataTable.GetValue("ContractID", intRow)
                If oServiceCon.Add() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else
                    strString = "Update OCTR set U_Z_Renewal='Y' where ContractID=" & oGrid.DataTable.GetValue("ContractID", intRow) & ""
                    oRec.DoQuery(strString)
                End If
            End If
        Next
        Databind(aform)
        Return True
    End Function


    Private Function CreateContracts(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim Startdate, EndDate As Date
        Dim strString As String
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oCheckbox As SAPbouiCOM.CheckBoxColumn
        Dim oServiceCon, oMainContract As SAPbobsCOM.ServiceContracts
        oServiceCon = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceContracts)
        oMainContract = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceContracts)
        oGrid = aform.Items.Item("6").Specific
        Try
            aform.Freeze(True)
         
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oCheckbox = oGrid.Columns.Item(0)
                If oCheckbox.IsChecked(intRow) Then
                    oServiceCon = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceContracts)
                    oMainContract = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceContracts)
                    If oApplication.Company.InTransaction() Then
                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    End If
                    oApplication.Company.StartTransaction()
                    If oMainContract.GetByKey(oGrid.DataTable.GetValue("ContractID", intRow)) Then
                        If oMainContract.UserFields.Fields.Item("U_Z_Renewal").Value <> "Y" Then
                            oServiceCon.CustomerCode = oMainContract.CustomerCode
                            Startdate = oGrid.DataTable.GetValue("EndDate", intRow)
                            oServiceCon.StartDate = Startdate.AddDays(1)
                            EndDate = Startdate.AddDays(1)
                            oServiceCon.EndDate = EndDate.AddYears(1).AddDays(-1)
                            oServiceCon.ContractTemplate = oMainContract.ContractTemplate
                            oServiceCon.ContractType = oMainContract.ContractType
                            oServiceCon.FridayEnabled = oMainContract.FridayEnabled
                            oServiceCon.Description = oMainContract.Description
                            oServiceCon.Remarks = oMainContract.Remarks

                            oServiceCon.FridayEnd = oMainContract.FridayEnd
                            oServiceCon.FridayStart = oMainContract.FridayStart
                            oServiceCon.IncludeHolidays = oMainContract.IncludeHolidays
                            oServiceCon.IncludeLabor = oMainContract.IncludeLabor
                            oServiceCon.IncludeParts = oMainContract.IncludeParts
                            oServiceCon.IncludeTravel = oMainContract.IncludeTravel
                            oServiceCon.MondayEnabled = oMainContract.MondayEnabled
                            oServiceCon.MondayEnd = oMainContract.MondayEnd
                            oServiceCon.MondayStart = oMainContract.MondayStart
                            oServiceCon.Owner = oMainContract.Owner
                            oServiceCon.Remarks = oMainContract.Remarks
                            oServiceCon.ReminderTime = oMainContract.ReminderTime
                            oServiceCon.RemindUnit = oMainContract.RemindUnit
                            oServiceCon.Renewal = oMainContract.Renewal
                            oServiceCon.ResolutionTime = oMainContract.ResolutionTime
                            oServiceCon.ResolutionUnit = oMainContract.ResolutionUnit
                            oServiceCon.ResponseUnit = oMainContract.ResponseUnit
                            oServiceCon.ResponseTime = oMainContract.ResponseTime
                            oServiceCon.SaturdayEnabled = oMainContract.SaturdayEnabled
                            oServiceCon.SaturdayEnd = oMainContract.SaturdayEnd
                            oServiceCon.SaturdayStart = oMainContract.SaturdayStart
                            oServiceCon.ServiceType = oMainContract.ServiceType
                            oServiceCon.Status = SAPbobsCOM.BoSvcContractStatus.scs_Approved
                            For intLoop As Integer = 0 To oMainContract.UserFields.Fields.Count - 1
                                Try
                                    oServiceCon.UserFields.Fields.Item(intLoop).Value = oMainContract.UserFields.Fields.Item(intLoop).Value
                                Catch ex As Exception
                                End Try
                            Next
                            If oMainContract.UserFields.Fields.Item("U_Z_WarType").Value = "" Then
                                oServiceCon.UserFields.Fields.Item("U_Z_WarType").Value = "F"
                            End If
                            oServiceCon.UserFields.Fields.Item("U_Z_Invoice").Value = "N"
                            oServiceCon.UserFields.Fields.Item("U_Z_InvNo").Value = ""
                            oServiceCon.UserFields.Fields.Item("U_Z_Renewal").Value = "N"
                            Try
                                oServiceCon.UserFields.Fields.Item("U_Contract_S").Value = "N"
                            Catch ex As Exception

                            End Try
                            Try
                                oServiceCon.UserFields.Fields.Item("U_Z_ContAmt").Value = oMainContract.UserFields.Fields.Item("U_Z_ContAmt").Value
                            Catch ex As Exception

                            End Try
                            If oServiceCon.ContractType = SAPbobsCOM.BoContractTypes.ct_ItemGroup Then
                                For intLoop1 As Integer = 0 To oMainContract.Lines.Count - 1
                                    oMainContract.Lines.SetCurrentLine(intLoop1)
                                    If intLoop1 > 0 Then
                                        oServiceCon.Lines.Add()
                                    End If
                                    oServiceCon.Lines.SetCurrentLine(intLoop1)
                                    oServiceCon.Lines.ItemGroup = oMainContract.Lines.ItemGroup
                                Next
                            End If
                            If oServiceCon.ContractType = SAPbobsCOM.BoContractTypes.ct_SerialNumber Then
                                For intLoop11 As Integer = 0 To oMainContract.Lines.Count - 1
                                    oMainContract.Lines.SetCurrentLine(intLoop11)
                                    If intLoop11 > 0 Then
                                        oServiceCon.Lines.Add()
                                    End If
                                    oServiceCon.Lines.SetCurrentLine(intLoop11)
                                    oServiceCon.Lines.ItemCode = oMainContract.Lines.ItemCode
                                    oServiceCon.Lines.ManufacturerSerialNum = oMainContract.Lines.ManufacturerSerialNum
                                    oServiceCon.Lines.InternalSerialNum = oMainContract.Lines.InternalSerialNum
                                    oServiceCon.Lines.ItemGroup = oMainContract.Lines.ItemGroup
                                    For intUDF As Integer = 0 To oMainContract.Lines.UserFields.Fields.Count - 1
                                        Try
                                            oServiceCon.Lines.UserFields.Fields.Item(intUDF).Value = oMainContract.Lines.UserFields.Fields.Item(intUDF).Value
                                        Catch ex As Exception

                                        End Try
                                        oServiceCon.Lines.UserFields.Fields.Item(intUDF).Value = oMainContract.Lines.UserFields.Fields.Item(intUDF).Value
                                    Next
                                Next
                            End If
                            oServiceCon.UserFields.Fields.Item("U_Z_BaseCoNo").Value = oMainContract.ContractID
                            If oServiceCon.Add() <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                If oApplication.Company.InTransaction() Then
                                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                End If
                                '  oApplication.Company.StartTransaction()
                                aform.Freeze(False)
                                Return False
                            Else
                                Dim strDocNum As String
                                oApplication.Company.GetNewObjectCode(strDocNum)
                                If CreateServiceCall(CInt(strDocNum)) = True Then
                                    oMainContract.Status = SAPbobsCOM.BoSvcContractStatus.scs_Terminated
                                    oMainContract.TerminationDate = oMainContract.EndDate
                                    oMainContract.UserFields.Fields.Item("U_Z_Renewal").Value = "Y"
                                    If oMainContract.Update <> 0 Then
                                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        If oApplication.Company.InTransaction() Then
                                            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                        End If
                                    Else
                                        If oApplication.Company.InTransaction() Then
                                            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                        End If
                                    End If
                                Else
                                    If oApplication.Company.InTransaction() Then
                                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                    End If
                                    aform.Freeze(False)
                                    Return False
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            Databind(aform)
            Return True
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
            Return False
        End Try
    End Function

    Private Function CreateServiceCall(ByVal aContID As Integer) As Boolean
        Dim oServiceCall As SAPbobsCOM.ServiceCalls
        Dim oContrcat As SAPbobsCOM.ServiceContracts
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim intOrion As Integer = -1

        oTest.DoQuery("Select * from OSCO where (Name like 'Contract' or descriptio like 'Contract')")
        If oTest.RecordCount > 0 Then
            intOrion = oTest.Fields.Item("OriginID").Value
        End If
        oContrcat = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceContracts)
        If oContrcat.GetByKey(aContID) Then
            If oContrcat.ContractType = SAPbobsCOM.BoContractTypes.ct_SerialNumber Then
                For intLoop11 As Integer = 0 To oContrcat.Lines.Count - 1
                    oServiceCall = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
                    oContrcat.Lines.SetCurrentLine(intLoop11)
                    oServiceCall.CustomerCode = oContrcat.CustomerCode
                    oServiceCall.ItemCode = oContrcat.Lines.ItemCode
                    oServiceCall.InternalSerialNum = oContrcat.Lines.InternalSerialNum
                    oServiceCall.ManufacturerSerialNum = oContrcat.Lines.ManufacturerSerialNum
                    oServiceCall.StartDate = oContrcat.StartDate
                    oServiceCall.EndDuedate = oContrcat.EndDate
                    If oContrcat.Description = "" Then
                        oServiceCall.Subject = "Elevator maintenance" ' oContrcat.Lines.ItemName
                    Else
                        oServiceCall.Subject = "Elevator maintenance" 'oContrcat.Description
                    End If
                    oServiceCall.ContractID = oContrcat.ContractID
                    oTest.DoQuery("Select * from OSCO where (Name like 'Contract' or descriptio like 'Contract')")
                    If oTest.RecordCount > 0 Then
                        oServiceCall.Origin = oTest.Fields.Item("OriginID").Value
                    Else
                        oServiceCall.Origin = 1
                    End If
                    oTest.DoQuery("Select * from OSCP where (Name like 'Monthly maintenance' or descriptio like 'Monthly maintenance')")
                    If oTest.RecordCount > 0 Then
                        oServiceCall.ProblemType = oTest.Fields.Item("prblmTypID").Value
                    Else
                        ' oServiceCall.Origin = 1
                        oServiceCall.ProblemType = 1
                    End If

                    If oServiceCall.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False

                    End If


                Next
            End If
        End If
        Return True

    End Function
#End Region

    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try

            Dim EndFrom, EndTo, custFrom, custTo As String
            oCombobox = oForm.Items.Item("18").Specific
            oCombobox1 = oForm.Items.Item("20").Specific
            oCombobox2 = oForm.Items.Item("22").Specific
         
            EndFrom = oApplication.Utilities.getEdittextvalue(aForm, "10")
            EndTo = oApplication.Utilities.getEdittextvalue(aForm, "1000002")
            custFrom = oApplication.Utilities.getEdittextvalue(aForm, "14")
            custTo = oApplication.Utilities.getEdittextvalue(aForm, "16")
            'If EndFrom = "" And EndTo = "" And custFrom = "" And custTo = "" And oCombobox.Selected.Value = "" And oCombobox1.Selected.Value = "" And oCombobox2.Selected.Value = "" Then
            '    ' oApplication.Utilities.Message("Select any one...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    ' Return False
            'End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function


#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_ContRenewal Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "3" And oForm.PaneLevel = 2 Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                '  oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
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
                                        '  If AddtoUDT1(oForm) = True Then
                                        If oApplication.SBO_Application.MessageBox("Do you want to renew the selected contracts?", , "Continue", "Cancle") = 2 Then
                                            Exit Sub
                                        End If
                                        Try
                                            'If oApplication.Company.InTransaction() Then
                                            '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                            'End If
                                            'oApplication.Company.StartTransaction()
                                            If CreateContracts(oForm) = True Then
                                                oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                oForm.Close()
                                                'If oApplication.Company.InTransaction() Then
                                                '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                'End If
                                            Else
                                                'If oApplication.Company.InTransaction() Then
                                                '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                'End If
                                                Databind(oForm)
                                            End If
                                           
                                        Catch ex As Exception
                                            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                            'If oApplication.Company.InTransaction() Then
                                            '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                            'End If
                                        End Try
                                      
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
                Case mnu_ContRenewal
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
                'Select Case pVal.MenuUID
                '    Case mnu_LeaveMaster
                '        oMenuobject = New clsEarning
                '        oMenuobject.MenuEvent(pVal, BubbleEvent)
                'End Select
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class
