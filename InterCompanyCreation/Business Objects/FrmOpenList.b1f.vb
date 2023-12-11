Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Data
Namespace InterCompanyCreation
    <FormAttribute("OpenList", "Business Objects/FrmOpenList.b1f")>
    Friend Class FrmOpenList
        Inherits UserFormBase
        Public WithEvents objform, objformUDF, objformNew As SAPbouiCOM.Form
        Public WithEvents objText As SAPbouiCOM.EditText
        Private WithEvents objMatrix As SAPbouiCOM.Matrix
        Private WithEvents odbdsDetails As SAPbouiCOM.DBDataSource
        Public objCompany1 As New SAPbobsCOM.Company
        Public UDFFormID As String
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("11").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("MtxOpen").Specific, SAPbouiCOM.Matrix)
            Me.StaticText0 = CType(Me.GetItem("Item_8").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("Item_9").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try

                objform = objaddon.objapplication.Forms.GetForm("OpenList", 0)
                objform.Settings.Enabled = True
                bModal = True
                'objform = objaddon.objapplication.Forms.ActiveForm
                'If objaddon.objapplication.Forms.ActiveForm.Type = "139" Then
                '    objformNew = objaddon.objapplication.Forms.GetForm("139", 0)
                '    FormID = objaddon.objapplication.Forms.GetEventForm(objformNew.Type).TypeID.ToString
                '    UDFFormID = -FormID
                '    If Not objaddon.objapplication.Menus.Item("6913").Checked = True Then
                '        objaddon.objapplication.Menus.Item("6913").Activate()
                '        objformUDF = objaddon.objapplication.Forms.GetForm("-139", 0)
                '        objText = objformUDF.Items.Item("U_IComNum").Specific
                '    End If
                'ElseIf objaddon.objapplication.Forms.ActiveForm.Type = "142" Then
                '    objformNew = objaddon.objapplication.Forms.GetForm("142", 0)
                '    FormID = objaddon.objapplication.Forms.GetEventForm(objformNew.Type).TypeID.ToString
                '    UDFFormID = -FormID
                '    If Not objaddon.objapplication.Menus.Item("6913").Checked = True Then
                '        objaddon.objapplication.Menus.Item("6913").Activate()
                '        objformUDF = objaddon.objapplication.Forms.GetForm("-142", 0)
                '        objText = objformUDF.Items.Item("U_IComNum").Specific
                '    End If
                'End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix

        
        Public Function GetList(ByVal Query As String, ByVal DBName As String) As SAPbobsCOM.Recordset
            Dim objRS As SAPbobsCOM.Recordset
            objCompany1 = objaddon.objglobalmethods.ConnectToCompany(DBName)
            objRS = objCompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery(Query)

            Return objRS
        End Function

        Public Sub LoadMatrixFromGRPO_Deliveries(ByVal DocName As String)
            Try
                Dim StrQuery As String = "", DBName As String = ""
                Dim objDTable As SAPbouiCOM.DataTable
                Dim objRS1 As SAPbobsCOM.Recordset
                objform = objaddon.objapplication.Forms.GetForm("OpenList", 0)
                objform = objaddon.objapplication.Forms.ActiveForm
                objMatrix = objform.Items.Item("MtxOpen").Specific
                objRS1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If DocName = "GRPO" Then
                    objform.Title = "GRPO Open Entries"
                    StrQuery = "select  ROW_NUMBER() OVER () AS ""LineId"",T0.""DocEntry"",T0.""DocNum"",TO_VARCHAR(T0.""DocDate"",'dd/MM/yy') as ""DocDate"",TO_VARCHAR(T0.""DocDueDate"",'dd/MM/yy') as ""DocDueDate"",T0.""CardCode"",T0.""CardName"",T0.""DocTotal"",T0.""ObjType"""
                    StrQuery += vbCrLf + " from OPDN T0 left join ORDR T1 on  T0.""DocEntry""=T1.""U_IComNum"" where T0.""DocStatus""='O' and T1.""U_IComNum"" is  null or (T1.""U_IComNum"" is not null and T1.""CANCELED""='Y') Order by T0.""DocDate"" Desc"
                ElseIf DocName = "Delivery" Then
                    objform.Title = "Delivery Open Entries"
                    DBName = objaddon.objglobalmethods.getSingleValue("select Top 1 ""U_DBName"" from ""@DB_LIST"" where ""U_DBName"" is not null")
                    StrQuery = "select  ROW_NUMBER() OVER () AS ""LineId"",T0.""DocEntry"",T0.""DocNum"",TO_VARCHAR(T0.""DocDate"",'dd/MM/yy') as ""DocDate"",TO_VARCHAR(T0.""DocDueDate"",'dd/MM/yy') as ""DocDueDate"",T0.""CardCode"",T0.""CardName"",T0.""DocTotal"",T0.""ObjType"""
                    StrQuery += vbCrLf + " from """ & DBName & """.ODLN T0 left join OPOR T1 on  T0.""DocEntry""=T1.""U_IComNum"" where T0.""DocStatus""='O' and T1.""U_IComNum"" is  null or (T1.""U_IComNum"" is not null and T1.""CANCELED""='Y') Order by T0.""DocDate"" Desc"
                End If

                If objform.DataSources.DataTables.Count.Equals(0) Then
                    objform.DataSources.DataTables.Add("dtLoad")
                Else
                    objform.DataSources.DataTables.Item("dtLoad").Clear()
                End If

                objDTable = objform.DataSources.DataTables.Item("dtLoad")
                'objform.DataSources.DataTables.Item("dtLoad").ExecuteQuery(StrQuery)
                If DocName = "Delivery" Then
                    objRS1.DoQuery(StrQuery) '= GetList(StrQuery, DBName)
                    objDTable.Columns.Add("LineId", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254)
                    objDTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254)
                    objDTable.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254)
                    objDTable.Columns.Add("DocDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254)
                    objDTable.Columns.Add("DocDueDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254)
                    objDTable.Columns.Add("CardCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254)
                    objDTable.Columns.Add("CardName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254)
                    objDTable.Columns.Add("DocTotal", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254)
                    objDTable.Columns.Add("ObjType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254)
                Else
                    objDTable.ExecuteQuery(StrQuery)
                    objRS1.DoQuery(StrQuery)
                End If
                Dim objColumn As SAPbouiCOM.Column
                objColumn = objMatrix.Columns.Item("0")
                objColumn.DataBind.Bind("dtLoad", "LineId")
                objColumn = objMatrix.Columns.Item("1")
                objColumn.DataBind.Bind("dtLoad", "DocEntry")
                objColumn = objMatrix.Columns.Item("2")
                objColumn.DataBind.Bind("dtLoad", "DocNum")
                objColumn = objMatrix.Columns.Item("3")
                objColumn.DataBind.Bind("dtLoad", "DocDate")
                objColumn = objMatrix.Columns.Item("4")
                objColumn.DataBind.Bind("dtLoad", "DocDueDate")
                objColumn = objMatrix.Columns.Item("5")
                objColumn.DataBind.Bind("dtLoad", "CardCode")
                objColumn = objMatrix.Columns.Item("6")
                objColumn.DataBind.Bind("dtLoad", "CardName")
                objColumn = objMatrix.Columns.Item("7")
                objColumn.DataBind.Bind("dtLoad", "DocTotal")
                objColumn = objMatrix.Columns.Item("8")
                objColumn.DataBind.Bind("dtLoad", "ObjType")
                objDTable.Rows.Clear()

                'objMatrix.Columns.Item("0").TitleObject.Sortable = True
                'objMatrix.Columns.Item("1").TitleObject.Sortable = True
                'objMatrix.Columns.Item("2").TitleObject.Sortable = True
                'objMatrix.Columns.Item("3").TitleObject.Sortable = True
                'objMatrix.Columns.Item("4").TitleObject.Sortable = True
                'objMatrix.Columns.Item("5").TitleObject.Sortable = True
                'objMatrix.Columns.Item("6").TitleObject.Sortable = True
                'objMatrix.Columns.Item("7").TitleObject.Sortable = True

                'objDTable.Rows.Add(objRS1.RecordCount)
                For i As Integer = 0 To Matrix0.Columns.Count - 1
                    objMatrix.Columns.Item(i).TitleObject.Sortable = True
                Next
                While Not objRS1.EoF
                    objDTable.Rows.Add()
                    objDTable.SetValue(0, objDTable.Rows.Count - 1, objRS1.Fields.Item("LineId").Value)
                    objDTable.SetValue(1, objDTable.Rows.Count - 1, objRS1.Fields.Item("DocEntry").Value)
                    objDTable.SetValue(2, objDTable.Rows.Count - 1, objRS1.Fields.Item("DocNum").Value)
                    objDTable.SetValue(3, objDTable.Rows.Count - 1, objRS1.Fields.Item("DocDate").Value)
                    objDTable.SetValue(4, objDTable.Rows.Count - 1, objRS1.Fields.Item("DocDueDate").Value)
                    objDTable.SetValue(5, objDTable.Rows.Count - 1, objRS1.Fields.Item("CardCode").Value)
                    objDTable.SetValue(6, objDTable.Rows.Count - 1, objRS1.Fields.Item("CardName").Value)
                    objDTable.SetValue(7, objDTable.Rows.Count - 1, objRS1.Fields.Item("DocTotal").Value)
                    objDTable.SetValue(8, objDTable.Rows.Count - 1, objRS1.Fields.Item("ObjType").Value)
                    objRS1.MoveNext()
                End While

                objMatrix.Clear()
                objMatrix.LoadFromDataSource()
                'objMatrix.LoadFromDataSourceEx()
                'objaddon.objapplication.Menus.Item("1300").Activate()
                Matrix0.SelectRow(1, True, False)
                objform.Settings.Enabled = True
                objDTable = Nothing
                'objMatrix.Columns.Item("8").Visible = False
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                If pVal.ActionSuccess = True Then
                    Dim docent, docnum, objtype As String
                    Dim GetValues As List(Of Transaction) = GetSelectedEntry()
                    objform.Close()
                    If GetValues.Count > 0 Then
                        Dim DocEntryList = (From gv In GetValues Select New String(gv.DocEntry)).ToList()
                        docent = String.Join(",", DocEntryList)
                        Dim DocNumList = (From gv In GetValues Select New String(gv.DocNum)).ToList()
                        docnum = String.Join(",", DocNumList)
                        Dim TypeList = (From gv In GetValues Select New String(gv.ObjType)).ToList()
                        objtype = TypeList(0) 'String.Join(",", TypeList)
                        If objtype = "20" Then
                            LoadFromGRPO(docent, docnum)
                        Else
                            LoadFromDeliveries(docent, docnum)
                        End If
                    End If
                End If
            Catch ex As Exception
            End Try

        End Sub
        Public Class Transaction
            Public DocEntry As String
            Public DocNum As String
            Public ObjType As String
        End Class
        Private Function GetSelectedEntry()
            Try
                Dim GetTranValues As New List(Of Transaction)
                For i As Integer = 1 To Matrix0.VisualRowCount
                    If Matrix0.IsRowSelected(i) Then
                        Dim GetTran As New Transaction
                        GetTran.DocEntry = Matrix0.Columns.Item("1").Cells.Item(i).Specific.String
                        GetTran.DocNum = Matrix0.Columns.Item("2").Cells.Item(i).Specific.String
                        GetTran.ObjType = Matrix0.Columns.Item("8").Cells.Item(i).Specific.String
                        GetTranValues.Add(GetTran)
                        ' DocEntry = Matrix0.Columns.Item("1").Cells.Item(i).Specific.String
                        'ObjType = Matrix0.Columns.Item("8").Cells.Item(i).Specific.String
                        'DocNum = Matrix0.Columns.Item("2").Cells.Item(i).Specific.String
                    End If
                Next
                Return GetTranValues
            Catch ex As Exception
                Return New List(Of Transaction)
            End Try
        End Function

        Private Sub LoadFromGRPO(ByVal DocEntry As String, ByVal DocNum As String)
            Dim StrQry As String
            Dim FormID As String = ""
            Dim objRS1, objRS As SAPbobsCOM.Recordset
            Dim objSOform As SAPbouiCOM.Form
            Dim objSOMatrix As SAPbouiCOM.Matrix
            Dim Row As Integer = 0
            Try
                If DocEntry = "" Then
                    Exit Sub
                End If
                GRPODocEntry = DocEntry
                objaddon.objapplication.SetStatusBarMessage("GRPO Loading to Sales Order Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                objSOform = objaddon.objapplication.Forms.GetForm("139", 0)
                objSOMatrix = objSOform.Items.Item("38").Specific
                objRS1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objaddon.objapplication.SetStatusBarMessage("GRPO Loading to Sales Order Please wait... DocumentNumber-> " & DocNum, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                StrQry = "Select T0.""DocEntry"",T1.""LineNum"",T1.""ItemCode"",T1.""Dscription"",T1.""Quantity"", ((T1.""Price""*T0.""DocRate"")+((T1.""Price""*T0.""DocRate"")*((select Top 1 ""U_MarkUp"" from ""@MI_MARKUP"" order by ""Code"" Desc)/100))) as ""UnitPrice"","
                StrQry += vbCrLf + "  T1.""TaxCode"",T1.""WhsCode"", (select Top 1 ""U_MarkUp"" from ""@MI_MARKUP"" order by ""Code"" Desc) as ""MarkUp"""
                StrQry += vbCrLf + " from OPDN T0 join PDN1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""DocEntry"" in(" & DocEntry & "); "
                objRS1.DoQuery(StrQry)
                objSOMatrix.Clear()
                objSOMatrix.AddRow()
                'objform = objaddon.objapplication.Forms.GetForm(UDFFormID, 0)
                'objform.Items.Item("U_IComNum").Specific.String = DocEntry
                If objRS1.RecordCount > 0 Then
                    For i As Integer = 0 To objRS1.RecordCount - 1
                        Row += 1
                        objSOMatrix.Columns.Item("1").Cells.Item(Row).Specific.String = objRS1.Fields.Item("ItemCode").Value.ToString
                        objSOMatrix.Columns.Item("11").Cells.Item(Row).Specific.String = objRS1.Fields.Item("Quantity").Value.ToString
                        If objSOMatrix.Columns.Item("U_MarkUp").Visible = True Then
                            objSOMatrix.Columns.Item("U_MarkUp").Cells.Item(Row).Specific.String = objRS1.Fields.Item("MarkUp").Value.ToString
                        End If
                        objSOMatrix.Columns.Item("U_GRNum").Cells.Item(Row).Specific.String = objRS1.Fields.Item("DocEntry").Value.ToString
                        objSOMatrix.Columns.Item("U_GRLine").Cells.Item(Row).Specific.String = objRS1.Fields.Item("LineNum").Value.ToString
                        objSOMatrix.Columns.Item("14").Cells.Item(Row).Specific.String = objRS1.Fields.Item("UnitPrice").Value.ToString
                        objSOMatrix.Columns.Item("160").Cells.Item(Row).Specific.String = objRS1.Fields.Item("TaxCode").Value.ToString
                        objSOMatrix.Columns.Item("24").Cells.Item(Row).Specific.String = objRS1.Fields.Item("WhsCode").Value.ToString
                        objRS1.MoveNext()
                    Next
                End If

                objaddon.objapplication.StatusBar.SetText("GRPO Loaded to Sales Order Successfully!!! ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objRS1 = Nothing
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub LoadFromDeliveries(ByVal DocEntry As String, ByVal DocNum As String)
            Dim StrQry, DBName As String
            Dim objRS1 As SAPbobsCOM.Recordset
            Dim objPOform As SAPbouiCOM.Form
            Dim objPOMatrix As SAPbouiCOM.Matrix
            Dim Row As Integer = 0
            Try
                If DocEntry = "" Then
                    Exit Sub
                End If
                DeliveryDocEntry = DocEntry
                objPOform = objaddon.objapplication.Forms.GetForm("142", 0)
                objPOMatrix = objPOform.Items.Item("38").Specific
                objRS1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                DBName = objaddon.objglobalmethods.getSingleValue("select ""U_DBName"" from ""@DB_LIST""")
                objaddon.objapplication.SetStatusBarMessage("Delivery Loading to Purchase Order Please Wait... DocumentNumber->" & DocNum, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                StrQry = " Select T1.""ItemCode"",T1.""Dscription"",T1.""Quantity"",(T0.""DocRate""*T1.""Price"") as ""Price"", T1.""TaxCode"",(select ""U_ToWhse"" from ""@MI_WHSE"" where ""U_FrmWhse""= T1.""WhsCode"") as ""WhsCode"" "
                StrQry += vbCrLf + "  from """ & DBName & """.ODLN T0 join """ & DBName & """.DLN1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""DocEntry""in (" & DocEntry & ") "
                objRS1.DoQuery(StrQry)
                objPOMatrix.Clear()
                objPOMatrix.AddRow()
                'If Not objaddon.objapplication.Menus.Item("6913").Checked = True Then
                '    objaddon.objapplication.Menus.Item("6913").Activate()
                'End If
                'objPOform = objaddon.objapplication.Forms.GetForm("-142", 1)
                'objPOform.Items.Item("U_IComNum").Specific.String = DocEntry
                If objRS1.RecordCount > 0 Then
                    For i As Integer = 0 To objRS1.RecordCount - 1
                        Row += 1
                        objPOMatrix.Columns.Item("1").Cells.Item(Row).Specific.String = objRS1.Fields.Item("ItemCode").Value.ToString
                        objPOMatrix.Columns.Item("11").Cells.Item(Row).Specific.String = objRS1.Fields.Item("Quantity").Value.ToString
                        objPOMatrix.Columns.Item("14").Cells.Item(Row).Specific.String = objRS1.Fields.Item("Price").Value.ToString
                        objPOMatrix.Columns.Item("160").Cells.Item(Row).Specific.String = objRS1.Fields.Item("TaxCode").Value.ToString
                        objPOMatrix.Columns.Item("24").Cells.Item(Row).Specific.String = objRS1.Fields.Item("WhsCode").Value.ToString
                        objRS1.MoveNext()
                    Next
                End If
                objaddon.objapplication.StatusBar.SetText("Delivery Loaded to Purchase Order Successfully!!! ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objRS1 = Nothing
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub EditText1_KeyDownAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText1.KeyDownAfter
            Try
                Dim FindString As String
                FindString = EditText1.Value
                Dim Flag As Boolean = False
                Dim ColumnNum As Integer
                For i As Integer = 1 To Matrix0.Columns.Count - 1
                    If Matrix0.Columns.Item(i).TitleObject.Sortable = True Then
                        Flag = True
                        ColumnNum = i
                        Exit For
                    Else
                        Flag = False
                    End If
                Next
                If Flag Then
                    For j As Integer = 1 To Matrix0.RowCount
                        If FindString Like Matrix0.Columns.Item(ColumnNum).Cells.Item(j).Specific.String Then
                            Matrix0.SelectRow(j, True, False)
                        End If
                    Next
                Else
                    For j As Integer = 1 To Matrix0.RowCount
                        If FindString Like Matrix0.Columns.Item(ColumnNum).Cells.Item(j).Specific.String Then
                            Matrix0.SelectRow(j, True, False)
                        End If
                    Next
                End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_DoubleClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.DoubleClickAfter
            Try
                If pVal.Row = 0 Then
                    Exit Sub
                End If
                If pVal.ActionSuccess = True And pVal.Row <> 0 Then
                    Dim docent, docnum, objtype As String
                    Dim GetValues As List(Of Transaction) = GetSelectedEntry()
                    objform.Close()
                    If GetValues.Count > 0 Then
                        Dim DocEntryList = (From gv In GetValues Select New String(gv.DocEntry)).ToList()
                        docent = String.Join(",", DocEntryList)
                        Dim DocNumList = (From gv In GetValues Select New String(gv.DocNum)).ToList()
                        docnum = String.Join(",", DocNumList)
                        Dim TypeList = (From gv In GetValues Select New String(gv.ObjType)).ToList()
                        objtype = TypeList(0) 'String.Join(",", TypeList)
                        If objtype = "20" Then
                            LoadFromGRPO(docent, docnum)
                        Else
                            LoadFromDeliveries(docent, docnum)
                        End If
                    End If

                End If
              
                'If pVal.Row <> 0 Then
                '    Dim DocEntry() As String, DocEntry1() As String = {""}, DocNum() As String = {""}
                '    Dim ObjType As String = ""
                '    DocEntry = GetSelectedEntry()
                '    DocEntry1 = DocEntry
                '    DocNum = DocEntry
                '    ObjType = DocEntry(1)
                '    'DocEntry1 = DocEntry(0)
                '    'ObjType = DocEntry(1)
                '    'DocNum = DocEntry(2)
                '    If ObjType = "20" Then
                '        'LoadFromGRPO(DocEntry1, DocNum)
                '    Else
                '        LoadFromDeliveries(DocEntry1, DocNum)
                '    End If

                'End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ClickAfter
            Try
                If pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_SHIFT Or pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_CTRL Then
                    If pVal.Row <> 0 Then
                        Matrix0.SelectRow(pVal.Row, True, True)
                    Else
                        Matrix0.Columns.Item(pVal.ColUID).TitleObject.Sortable = True
                    End If
                Else
                    If pVal.Row <> 0 Then
                        Matrix0.SelectRow(pVal.Row, True, False)
                    Else
                        Matrix0.Columns.Item(pVal.ColUID).TitleObject.Sortable = True
                    End If
                End If

            Catch ex As Exception
            End Try
        End Sub

        'Private Sub Matrix0_KeyDownAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.KeyDownAfter
        '    Try
        '        If pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_SHIFT Or pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_CTRL Then
        '            If pVal.CharPressed = 38 Or pVal.Row <> 0 Then
        '                Matrix0.SelectRow(pVal.Row - 1, True, True)
        '            End If
        '        Else
        '            If pVal.CharPressed = 40 Or pVal.Row <> 0 Then
        '                Matrix0.SelectRow(pVal.Row + 1, True, False)
        '            End If
        '        End If
        '    Catch ex As Exception
        '    End Try
        'End Sub

        'Private Sub Matrix0_KeyDownBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Matrix0.KeyDownBefore
        '    Try
        '        If pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_SHIFT Or pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_CTRL Then
        '            If pVal.CharPressed = 38 Or pVal.Row <> 0 Then
        '                Matrix0.SelectRow(pVal.Row - 1, True, True)
        '            End If
        '        Else
        '            If pVal.CharPressed = 40 Or pVal.Row <> 0 Then
        '                Matrix0.SelectRow(pVal.Row + 1, True, False)
        '            End If
        '        End If

        '    Catch ex As Exception
        '    End Try

        'End Sub

    End Class
End Namespace
