Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace InterCompanyCreation
    <FormAttribute("139", "Business Objects/SalesOrder.b1f")>
    Friend Class SalesOrder
        Inherits SystemFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents objCombo As SAPbouiCOM.ComboBox
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.EditText0 = CType(Me.GetItem("4").Specific, SAPbouiCOM.EditText)
            Me.ComboBox0 = CType(Me.GetItem("10000330").Specific, SAPbouiCOM.ComboBox)
            Me.Matrix0 = CType(Me.GetItem("38").Specific, SAPbouiCOM.Matrix)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataAddAfter, AddressOf Me.Form_DataAddAfter

        End Sub

        Private WithEvents EditText0 As SAPbouiCOM.EditText

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("139", 0)

            Catch ex As Exception
            End Try
        End Sub

        Private Sub EditText0_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText0.ChooseFromListAfter
            Try
                objCombo = objform.Items.Item("10000330").Specific

                If Not objaddon.ValidateItemSync() Then
                    If objaddon.Validate_Transaction Then
                        objCombo.ValidValues.Add("GRPO", "GRPO")
                        objCombo.ExpandType = SAPbouiCOM.BoExpandType.et_ValueOnly
                    End If
                End If

            Catch ex As Exception

            End Try
        End Sub

        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox

        Private Sub ComboBox0_ComboSelectAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ComboBox0.ComboSelectAfter
            Try
                objCombo = objform.Items.Item("10000330").Specific
                If objCombo.Selected.Value = "GRPO" Then
                    Dim activeform As New FrmOpenList
                    activeform.Show()
                    activeform.LoadMatrixFromGRPO_Deliveries(objCombo.Selected.Value)
                    activeform.objform.Left = objform.Left + 100
                    activeform.objform.Top = objform.Top + 100
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                If pVal.ActionSuccess = True Then
                    Dim UpdateQuery As String = "", MarkUp As String = ""
                    MarkUp = objaddon.objglobalmethods.getSingleValue("select Top 1 ""U_MarkUp"" from ""@MI_MARKUP"" order by ""Code"" Desc")
                    Dim objRS As SAPbobsCOM.Recordset
                    objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    UpdateQuery = "Update T1 set T1.""U_MarkUp""='" & MarkUp & "' from RDR1 T1 join ORDR T0 on T1.""DocEntry""=T0.""DocEntry"" where T0.""DocEntry""='" & objform.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0) & "'"
                    objRS.DoQuery(UpdateQuery)
                    'UpdateQuery = " update ORDR Set ""U_IComNum""='" & GRPODocEntry & "' where ""DocEntry""='" & objform.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0) & "'"
                    'objRS.DoQuery(UpdateQuery)
                    objRS = Nothing
                    GRPODocEntry = Nothing
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private WithEvents Matrix0 As SAPbouiCOM.Matrix

        Private Sub Matrix0_ValidateAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ValidateAfter
            Try
                Dim Markup As Double
                Dim LineId As Integer
                Dim objRs As SAPbobsCOM.Recordset
                Dim StrQuery As String = "", DocEntry As String = ""
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Select Case pVal.ColUID
                    Case "U_MarkUp"
                        If pVal.ItemChanged = True And pVal.ActionSuccess = True Then
                            If Matrix0.Columns.Item("U_MarkUp").Cells.Item(pVal.Row).Specific.String <> "" Then
                                LineId = CInt(Matrix0.Columns.Item("U_GRLine").Cells.Item(pVal.Row).Specific.String)
                                DocEntry = Matrix0.Columns.Item("U_GRNum").Cells.Item(pVal.Row).Specific.String
                                StrQuery = "Select T1.""ItemCode"",T0.""DocRate"",T1.""Price"" from OPDN T0 join PDN1 T1 on T0.""DocEntry""=T1.""DocEntry"""
                                StrQuery += vbCrLf + " where T0.""DocEntry""= '" & DocEntry & "' and T1.""ItemCode""='" & Matrix0.Columns.Item("1").Cells.Item(pVal.Row).Specific.String & "'  and T1.""LineNum""=" & LineId & ""
                                objRs.DoQuery(StrQuery)
                                Markup = CDbl(Matrix0.Columns.Item("U_MarkUp").Cells.Item(pVal.Row).Specific.String)
                                Markup = ((CDbl(objRs.Fields.Item("Price").Value.ToString) * CDbl(objRs.Fields.Item("DocRate").Value.ToString)) + ((CDbl(objRs.Fields.Item("Price").Value.ToString) * CDbl(objRs.Fields.Item("DocRate").Value.ToString)) * (Markup / 100)))
                                Matrix0.Columns.Item("14").Cells.Item(pVal.Row).Specific.String = Markup
                            End If
                        End If
                End Select
                objRs = Nothing
            Catch ex As Exception
            End Try
        End Sub
    End Class
End Namespace
