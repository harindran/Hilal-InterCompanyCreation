Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace InterCompanyCreation
    <FormAttribute("142", "Business Objects/PurchaseOrder.b1f")>
    Friend Class PurchaseOrder
        Inherits SystemFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents objCombo As SAPbouiCOM.ComboBox
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.EditText0 = CType(Me.GetItem("4").Specific, SAPbouiCOM.EditText)
            Me.ComboBox0 = CType(Me.GetItem("10000330").Specific, SAPbouiCOM.ComboBox)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataAddAfter, AddressOf Me.Form_DataAddAfter

        End Sub

        Private WithEvents EditText0 As SAPbouiCOM.EditText

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("142", 0)
            Catch ex As Exception
            End Try
        End Sub

        Private Sub EditText0_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText0.ChooseFromListAfter
            Try
                objCombo = objform.Items.Item("10000330").Specific
                If objaddon.ValidateItemSync() Then
                    If objaddon.Validate_Transaction Then
                        objCombo.ValidValues.Add("Delivery", "Delivery")
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
                If objCombo.Selected.Value = "Delivery" Then
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
                'Dim UpdateQuery As String = ""
                'If pVal.ActionSuccess = True Then
                '    Dim objRS As SAPbobsCOM.Recordset
                '    objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '    'UpdateQuery = " update OPOR Set ""U_IComNum""='" & DeliveryDocEntry & "' where ""DocEntry""='" & objform.DataSources.DBDataSources.Item("OPOR").GetValue("DocEntry", 0) & "'"
                '    'objRS.DoQuery(UpdateQuery)
                '    objRS = Nothing
                '    DeliveryDocEntry = Nothing
                'End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub
    End Class
End Namespace
