Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace InterCompanyCreation
    <FormAttribute("228", "Business Objects/FrmDocSettings.b1f")>
    Friend Class FrmDocSettings
        Inherits SystemFormBase

        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.CheckBox0 = CType(Me.GetItem("ItemSync").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox1 = CType(Me.GetItem("ChkMaster").Specific, SAPbouiCOM.CheckBox)
            Me.CheckBox2 = CType(Me.GetItem("ChkTran").Specific, SAPbouiCOM.CheckBox)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub
        Private WithEvents CheckBox0 As SAPbouiCOM.CheckBox

        Private Sub OnCustomInitialize()

        End Sub
        Private WithEvents CheckBox1 As SAPbouiCOM.CheckBox
        Private WithEvents CheckBox2 As SAPbouiCOM.CheckBox
    End Class
End Namespace
