Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace InterCompanyCreation
    <FormAttribute("138", "Business Objects/FrmGeneralSettings.b1f")>
    Friend Class FrmGeneralSettings
        Inherits SystemFormBase

        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.CheckBox0 = CType(Me.GetItem("ChkAddon").Specific, SAPbouiCOM.CheckBox)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

        Private Sub OnCustomInitialize()

        End Sub
        Private WithEvents CheckBox0 As SAPbouiCOM.CheckBox
    End Class
End Namespace
