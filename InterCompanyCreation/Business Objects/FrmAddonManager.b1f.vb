Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace InterCompanyCreation
    <FormAttribute("60059", "Business Objects/FrmAddonManager.b1f")>
    Friend Class FrmAddonManager
        Inherits SystemFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents objMatrix As SAPbouiCOM.Matrix
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("11").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            objform = objaddon.objapplication.Forms.GetForm("60059", 0)
            objMatrix = objform.Items.Item("3").Specific
        End Sub
        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            Try
                Dim addonname As String = "", GetAddonName As String = "", EnableToDisconnect As String
                Dim Stat As Boolean = False
                GetAddonName = objaddon.objglobalmethods.getSingleValue("Select Top 1 T1.""U_AddOnName"" from ""@DB_LIST"" T1 where  T1.""U_AddOnName""<>''")
                GetAddonName = GetAddonName & "*"
                For i As Integer = 1 To objMatrix.VisualRowCount
                    addonname = objMatrix.Columns.Item("1").Cells.Item(i).Specific.String
                    If objMatrix.IsRowSelected(i) Then
                        If addonname.ToUpper Like GetAddonName.ToUpper Then
                            Stat = True
                        End If
                    End If
                Next
                EnableToDisconnect = objaddon.objglobalmethods.getSingleValue("Select ""U_AddonDis"" from OADM")
                If EnableToDisconnect = "Y" Then
                    Stat = False
                End If
                If Stat = True Then
                    objaddon.objapplication.StatusBar.SetText("You are not authorized to disconnect the " & addonname & "  Add-on...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False : Exit Sub
                End If
            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
