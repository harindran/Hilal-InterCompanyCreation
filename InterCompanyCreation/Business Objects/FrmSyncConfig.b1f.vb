Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace InterCompanyCreation
    <FormAttribute("MISYNC", "Business Objects/FrmSyncConfig.b1f")>
    Friend Class FrmSyncConfig
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents odbdsDetails As SAPbouiCOM.DBDataSource
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("MtxConfig").Specific, SAPbouiCOM.Matrix)
            Me.EditText1 = CType(Me.GetItem("txtName").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("lblName").Specific, SAPbouiCOM.StaticText)
            Me.StaticText0 = CType(Me.GetItem("lblDB").Specific, SAPbouiCOM.StaticText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("MISYNC", 0)
                objform = objaddon.objapplication.Forms.ActiveForm
                'EditText0.Value = objaddon.objglobalmethods.GetNextCode_Value("@MIPL_SYNC")
            Catch ex As Exception

            End Try
        End Sub
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix


        Public Sub LoadMatrix(ByVal FormUID As String)
            Dim objrs As SAPbobsCOM.Recordset
            Dim strsql As String = ""
            Dim i As Integer = 0
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strsql = "select ""U_DBName"" from ""@DB_LIST"" "
                objrs.DoQuery(strsql)
                objform.Freeze(True)
                'odbdsDetails = objform.DataSources.DBDataSources.Item(CType(1, Object))
                odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_SYNC1")
                Matrix0.Clear()
                odbdsDetails.Clear()
                If objrs.RecordCount > 0 Then
                    objaddon.objapplication.StatusBar.SetText("Loading Database List Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    While Not objrs.EoF
                        Matrix0.AddRow()
                        'odbdsDetails.Clear()
                        Matrix0.GetLineData(Matrix0.RowCount)
                        odbdsDetails.SetValue("LineId", 0, i + 1)
                        odbdsDetails.SetValue("U_DBName", 0, objrs.Fields.Item("U_DBName").Value.ToString)
                        Matrix0.SetLineData(Matrix0.RowCount)
                        objrs.MoveNext()
                        i += 1
                    End While
                End If
                objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception
            Finally
                objform.Freeze(False)
            End Try
        End Sub

        Private Function CheckExistingItem(ByVal ItemCode1 As String) As Boolean
            Dim Str As String
            Dim objRecordset As SAPbobsCOM.Recordset
            Str = "select ""Code"", ""U_ItemName"" from """ & objaddon.objcompany.CompanyDB & """.""@MIDBSL"" where ""Code"" ='" & ItemCode1 & "'"
            objRecordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecordset.DoQuery(Str)
            If objRecordset.RecordCount > 0 Then
                Return True
            Else
                Return False
            End If
        End Function

        Private Sub UpdatingSyncData(ByVal DBName As String, ByVal ItemCode As String)
            Dim Str As String
            Dim objRecordset As SAPbobsCOM.Recordset
            Try
                Str = "update " & objaddon.objcompany.CompanyDB & ".""@MIDBSL1"" set ""U_Sync""='Y' where ""U_DBName"" ='" & DBName & "' and ""Code""='" & ItemCode & "'"
                objRecordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordset.DoQuery(Str)
                'objAddOn.WriteSMSLog(str)
                objaddon.objapplication.SetStatusBarMessage("Sync updated", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
            End Try
        End Sub

        Private Sub DisablingRows(ByVal RowID As Integer)
            Matrix0.CommonSetting.SetCellEditable(RowID, 1, False)
            Matrix0.CommonSetting.SetCellEditable(RowID, 2, False)
            Matrix0.CommonSetting.SetCellEditable(RowID, 3, False)
            'Matrix0.CommonSetting.SetCellEditable(RowID, 4, False)
            'Matrix0.CommonSetting.SetCellEditable(RowID, 5, False)
        End Sub

        Private Sub Disabling()
            Dim chkselect As SAPbouiCOM.CheckBox
            For i = 1 To Matrix0.RowCount
                chkselect = Matrix0.Columns.Item("Select").Cells.Item(i).Specific
                If chkselect.Checked = True Then
                    DisablingRows(i)
                End If
            Next
        End Sub
        
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText

        Private Sub Button0_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.PressedAfter
            Try
                objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                objform.Items.Item("2").Click()
            Catch ex As Exception

            End Try
        End Sub
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText

        Private Sub Button1_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button1.ClickAfter
            Try
                'Dim objItemForm As SAPbouiCOM.Form
                'Dim objcheck As SAPbouiCOM.CheckBox
                'objItemForm = objaddon.objapplication.Forms.GetForm("150", 0)
                'objcheck = objItemForm.Items.Item("Global").Specific
                'objcheck.Item.Enabled = True
                'objcheck.Checked = False
            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
