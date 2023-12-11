Option Strict Off
Option Explicit On

Imports SAPbobsCOM
Imports SAPbouiCOM.Framework

Namespace InterCompanyCreation
    <FormAttribute("GETVAL", "Business Objects/FrmGetUserData.b1f")>
    Friend Class FrmGetUserData
        Inherits UserFormBase
        Dim FormCount As Integer = 0
        Private WithEvents objform As SAPbouiCOM.Form
        Dim Strq As String
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("btnok").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("btncan").Specific, SAPbouiCOM.Button)
            Me.EditText0 = CType(Me.GetItem("tpass").Specific, SAPbouiCOM.EditText)
            Me.StaticText0 = CType(Me.GetItem("lpass").Specific, SAPbouiCOM.StaticText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler CloseBefore, AddressOf Me.Form_CloseBefore

        End Sub

        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("GETVAL", FormCount)
                bModal = True
                objform.Freeze(True)
                objform.Left = AlignLeft + (AlignLeft / 2)
                objform.Top =  (AlignTop / 2)
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                'Dim Res As SAPbobsCOM.AuthenticateUserResultsEnum = objaddon.objcompany.AuthenticateUser(objaddon.objcompany.UserName, EditText0.Value)
                'If Not Res = AuthenticateUserResultsEnum.aturUsernamePasswordMatch Then
                '    objaddon.objapplication.StatusBar.SetText("Incorrect Password...Please enter the valid password..", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    BubbleEvent = False : Exit Sub
                'End If
                Strq = objaddon.objglobalmethods.getSingleValue("select Distinct T1.""U_DBName"" from ""@DB_LIST"" T1  where  ifnull(""U_Select"",'')='Y' and T1.""U_DBName""<>''")
                If GetValidCompany(Strq) = False Then
                    objaddon.objapplication.StatusBar.SetText("Incorrect Password...Please enter the valid password..", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False : Exit Sub
                End If
                SAPPassWord = EditText0.Value
                objform.Close()
            Catch ex As Exception
            End Try

        End Sub

        Private Function GetValidCompany(ByVal DBName As String) As Boolean
            Dim lRetCode As Integer
            Dim sErrMsg As String = ""
            Dim objRecordset As SAPbobsCOM.Recordset
            Dim strQuery As String
            Dim objcompanynew As SAPbobsCOM.Company
            Try
                strQuery = "Select ""U_DBName"", ""U_UserName"", ""U_Password"", ""U_DBUser"",""U_DBPass"",""U_Server"",""U_LicServer"" from """ & objaddon.objcompany.CompanyDB & """.""@DB_LIST"" where ""U_DBName""='" & DBName & "' and ifnull(""U_Select"",'')='Y' and ""U_DBName""<>''"
                objRecordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordset.DoQuery(strQuery)
                objcompanynew = New SAPbobsCOM.Company
                objcompanynew.Server = Trim(objRecordset.Fields.Item("U_Server").Value)
                objcompanynew.LicenseServer = Trim(objRecordset.Fields.Item("U_LicServer").Value)
                objcompanynew.SLDServer = Trim(objRecordset.Fields.Item("U_LicServer").Value)
                objcompanynew.language = SAPbobsCOM.BoSuppLangs.ln_English
                objcompanynew.UseTrusted = False
                objcompanynew.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
                objcompanynew.DbUserName = Trim(objRecordset.Fields.Item("U_DBUser").Value)
                objcompanynew.DbPassword = Trim(objRecordset.Fields.Item("U_DBPass").Value)
                objcompanynew.CompanyDB = Trim(objRecordset.Fields.Item("U_DBName").Value)
                objcompanynew.UserName = objaddon.objcompany.UserName
                objcompanynew.Password = EditText0.Value 'SAPPassWord

                If objcompanynew.Connected = True Then
                    objcompanynew.Disconnect()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompanynew)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompanynew)
                    objcompanynew = Nothing
                    Return True
                End If
                lRetCode = objcompanynew.Connect()
                If lRetCode <> 0 Then
                    EditText0.Value = ""
                    SAPPassWord = ""
                    Return False
                Else
                    objcompanynew.Disconnect()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompanynew)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompanynew)
                    objcompanynew = Nothing
                    'objaddon.objapplication.SetStatusBarMessage("Connected to " & objcompanynew.CompanyDB, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    Return True
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                'MsgBox(ex.ToString)
            End Try
            Return Nothing
        End Function

        Private Sub Button1_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button1.ClickBefore
            Try
                objform.Close()
            Catch ex As Exception
            End Try
        End Sub

        Private Sub Form_CloseBefore(pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)
            Try
                If pVal.InnerEvent = False Then Exit Sub
                BubbleEvent = False
            Catch ex As Exception
            End Try

        End Sub
    End Class
End Namespace
