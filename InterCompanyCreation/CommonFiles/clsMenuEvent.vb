Imports SAPbouiCOM
Namespace InterCompanyCreation

    Public Class clsMenuEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods

        Public Sub MenuEvent_For_StandardMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "150"
                        ItemMaster_MenuEvent(pVal, BubbleEvent)

                End Select
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Default_Sample_MenuEvent(ByVal pval As SAPbouiCOM.MenuEvent, ByVal BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                If pval.BeforeAction = True Then
                Else
                    Select Case pval.MenuUID
                        Case "1281"
                        Case Else
                    End Select
                End If
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub



        Public Sub RemovingItemNew(ByVal ItemCode As String)
            Try
                Dim Str As String
                Dim objRecordset As SAPbobsCOM.Recordset
                Dim objitem1 As SAPbobsCOM.Items
                Dim objcompany1 As SAPbobsCOM.Company
                Str = "select Distinct T1.""U_DBName"" from ""@DB_LIST"" T1  where  ifnull(""U_Select"",'')='Y'"
                objRecordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordset.DoQuery(Str)
                If objRecordset.RecordCount > 0 Then
                    While Not objRecordset.EoF
                        objcompany1 = objaddon.objglobalmethods.ConnectToCompany(objRecordset.Fields.Item("U_DBName").Value)
                        objitem1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                        If objitem1.GetByKey(objform.Items.Item("5").Specific.String) Then
                            If objitem1.Remove() <> 0 Then
                                objaddon.objapplication.SetStatusBarMessage(objcompany1.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                ' objAddOn.objApplication.MessageBox(ErrCode & " " & ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, vbOKCancel)
                            Else
                                objaddon.objapplication.SetStatusBarMessage("Item Removed from " & objRecordset.Fields.Item("U_DBName").Value, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            End If
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany1)
                            GC.Collect()
                        End If
                        objRecordset.MoveNext()

                    End While
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub ItemMaster_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim objCheck As SAPbouiCOM.CheckBox
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                objCheck = objform.Items.Item("Global").Specific

                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            'If objaddon.ValidateItemSync Then
                            objaddon.objapplication.SetStatusBarMessage("Removing Item is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False : Exit Sub
                            'Else
                            '    If objaddon.Validate_Master() Then
                            '        RemovingItemNew(objform.Items.Item("5").Specific.String)
                            '    End If
                            'End If
                        Case "1293"

                        Case "1292"

                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode  
                        Case "1282" ' Add Mode
                            Dim objcombo As SAPbouiCOM.ComboBox
                            objcombo = objform.Items.Item("1320002059").Specific
                            If objcombo.Selected.Description = "Saudi" Then
                                objCheck.Item.Enabled = True
                            Else
                                objCheck.Item.Enabled = False
                            End If

                            'objform.Items.Item("txtentry").Specific.string = objaddon.objglobalmethods.getSingleValue("select Count(*)+1 ""DocEntry"" from ""@MIPL_OBOM""")
                        Case "1288", "1289", "1290", "1291"
                            If objCheck.Checked = True Then
                                objCheck.Item.Enabled = False
                            Else
                                objCheck.Item.Enabled = True
                            End If
                            'objaddon.objapplication.Menus.Item("1300").Activate()
                        Case "1293"
                            'If objform.Mode = BoFormMode.fm_OK_MODE Then objform.Mode = BoFormMode.fm_UPDATE_MODE
                            'objform.Update()
                            'objform.Refresh()
                        Case "1292"

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub



    End Class
End Namespace