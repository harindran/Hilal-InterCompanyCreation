Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SAPbobsCOM


Namespace InterCompanyCreation
    <FormAttribute("150", "Business Objects/FrmItemMaster.b1f")>
    Friend Class FrmItemMaster
        Inherits SystemFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Private Shared FormCount As Integer = 0
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.CheckBox0 = CType(Me.GetItem("Global").Specific, SAPbouiCOM.CheckBox)
            Me.EditText0 = CType(Me.GetItem("18").Specific, SAPbouiCOM.EditText)
            Me.ComboBox0 = CType(Me.GetItem("10002056").Specific, SAPbouiCOM.ComboBox)
            Me.ComboBox1 = CType(Me.GetItem("39").Specific, SAPbouiCOM.ComboBox)
            Me.ComboBox2 = CType(Me.GetItem("214").Specific, SAPbouiCOM.ComboBox)
            Me.ComboBox3 = CType(Me.GetItem("1320002059").Specific, SAPbouiCOM.ComboBox)
            Me.ComboBox4 = CType(Me.GetItem("cmbType").Specific, SAPbouiCOM.ComboBox)
            Me.StaticText0 = CType(Me.GetItem("IType").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox5 = CType(Me.GetItem("24").Specific, SAPbouiCOM.ComboBox)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataLoadAfter, AddressOf Me.FrmItemMaster_DataLoadAfter
            AddHandler DataLoadAfter, AddressOf Me.FrmItemMaster_DataLoadAfter
            AddHandler DataLoadAfter, AddressOf Me.FrmItemMaster_DataLoadAfter
            AddHandler DataAddAfter, AddressOf Me.Form_DataAddAfter
            AddHandler DataLoadAfter, AddressOf Me.FrmItemMaster_DataLoadAfter
            AddHandler DataLoadAfter, AddressOf Me.FrmItemMaster_DataLoadAfter
            AddHandler DataLoadAfter, AddressOf Me.FrmItemMaster_DataLoadAfter
            AddHandler DataLoadAfter, AddressOf Me.FrmItemMaster_DataLoadAfter
            AddHandler CloseAfter, AddressOf Me.Form_CloseAfter
            AddHandler DataLoadAfter, AddressOf Me.FrmItemMaster_DataLoadAfter


        End Sub

        Private Sub OnCustomInitialize()
            Try
                FormCount += 1
                objform = objaddon.objapplication.Forms.GetForm("150", FormCount)

                'CheckBox0.Item.Enabled = True
                If objaddon.ValidateItemSync Then
                    CheckBox0.Item.Visible = False
                Else
                    If objaddon.Validate_Master Then
                        CheckBox0.Item.Visible = True
                    Else
                        CheckBox0.Item.Visible = False
                    End If
                End If

                ComboBox4.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                'ComboBox4.Select("-", SAPbouiCOM.BoSearchKey.psk_ByDescription)
            Catch ex As Exception
            End Try
        End Sub

        Public Sub CreateCheckBox()

            Dim objCheckbox As SAPbouiCOM.CheckBox
            Dim objItem As SAPbouiCOM.Item
            Try
                objform = objaddon.objapplication.Forms.GetForm("150", 0)
                objItem = objform.Items.Add("BtnCheck", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
                objItem.Left = objform.Items.Item("39").Left + objform.Items.Item("39").Width + 110
                objItem.Width = 100
                objItem.Top = objform.Items.Item("39").Top
                objItem.Height = objform.Items.Item("39").Height
                objCheckbox = objItem.Specific
                objCheckbox.Caption = "Global Item"
                objCheckbox.DataBind.SetBound(True, "OITM", "U_Global")
                objCheckbox.Item.FontSize = 14
                objaddon.objapplication.SetStatusBarMessage("CheckBox Created", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Catch ex As Exception
            End Try

        End Sub

        Private Function AddItemToDestinationDB(ByVal ItemCode As String) As Boolean
            Dim RetVal As Long
            Dim Str, StrItmG As String
            Dim objitem As SAPbobsCOM.Items
            Dim objitem1 As SAPbobsCOM.Items
            'Dim objitemWhs As SAPbobsCOM.ItemWarehouseInfo
            Dim objRecordset, objRs, objRsItmg As SAPbobsCOM.Recordset
            Dim Series, ItemSyncFlag As String
            Dim objcompany1 As SAPbobsCOM.Company
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim oSeriesService As SAPbobsCOM.SeriesService
            Dim oSeries As SAPbobsCOM.Series
            Dim oSeries1 As SAPbobsCOM.SeriesCollection
            Dim oDocumentTypeParams As SAPbobsCOM.DocumentTypeParams
            Dim ValidSeries As Boolean = False, Flag As Boolean = False
            Dim SeriesName As String = "", ItemMasterSeries As String = "", ItemMasterSeriesName As String = ""

            Try
                objform = objaddon.objapplication.Forms.Item(objform.UniqueID)
                ' Str = "select T1.""Code"",T1.""U_Select"",T1.""U_DBName"" from """ & objaddon.objcompany.CompanyDB & """.""@MIPL_SYNC1"" T1  where T1.""Code""='" & ItemCode & "' and ifnull(T1.""U_Select"",'')='Y'"
                Str = "select Distinct T1.""U_DBName"" from ""@DB_LIST"" T1  where  ifnull(T1.""U_Select"",'')='Y' and T1.""U_DBName""<>''"
                objRecordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRsItmg = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordset.DoQuery(Str)

                If ItemCode <> "" Then

                    objitem = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    oCmpSrv = objaddon.objcompany.GetCompanyService
                    oSeriesService = oCmpSrv.GetBusinessService(ServiceTypes.SeriesService)
                    oSeries = oSeriesService.GetDataInterface(SeriesServiceDataInterfaces.ssdiSeries)
                    oDocumentTypeParams = oSeriesService.GetDataInterface(SeriesServiceDataInterfaces.ssdiDocumentTypeParams)
                    oDocumentTypeParams.Document = "4"
                    oSeries1 = oSeriesService.GetDocumentSeries(oDocumentTypeParams)
                    ItemMasterSeries = objform.DataSources.DBDataSources.Item("OITM").GetValue("Series", 0)
                    ' CurrentItemCode = objaddon.objglobalmethods.getSingleValue("select ""NextNumber""-1 from NNM1 where ""ObjectCode""='4' and ""Series""='" & ItemMasterSeries & "'")
                    ItemSyncFlag = objform.DataSources.DBDataSources.Item("OITM").GetValue("U_Global", 0)
                    If ItemSyncFlag = "N" Then
                        objaddon.objapplication.SetStatusBarMessage("Please select the Saudi Series or check the add-on connection...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Flag = False
                        Return Flag
                    End If
                    ItemMasterSeriesName = objaddon.objglobalmethods.getSingleValue("select ""SeriesName"" from NNM1 where ""ObjectCode""='4' and ""Series""='" & ItemMasterSeries & "'")
                    For i As Integer = 0 To oSeries1.Count - 1
                        If ItemMasterSeriesName = "Saudi" Then  '"Saudi"
                            SeriesName = oSeries1.Item(i).Name
                            If SeriesName = "Saudi" Then
                                ValidSeries = True
                                Exit For
                            End If
                        End If
                    Next
                    If ValidSeries Then
                        If objRecordset.RecordCount > 0 Then
                            While Not objRecordset.EoF
                                objaddon.objapplication.SetStatusBarMessage("Connecting to " & objRecordset.Fields.Item("U_DBName").Value, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                objcompany1 = objaddon.objglobalmethods.ConnectToCompany(objRecordset.Fields.Item("U_DBName").Value)
                                
                                objRs = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Series = "select ""Series"" from NNM1 where ""ObjectCode""='4' and ""SeriesName""='" & SeriesName & "'"
                                objRs.DoQuery(Series)
                                If objRs.RecordCount > 0 Then
                                    Series = objRs.Fields.Item(0).Value.ToString()
                                Else
                                    Series = ""
                                End If
                                'DItemCode = "select ""NextNumber"" from NNM1 where ""ObjectCode""='4' and ""Series""='" & Series & "'"
                                'objRs.DoQuery(DItemCode)
                                'DItemCode = objRs.Fields.Item(0).Value.ToString()
                                'If Not CurrentItemCode = DItemCode Then
                                '    objaddon.objapplication.SetStatusBarMessage("Seems Sequence Itemcode Not Found.Please Check in from & to database...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                '    Flag = False
                                '    Return Flag
                                'End If
                                objitem.GetByKey(ItemCode.Trim)
                                objitem1 = objcompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                                If Series <> "" Then
                                    objitem1.Series = Series
                                Else
                                    objaddon.objapplication.SetStatusBarMessage("Series Not defined for Item Creation...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    Flag = False
                                    Return Flag
                                End If
                                objitem1.ItemCode = objitem.ItemCode
                                objitem1.ItemName = objitem.ItemName
                                objitem1.UserFields.Fields.Item("U_Model").Value = objitem.UserFields.Fields.Item("U_Model").Value
                                objitem1.UserFields.Fields.Item("U_Type").Value = objitem.UserFields.Fields.Item("U_Type").Value
                                objitem1.UserFields.Fields.Item("U_IType").Value = objitem.UserFields.Fields.Item("U_IType").Value
                                objitem1.UserFields.Fields.Item("U_ItemSync").Value = "Y"
                                objitem1.UserFields.Fields.Item("U_AddtExp").Value = objitem.UserFields.Fields.Item("U_AddtExp").Value
                                objitem1.UserFields.Fields.Item("U_StkItem").Value = objitem.UserFields.Fields.Item("U_StkItem").Value
                                objitem1.ForeignName = objitem.ForeignName
                                objitem1.ItemsGroupCode = objitem.ItemsGroupCode
                                objitem1.ChapterID = objitem.ChapterID
                                objitem1.SalesUnitWeight = objitem.SalesUnitWeight
                                objitem1.InventoryUOM = objitem.InventoryUOM
                                objitem1.PurchaseUnit = objitem.PurchaseUnit
                                objitem1.SalesUnit = objitem.SalesUnit
                                objitem1.ManageBatchNumbers = objitem.ManageBatchNumbers
                                objitem1.ManageSerialNumbers = objitem.ManageSerialNumbers
                                objitem1.TaxType = objitem.TaxType
                                objitem1.IsPhantom = objitem.IsPhantom
                                objitem1.PlanningSystem = objitem.PlanningSystem
                                objitem1.ProcurementMethod = objitem.ProcurementMethod
                                objitem1.ServiceCategoryEntry = objitem.ServiceCategoryEntry
                                objitem1.ItemClass = objitem.ItemClass
                                objitem1.SACEntry = objitem.SACEntry
                                'objitem1.VatLiable = objitem.VatLiable
                                'objitem1.WTLiable = objitem.WTLiable
                                objitem1.UoMGroupEntry = objitem.UoMGroupEntry
                                objitem1.SalesVATGroup = objitem.SalesVATGroup
                                objitem1.SalesVolumeUnit = objitem.SalesVolumeUnit
                                objitem1.PurchaseVolumeUnit = objitem.PurchaseVolumeUnit
                                objitem1.ManageStockByWarehouse = objitem.ManageStockByWarehouse
                                objitem1.MaxInventory = objitem.MaxInventory
                                objitem1.MinInventory = objitem.MinInventory
                                objitem1.MinOrderQuantity = objitem.MinOrderQuantity
                                objitem1.Valid = objitem.Valid
                                objitem1.ValidFrom = objitem.ValidFrom
                                objitem1.ValidTo = objitem.ValidTo
                                objitem1.ValidRemarks = objitem.ValidRemarks

                                objitem1.Frozen = objitem.Frozen
                                objitem1.FrozenFrom = objitem.FrozenFrom
                                objitem1.FrozenTo = objitem.FrozenTo
                                objitem1.FrozenRemarks = objitem.FrozenRemarks

                                objitem1.TypeOfAdvancedRules = objitem.TypeOfAdvancedRules

                                'objitem1.GSTRelevnt = objitem.GSTRelevnt
                                objitem1.GLMethod = objitem.GLMethod
                                objitem1.ItemType = objitem.ItemType
                                'objitem1.GSTTaxCategory = objitem.GSTTaxCategory
                                objitem1.MaterialType = objitem.MaterialType
                                objitem1.PurchaseItem = objitem.PurchaseItem
                                objitem1.SalesItem = objitem.SalesItem
                                objitem1.InventoryItem = objitem.InventoryItem
                                objitem1.ShipType = objitem.ShipType

                                objitem1.Manufacturer = objitem.Manufacturer

                                objitem1.PurchasePackagingUnit = objitem.PurchasePackagingUnit
                                objitem1.PurchaseQtyPerPackUnit = objitem.PurchaseQtyPerPackUnit
                                objitem1.SalesPackagingUnit = objitem.SalesPackagingUnit
                                objitem1.SalesQtyPerPackUnit = objitem.SalesQtyPerPackUnit

                                If objitem.ManageSerialNumbers = BoYesNoEnum.tNO And objitem.ManageBatchNumbers = BoYesNoEnum.tNO Then
                                    StrItmG = "Select ""Name"" from ""@MI_ITMSL"" where ""Code""<>''"
                                    objRsItmg.DoQuery(StrItmG)
                                    For Rec As Integer = 0 To objRsItmg.RecordCount - 1
                                        If objitem.ItemsGroupCode = objRsItmg.Fields.Item("Name").Value Then
                                            objitem1.ManageSerialNumbers = BoYesNoEnum.tYES
                                            objitem1.SRIAndBatchManageMethod = BoManageMethod.bomm_OnReleaseOnly
                                            objitem1.CostAccountingMethod = BoInventorySystem.bis_MovingAverage
                                            objitem.IssueMethod = BoIssueMethod.im_Manual
                                        End If
                                        objRsItmg.MoveNext()
                                    Next
                                Else
                                    objitem1.SRIAndBatchManageMethod = objitem.SRIAndBatchManageMethod
                                    objitem1.CostAccountingMethod = objitem.CostAccountingMethod
                                    objitem1.IssueMethod = objitem.IssueMethod
                                End If
                                Dim ItemCod As String = objform.Items.Item("5").Specific.String
                                RetVal = 0
                                If objitem1.GetByKey(objform.Items.Item("5").Specific.String) Then
                                    RetVal = objitem1.Update()
                                Else
                                    RetVal = objitem1.Add()
                                End If
                                Dim objItemWarehouse As SAPbouiCOM.DBDataSource
                                objItemWarehouse = objform.DataSources.DBDataSources.Item("OITW")
                                If RetVal <> 0 Then
                                    objaddon.objapplication.SetStatusBarMessage(objcompany1.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    ' objAddOn.objApplication.MessageBox(ErrCode & " " & ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, vbOKCancel)
                                Else
                                    'objaddon.objapplication.SetStatusBarMessage("Item Added to " & objRecordset.Fields.Item("U_DBName").Value, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                    If objitem1.GetByKey(ItemCod) Then
                                        Dim Whs As String = ""
                                        Dim objitemWhs As SAPbobsCOM.ItemWarehouseInfo
                                        For j As Integer = 0 To objitem.WhsInfo.Count - 1
                                            Whs = objaddon.objglobalmethods.getSingleValue("select Top 1 ""Name"" from ""@MI_WHSE"" where ""Code""='" & objItemWarehouse.GetValue("WhsCode", j) & "'") 'objItemWarehouse.GetValue("WhsCode", j)
                                            objitem.WhsInfo.SetCurrentLine(j)
                                            For k As Integer = 0 To objitem1.WhsInfo.Count - 1
                                                objitemWhs = objitem1.WhsInfo
                                                objitemWhs.SetCurrentLine(k)
                                                If objitemWhs.WarehouseCode = Whs Then
                                                    objitem1.WhsInfo.SetCurrentLine(k)
                                                    objitemWhs.MaximalStock = objitem.WhsInfo.MaximalStock
                                                    objitemWhs.MinimalStock = objitem.WhsInfo.MinimalStock
                                                    objitemWhs.MinimalOrder = objitem.WhsInfo.MinimalOrder
                                                End If
                                            Next
                                        Next
                                        RetVal = objitem1.Update()
                                        If RetVal <> 0 Then
                                            objaddon.objapplication.SetStatusBarMessage(objcompany1.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        Else
                                            objaddon.objapplication.SetStatusBarMessage("Item Added to " & objRecordset.Fields.Item("U_DBName").Value, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                            Flag = True
                                        End If
                                    End If
                                End If
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany1)
                                objcompany1 = Nothing
                                GC.Collect()
                                objRecordset.MoveNext()
                            End While
                        End If
                    Else
                        objaddon.objapplication.SetStatusBarMessage("Item Not Synced.. Series should be Saudi...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Flag = False
                        Return Flag
                    End If
                Else
                    Return Flag
                End If

            Catch ex As Exception
                Flag = False
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
            Return Flag
        End Function

        Public Sub RemovingItem(ByVal ItemCode As String)
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
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub UpdateItemToDestinationDB(ByVal ItemCode As String)
            Try
                Dim RetVal As Long
                Dim ErrCode As Long
                Dim ErrMsg As String, Str As String
                Dim Targetcompany As SAPbobsCOM.Company
                Dim objitem As SAPbobsCOM.Items
                Dim objitem1 As SAPbobsCOM.Items
                Dim objRecordset, objRsItmg As SAPbobsCOM.Recordset
                Dim objItemHeader As SAPbouiCOM.DBDataSource
                'Dim i As Integer
                objform = objaddon.objapplication.Forms.Item(objform.UniqueID)
                Try
                    'Str = "select T1.""Code"",T1.""U_Select"",T1.""U_DBName"" from """ & objaddon.objcompany.CompanyDB & """.""@MIPL_SYNC1"" T1  where T1.""Code""='" & ItemCode & "' and ifnull(""U_Select"",'')='Y'"
                    Str = "select Distinct T1.""U_DBName"" from ""@DB_LIST"" T1  where  ifnull(""U_Select"",'')='Y'"
                    objRecordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    objRsItmg = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    objRecordset.DoQuery(Str)
                    objitem = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    objaddon.objapplication.SetStatusBarMessage("Updating Item Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    If objRecordset.RecordCount > 0 Then
                        While Not objRecordset.EoF
                            Targetcompany = objaddon.objglobalmethods.ConnectToCompany(objRecordset.Fields.Item("U_DBName").Value)
                            objitem1 = Targetcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            objItemHeader = objform.DataSources.DBDataSources.Item("OITM")
                            Dim objItemWarehouse As SAPbouiCOM.DBDataSource
                            objItemWarehouse = objform.DataSources.DBDataSources.Item("OITW")
                            Dim itmstr As String = ItemCode.Trim 'objform.Items.Item("5").Specific.string
                            'objitem.GetByKey(Trim(objform.Items.Item("5").Specific.string))
                            If Not objitem1.GetByKey(Trim(itmstr)) Then
                                Try
                                    If Targetcompany.Connected Then Targetcompany.Disconnect()
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Targetcompany)
                                    Targetcompany = Nothing
                                    GC.Collect()
                                    AddItemToDestinationDB(Trim(ItemCode))
                                    Exit Sub
                                Catch ex As Exception
                                End Try
                            End If
                            If objitem1.GetByKey(Trim(ItemCode)) Then
                                objitem1.ItemCode = ItemCode.Trim ' objform.Items.Item("5").Specific.string
                                objitem1.ItemName = objItemHeader.GetValue("ItemName", 0)
                                objitem1.UserFields.Fields.Item("U_Model").Value = objItemHeader.GetValue("U_Model", 0)
                                objitem1.UserFields.Fields.Item("U_Type").Value = objItemHeader.GetValue("U_Type", 0)
                                objitem1.UserFields.Fields.Item("U_IType").Value = objItemHeader.GetValue("U_IType", 0) 'ComboBox4.Selected.Description
                                objitem1.UserFields.Fields.Item("U_AddtExp").Value = objItemHeader.GetValue("U_AddtExp", 0)
                                objitem1.UserFields.Fields.Item("U_StkItem").Value = objItemHeader.GetValue("U_StkItem", 0)
                                objitem1.UserFields.Fields.Item("U_ItemSync").Value = "Y"
                                objitem1.ForeignName = objItemHeader.GetValue("FrgnName", 0)
                                objitem1.ItemsGroupCode = objItemHeader.GetValue("ItmsGrpCod", 0)
                                If objItemHeader.GetValue("PrchseItem", 0) = "Y" Then
                                    objitem1.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tYES
                                Else
                                    objitem1.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tNO
                                End If
                                If objItemHeader.GetValue("InvntItem", 0) = "Y" Then
                                    objitem1.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES
                                Else
                                    objitem1.InventoryItem = SAPbobsCOM.BoYesNoEnum.tNO
                                End If
                                If objItemHeader.GetValue("SellItem", 0) = "Y" Then
                                    objitem1.SalesItem = SAPbobsCOM.BoYesNoEnum.tYES
                                Else
                                    objitem1.SalesItem = SAPbobsCOM.BoYesNoEnum.tNO
                                End If
                                objitem1.PurchaseUnit = objItemHeader.GetValue("BuyUnitMsr", 0)
                                If objItemHeader.GetValue("ByWh", 0) = "Y" Then
                                    objitem1.ManageStockByWarehouse = SAPbobsCOM.BoYesNoEnum.tYES
                                Else
                                    objitem1.ManageStockByWarehouse = SAPbobsCOM.BoYesNoEnum.tNO
                                End If
                                Dim Whs As String = ""

                                Dim objitemWhs As SAPbobsCOM.ItemWarehouseInfo

                                For j As Integer = 0 To objitem.WhsInfo.Count - 1
                                    Whs = objaddon.objglobalmethods.getSingleValue("select Top 1 ""Name"" from ""@MI_WHSE"" where ""Code""='" & objItemWarehouse.GetValue("WhsCode", j) & "'") 'objItemWarehouse.GetValue("WhsCode", j)
                                    For k As Integer = 0 To objitem1.WhsInfo.Count - 1
                                        objitemWhs = objitem1.WhsInfo
                                        objitemWhs.SetCurrentLine(k)
                                        If objitemWhs.WarehouseCode = Whs Then
                                            objitemWhs.MaximalStock = objItemWarehouse.GetValue("MaxStock", j) 'objitem.WhsInfo.MaximalStock
                                            objitemWhs.MinimalStock = objItemWarehouse.GetValue("MinStock", j) 'objitem.WhsInfo.MinimalStock
                                            objitemWhs.MinimalOrder = objItemWarehouse.GetValue("MinOrder", j) 'objitem.WhsInfo.MinimalOrder
                                        End If
                                    Next
                                    objitem.WhsInfo.SetCurrentLine(j)
                                Next
                                objitem1.Manufacturer = objItemHeader.GetValue("FirmCode", 0)
                                objitem1.SalesUnit = objItemHeader.GetValue("SalUnitMsr", 0)
                                objitem1.ShipType = objItemHeader.GetValue("ShipType", 0)
                                If objItemHeader.GetValue("GLMethod", 0) = "C" Then
                                    objitem1.GLMethod = SAPbobsCOM.BoGLMethods.glm_ItemClass
                                ElseIf objItemHeader.GetValue("GLMethod", 0) = "W" Then
                                    objitem1.GLMethod = SAPbobsCOM.BoGLMethods.glm_WH
                                Else
                                    objitem1.GLMethod = SAPbobsCOM.BoGLMethods.glm_ItemLevel
                                End If
                                'If objItemHeader.GetValue("TaxType", 0) = "N" Then
                                '    objitem1.TaxType = SAPbobsCOM.BoTaxTypes.tt_No
                                'ElseIf objItemHeader.GetValue("TaxType", 0) = "Y" Then
                                '    objitem1.TaxType = SAPbobsCOM.BoTaxTypes.tt_Yes   ' 2 pending
                                'End If
                                If objItemHeader.GetValue("validFor", 0) = "Y" Then
                                    objitem1.Valid = SAPbobsCOM.BoYesNoEnum.tYES
                                Else
                                    objitem1.Valid = SAPbobsCOM.BoYesNoEnum.tNO
                                End If
                                If objItemHeader.GetValue("frozenFor", 0) = "Y" Then
                                    objitem1.Frozen = SAPbobsCOM.BoYesNoEnum.tYES
                                Else
                                    objitem1.Frozen = SAPbobsCOM.BoYesNoEnum.tNO
                                End If
                                If objItemHeader.GetValue("GLPickMeth", 0) = "A" Then
                                    objitem1.TypeOfAdvancedRules = SAPbobsCOM.TypeOfAdvancedRulesEnum.toar_General
                                ElseIf objItemHeader.GetValue("GLPickMeth", 0) = "C" Then
                                    objitem1.TypeOfAdvancedRules = SAPbobsCOM.TypeOfAdvancedRulesEnum.toar_ItemGroup
                                Else
                                    objitem1.TypeOfAdvancedRules = SAPbobsCOM.TypeOfAdvancedRulesEnum.toar_Warehouse
                                End If
                                If objItemHeader.GetValue("validFrom", 0) <> "" Then
                                    objitem1.ValidFrom = Date.ParseExact(objItemHeader.GetValue("validFrom", 0), "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                End If
                                If objItemHeader.GetValue("validTo", 0) <> "" Then
                                    objitem1.ValidTo = Date.ParseExact(objItemHeader.GetValue("validTo", 0), "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                End If

                                objitem1.ValidRemarks = objItemHeader.GetValue("validComm", 0)
                                If objItemHeader.GetValue("frozenFrom", 0) <> "" Then
                                    objitem1.FrozenFrom = Date.ParseExact(objItemHeader.GetValue("frozenFrom", 0), "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                End If
                                If objItemHeader.GetValue("frozenTo", 0) <> "" Then
                                    objitem1.FrozenTo = Date.ParseExact(objItemHeader.GetValue("frozenTo", 0), "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                End If

                                objitem1.FrozenRemarks = objItemHeader.GetValue("frozenComm", 0)

                                If objItemHeader.GetValue("ItemType", 0) = "I" Then
                                    objitem1.ItemType = SAPbobsCOM.ItemTypeEnum.itItems
                                ElseIf objItemHeader.GetValue("ItemType", 0) = "L" Then
                                    objitem1.ItemType = SAPbobsCOM.ItemTypeEnum.itLabor
                                ElseIf objItemHeader.GetValue("ItemType", 0) = "T" Then
                                    objitem1.ItemType = SAPbobsCOM.ItemTypeEnum.itTravel
                                Else
                                    objitem1.ItemType = SAPbobsCOM.ItemTypeEnum.itFixedAssets
                                End If
                                If objItemHeader.GetValue("Phantom", 0) = "N" Then
                                    objitem1.IsPhantom = SAPbobsCOM.BoYesNoEnum.tNO
                                Else
                                    objitem1.IsPhantom = SAPbobsCOM.BoYesNoEnum.tYES
                                End If
                                objitem1.InventoryUOM = objItemHeader.GetValue("InvntryUom", 0)

                                If objItemHeader.GetValue("PlaningSys", 0) = "M" Then
                                    objitem1.PlanningSystem = SAPbobsCOM.BoPlanningSystem.bop_MRP
                                Else
                                    objitem1.PlanningSystem = SAPbobsCOM.BoPlanningSystem.bop_None
                                End If
                                If objItemHeader.GetValue("PrcrmntMtd", 0) = "B" Then
                                    objitem1.ProcurementMethod = SAPbobsCOM.BoProcurementMethod.bom_Buy
                                Else
                                    objitem1.ProcurementMethod = SAPbobsCOM.BoProcurementMethod.bom_Make
                                End If
                                objitem1.ChapterID = objItemHeader.GetValue("ChapterID", 0)
                                objitem1.SalesUnitWeight = objItemHeader.GetValue("SWeight1", 0)
                                objitem1.MaterialType = objItemHeader.GetValue("MatType", 0)
                                objitem1.ServiceCategoryEntry = objItemHeader.GetValue("ServiceCtg", 0)
                                'If objItemHeader.GetValue("ItemClass", 0) = "2" Then
                                '    objitem1.ItemClass = SAPbobsCOM.ItemClassEnum.itcMaterial
                                'Else
                                '    objitem1.ItemClass = SAPbobsCOM.ItemClassEnum.itcService
                                'End If                               
                                'If objItemHeader.GetValue("GSTRelevnt", 0) = "Y" Then
                                '    objitem1.GSTRelevnt = SAPbobsCOM.BoYesNoEnum.tYES
                                'Else
                                '    objitem1.GSTRelevnt = SAPbobsCOM.BoYesNoEnum.tNO
                                'End If

                                objitem1.SACEntry = objItemHeader.GetValue("SACEntry", 0)
                                objitem1.UoMGroupEntry = objItemHeader.GetValue("UgpEntry", 0)
                                objitem1.SalesVATGroup = objItemHeader.GetValue("VatGourpSa", 0)
                                objitem1.SalesVolumeUnit = objItemHeader.GetValue("NumInSale", 0)
                                objitem1.PurchaseVolumeUnit = objItemHeader.GetValue("NumInBuy", 0)
                                'If objItemHeader.GetValue("GstTaxCtg", 0) = "R" Then
                                '    objitem1.GSTTaxCategory = SAPbobsCOM.GSTTaxCategoryEnum.gtc_Regular
                                'ElseIf objItemHeader.GetValue("GstTaxCtg", 0) = "N" Then
                                '    objitem1.GSTTaxCategory = SAPbobsCOM.GSTTaxCategoryEnum.gtc_NilRated
                                'Else
                                '    objitem1.GSTTaxCategory = SAPbobsCOM.GSTTaxCategoryEnum.gtc_Exempt
                                'End If
                                objitem1.PurchasePackagingUnit = objItemHeader.GetValue("PurPackMsr", 0)
                                objitem1.PurchaseQtyPerPackUnit = objItemHeader.GetValue("PurPackUn", 0)
                                objitem1.SalesPackagingUnit = objItemHeader.GetValue("SalPackMsr", 0)
                                objitem1.SalesQtyPerPackUnit = objItemHeader.GetValue("SalPackUn", 0)
                                Dim StrItmG As String

                                If objitem.ManageSerialNumbers = BoYesNoEnum.tNO And objitem.ManageBatchNumbers = BoYesNoEnum.tNO Then
                                    StrItmG = "Select ""Name"" from ""@MI_ITMSL"" where ""Code""<>''"
                                    objRsItmg.DoQuery(StrItmG)
                                    For Rec As Integer = 0 To objRsItmg.RecordCount - 1
                                        If objitem.ItemsGroupCode = objRsItmg.Fields.Item("Name").Value Then
                                            objitem1.ManageSerialNumbers = BoYesNoEnum.tYES
                                            objitem1.SRIAndBatchManageMethod = BoManageMethod.bomm_OnReleaseOnly
                                            objitem1.CostAccountingMethod = BoInventorySystem.bis_MovingAverage
                                            objitem.IssueMethod = BoIssueMethod.im_Manual
                                        End If
                                        objRsItmg.MoveNext()
                                    Next
                                Else
                                    If objItemHeader.GetValue("ManBtchNum", 0) = "Y" Then
                                        objitem1.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tYES
                                    Else
                                        objitem1.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tNO
                                    End If

                                    If objItemHeader.GetValue("ManSerNum", 0) = "Y" Then
                                        objitem1.ManageSerialNumbers = SAPbobsCOM.BoYesNoEnum.tYES
                                    Else
                                        objitem1.ManageSerialNumbers = SAPbobsCOM.BoYesNoEnum.tNO
                                    End If

                                    If objItemHeader.GetValue("IssueMthd", 0) = "M" Then
                                        objitem1.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Manual
                                    Else
                                        objitem1.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Backflush
                                    End If

                                    If objItemHeader.GetValue("MngMethod", 0) = "R" Then
                                        objitem1.SRIAndBatchManageMethod = SAPbobsCOM.BoManageMethod.bomm_OnReleaseOnly
                                    Else
                                        objitem1.SRIAndBatchManageMethod = SAPbobsCOM.BoManageMethod.bomm_OnEveryTransaction
                                    End If

                                    If objItemHeader.GetValue("EvalSystem", 0) = "F" Then
                                        objitem1.CostAccountingMethod = SAPbobsCOM.BoInventorySystem.bis_FIFO
                                    ElseIf objItemHeader.GetValue("EvalSystem", 0) = "A" Then
                                        objitem1.CostAccountingMethod = SAPbobsCOM.BoInventorySystem.bis_MovingAverage
                                    ElseIf objItemHeader.GetValue("EvalSystem", 0) = "S" Then
                                        objitem1.CostAccountingMethod = SAPbobsCOM.BoInventorySystem.bis_Standard
                                    Else
                                        objitem1.CostAccountingMethod = SAPbobsCOM.BoInventorySystem.bis_SNB
                                    End If
                                End If
                                'objItemPrice = objitem.PriceList
                                'For j As Integer = 0 To objItemPrice.Count - 1
                                '    objitem1.PriceList.SetCurrentLine(j)
                                '    objitem.PriceList.SetCurrentLine(j)
                                '    objitem1.PriceList.Price = objitem.PriceList.Price
                                '    objitem1.PriceList.Currency = objitem.PriceList.Currency
                                'Next j
                                RetVal = objitem1.Update()

                                If RetVal <> 0 Then
                                    Targetcompany.GetLastError(ErrCode, ErrMsg)
                                    objaddon.objapplication.SetStatusBarMessage(ErrCode & " " & ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Else
                                    objaddon.objapplication.SetStatusBarMessage("Item Updated to " & objRecordset.Fields.Item("U_DBName").Value, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                End If
                            End If
                            Try
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(Targetcompany)
                                Targetcompany = Nothing
                                GC.Collect()
                            Catch ex As Exception
                            End Try
                            objRecordset.MoveNext()
                        End While
                    End If
                Catch ex As Exception
                    objaddon.objapplication.SetStatusBarMessage(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    'MsgBox(ex.ToString)
                End Try
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

#Region "Fields"
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents CheckBox0 As SAPbouiCOM.CheckBox
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents ComboBox1 As SAPbouiCOM.ComboBox
        Private WithEvents ComboBox2 As SAPbouiCOM.ComboBox
        Private WithEvents ComboBox3 As SAPbouiCOM.ComboBox
        Private WithEvents ComboBox4 As SAPbouiCOM.ComboBox
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox5 As SAPbouiCOM.ComboBox
#End Region

        Private Sub CheckBox0_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles CheckBox0.PressedAfter
            'Try
            '    Dim objSyncForm As SAPbouiCOM.Form
            '    Dim objMatrix As SAPbouiCOM.Matrix
            '    Dim objStatic As SAPbouiCOM.StaticText
            '    'Dim objCheck As SAPbouiCOM.CheckBox
            '    If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
            '        If objform.Items.Item("5").Specific.String <> "" Then
            '            If CheckBox0.Checked = True Then
            '                Dim activeform As New FrmSyncConfig
            '                activeform.Show()
            '                objSyncForm = objaddon.objapplication.Forms.GetForm("MISYNC", 0)
            '                objSyncForm = objaddon.objapplication.Forms.ActiveForm
            '                objMatrix = objSyncForm.Items.Item("MtxConfig").Specific
            '                activeform.LoadMatrix(objSyncForm.UniqueID)
            '                objStatic = objSyncForm.Items.Item("lblDB").Specific
            '                objSyncForm.Items.Item("txtName").Specific.String = objform.Items.Item("5").Specific.String
            '                'objSyncForm.ActiveItem = "txtName"
            '                objStatic.Caption += " " & objaddon.objcompany.CompanyDB
            '                'objSyncForm.Items.Item("lblName").Visible = False
            '                'objSyncForm.Items.Item("txtCode").Visible = False
            '                CheckBox0.Item.Enabled = False
            '            Else
            '                CheckBox0.Checked = False
            '                Exit Sub
            '            End If
            '        Else
            '            CheckBox0.Checked = False
            '            objaddon.objapplication.SetStatusBarMessage("Please Create ItemCode...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            '            Exit Sub
            '        End If
            '    End If
            'Catch ex As Exception
            '    objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            'End Try
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    If CheckBox0.Checked = True Then
                        CheckBox0.Item.Enabled = False
                    Else
                        CheckBox0.Item.Enabled = True
                    End If
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    If objaddon.Validate_Master() Then
                        If CheckBox0.Checked = True Then
                            UpdateItemToDestinationDB(objform.Items.Item("5").Specific.String)
                        End If
                    End If
                End If

            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            Try
                Dim DBName, Str, Str1 As String
                Dim objRecordset, objrs, objrs1 As SAPbobsCOM.Recordset
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                Dim objcombo As SAPbouiCOM.ComboBox
                objcombo = objform.Items.Item("1320002059").Specific
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    'ItemCreate = objaddon.objglobalmethods.getSingleValue("select ""U_ItemSync"" from OADP")
                    DBName = objaddon.objglobalmethods.getSingleValue("select Top 1 ""U_DBName"" from ""@DB_LIST""")
                    If objaddon.ValidateItemSync() Then
                        objaddon.objapplication.SetStatusBarMessage(objform.Items.Item("5").Specific.String & " Item is Not Created or Updated in " & DBName & " Please Create in appropriate DB", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        BubbleEvent = False : Exit Sub
                    End If
                    If objcombo.Selected.Description.ToUpper = "Saudi".ToUpper Then
                        If ComboBox4.Selected.Description = "-" Then
                            objaddon.objapplication.SetStatusBarMessage("Please Select the Item Type Local/Import...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            BubbleEvent = False : Exit Sub
                        End If
                    End If
                End If
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    If objcombo.Selected.Description.ToUpper = "Saudi".ToUpper Then
                        Str = "select Distinct T1.""U_DBName"" from ""@DB_LIST"" T1  where  ifnull(""U_Select"",'')='Y'"
                        objRecordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objrs1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objRecordset.DoQuery(Str)
                        If objRecordset.RecordCount > 0 Then
                            While Not objRecordset.EoF
                                Str1 = "Select T0.""NextNumber""-1 as ""NextNumber1"", T1.""NextNumber"" as ""NextNumber2"",T1.""SeriesName"" as ""SeriesName"" from NNM1 T0 join """ & objRecordset.Fields.Item("U_DBName").Value.ToString & """.NNM1 T1 on T0.""ObjectCode""=T1.""ObjectCode"" and T0.""SeriesName""=T1.""SeriesName"""
                                Str1 += vbCrLf + "  where T1.""ObjectCode""='4' and T1.""SeriesName""='Saudi' "
                                objrs1.DoQuery(Str1)
                                If objrs1.RecordCount > 0 Then
                                    If CInt(objrs1.Fields.Item("NextNumber1").Value.ToString) > CInt(objrs1.Fields.Item("NextNumber2").Value.ToString) Then
                                        'Str1 = objaddon.objglobalmethods.getSingleValue("select Top 1 ""ItemCode"" from OITM where ""Series""='" & objcombo.Selected.Value & "' and ""ItemCode"" like '%" & objrs1.Fields.Item("NextNumber2").Value.ToString.Trim & "' Order by ""CreateDate"" desc")
                                        'If Str1 <> "" Then
                                        '    UpdateItemToDestinationDB(Str1)
                                        'End If
                                        objaddon.objapplication.SetStatusBarMessage("Please Update the " & objrs1.Fields.Item("NextNumber2").Value.ToString & " ending Item. Since Item Creation is not possible in the " & objRecordset.Fields.Item("U_DBName").Value.ToString & " DB...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        BubbleEvent = False : Exit Sub
                                    End If
                                End If
                                'End If
                                objRecordset.MoveNext()
                            End While
                        End If
                        objRecordset = Nothing
                        objrs = Nothing
                        objrs1 = Nothing
                    End If
                End If
                If SAPPassWord = "" Then
                    objaddon.objapplication.StatusBar.SetText("Please enter your SAP Password to sync the Item...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False
                    AlignLeft = objform.Left
                    AlignTop = objform.Height
                    Dim activeform As New FrmGetUserData
                    activeform.Show()
                End If
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try

        End Sub

        Private Sub Form_DataUpdateAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try

                If CheckBox0.Checked = True Then
                    CheckBox0.Item.Enabled = False
                Else
                    CheckBox0.Item.Enabled = True
                End If
            Catch ex As Exception

            End Try


        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                CheckBox0.Item.Left = objform.Items.Item("12").Left '456
                CheckBox0.Item.Top = objform.Items.Item("12").Top + 15 ' 54
                ComboBox4.Item.Left = objform.Items.Item("24").Left '150
                ComboBox4.Item.Top = objform.Items.Item("24").Top + 15 '113
                StaticText0.Item.Top = objform.Items.Item("25").Top + 15 '113
                StaticText0.Item.Left = objform.Items.Item("25").Left  ' 5
            Catch ex As Exception

            End Try

        End Sub

        Private Sub ComboBox3_ComboSelectAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ComboBox3.ComboSelectAfter
            Try
                Dim objcombo As SAPbouiCOM.ComboBox
                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    objcombo = objform.Items.Item("1320002059").Specific
                    If objcombo.Selected.Description = "Saudi" Then  '"Saudi"
                        If SAPPassWord = "" Then
                            AlignLeft = objform.Left
                            AlignTop = objform.Height
                            Dim activeform As New FrmGetUserData
                            activeform.Show()
                        End If
                        CheckBox0.Item.Enabled = True
                        CheckBox0.Checked = True
                        CheckBox0.Item.Enabled = False
                    Else
                        CheckBox0.Item.Enabled = True
                        CheckBox0.Checked = False
                        CheckBox0.Item.Enabled = False
                    End If
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub FrmItemMaster_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo) Handles Me.DataLoadAfter
            Try
                Dim objcombo As SAPbouiCOM.ComboBox
                objcombo = objform.Items.Item("1320002059").Specific
                If objcombo.Selected.Description = "Saudi" Then  '"Saudi"
                    CheckBox0.Item.Enabled = True
                    CheckBox0.Checked = True
                    CheckBox0.Item.Enabled = False
                Else
                    CheckBox0.Item.Enabled = True
                    CheckBox0.Checked = False
                    CheckBox0.Item.Enabled = False
                End If
                If CheckBox0.Checked = True Then
                    CheckBox0.Item.Enabled = False
                Else
                    CheckBox0.Item.Enabled = True
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Form_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    Try
                        If objaddon.Validate_Master() Then
                            If CheckBox0.Checked = True Then
                                If AddItemToDestinationDB(objform.Items.Item("5").Specific.String) Then
                                    objaddon.objapplication.StatusBar.SetText("Item Synced Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                Else
                                    objaddon.objapplication.StatusBar.SetText("Item Not Synced...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If

                            Else
                                Exit Sub
                            End If
                        End If

                    Catch ex As Exception
                        objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Finally
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    End Try
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_CloseAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                FormCount -= 1
            Catch ex As Exception
            End Try
        End Sub


    End Class
End Namespace
