Module EventHandler
#Region " ... Common Variables For SAP ..."

    Public WithEvents oApplication As SAPbouiCOM.Application

#End Region
#Region " ... 1) Menu Event ..."

    Private Sub oApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles oApplication.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID

                    Case EDefEntegratorSecim_FormUID
                        If oGFun.FormExist(EDefEntegratorSecim_FormUID) Then
                            oApplication.Forms.Item(EDefEntegratorSecim_FormUID).Visible = True
                            oApplication.Forms.Item(EDefEntegratorSecim_FormUID).Select()
                        Else
                            oEDefEntegratorSecim.LoadEDefEntegratorSecim()
                        End If
                    Case EDefListeRaporlar_FormUID
                        If oGFun.FormExist(EDefListeRaporlar_FormUID) Then
                            oApplication.Forms.Item(EDefListeRaporlar_FormUID).Visible = True
                            oApplication.Forms.Item(EDefListeRaporlar_FormUID).Select()
                        Else
                            oEDefListeRaporlar.LoadEDefListeRaporlar()
                        End If
                    Case EDefBelgeAyarlari_FormUID
                        If oGFun.FormExist(EDefBelgeAyarlari_FormUID) Then
                            oApplication.Forms.Item(EDefBelgeAyarlari_FormUID).Visible = True
                            oApplication.Forms.Item(EDefBelgeAyarlari_FormUID).Select()
                        Else
                            oEDefBelgeAyarlari.LoadEDefBelgeAyarlari()
                        End If
                        'EDefBelgeAyarlari_FormUID


                   
                End Select
                oForm = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "SenetHareketIptal", "CekHareketIptal", "1283", "1282", "1281", "1292", "1293", "1287", "519", "1284", "1286", "1288", "1291", "1289", "1290", "5890", "Delete Row"
                        '1288 bir sonraki kayit
                        '1291 son kayit
                        '1289 bir önceki kayıt
                        '1290 ilk kayıt
                        '1283 iptal
                        'sevda sağ klik
                        Select Case oForm.UniqueID
                            Case EDefEntegratorSecim_FormUID
                                oEDefEntegratorSecim.MenuEvent(pVal, BubbleEvent)
                            Case EDefBelgeAyarlari_FormUID
                                oEDefBelgeAyarlari.MenuEvent(pVal, BubbleEvent)

                        End Select

                        Select Case oForm.Type
                            
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("Menu Event Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
#End Region
#Region " ... 2) Item Event ..."
    Private Sub oApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles oApplication.ItemEvent
        Try

            Select Case pVal.FormUID
                Case EDefEntegratorSecim_FormUID
                    oEDefEntegratorSecim.ItemEvent(pVal.FormUID, pVal, BubbleEvent)
                Case EDefListeRaporlar_FormUID
                    oEDefListeRaporlar.ItemEvent(pVal.FormUID, pVal, BubbleEvent)
                Case EDefListeRaporlarDeger_FormUID
                    oEDefListeRaporlarDeger.ItemEvent(pVal.FormUID, pVal, BubbleEvent)
                Case EDefBelgeAyarlari_FormUID
                    oEDefBelgeAyarlari.ItemEvent(pVal.FormUID, pVal, BubbleEvent)
                Case EDefRapor_FormUID
                    oEDefRapor.ItemEvent(pVal.FormUID, pVal, BubbleEvent)


            End Select

            Select Case pVal.FormType


            End Select


        Catch ex As Exception
            oApplication.StatusBar.SetText("ItemEvent Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
#End Region
#Region " ... 6) Right Click Event ..."
    Private Sub oApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles oApplication.RightClickEvent
        Try
            'Delete the User Creation Menus from Main Menu ..
            If oApplication.Menus.Item("1280").SubMenus.Exists("AddEntry") = True Then oApplication.Menus.Item("1280").SubMenus.RemoveEx("AddEntry")
            If oApplication.Menus.Item("1280").SubMenus.Exists("UpdateEntry") = True Then oApplication.Menus.Item("1280").SubMenus.RemoveEx("UpdateEntry")
            '   If oApplication.Menus.Item("1280").SubMenus.Exists("Batch Creation") = True Then oApplication.Menus.Item("1280").SubMenus.RemoveEx("Batch Creation")
            If oApplication.Menus.Item("1280").SubMenus.Exists("DeleteEntry") = True Then oApplication.Menus.Item("1280").SubMenus.RemoveEx("DeleteEntry")
            Select Case eventInfo.FormUID
                

                    'Case ReceiptFormID
                    '    oReceiptMgmt.RightClickEvent(eventInfo, BubbleEvent)
                    'Case ReturnFormID
                    '    oReturnMgmt.RightClickEvent(eventInfo, BubbleEvent)
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(addonName & " : Right Click Event Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    'Private Sub oApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles oApplication.RightClickEvent
    '    If eventInfo.FormUID = "CME" And eventInfo.BeforeAction = True And eventInfo.ItemUID = "omatHar" Then
    '        If oApplication.Forms.Item(eventInfo.FormUID).Items.Item("omatHar").Enabled = True And eventInfo.ItemUID = "omatHar" Then
    '            If eventInfo.Row > 4 Then

    '            Else
    '                'oApplication.Forms.Item("CME").Items.Item("omatHar").Specific.Columns.Item("1").Cells.Item(oApplication.Forms.Item("CME").Items.Item("omatHar").Specific.RowCount).Specific.String = _
    '                'oApplication.Forms.Item("CME").Items.Item("omatHar").Specific.Columns.Item("omatHar").Cells.Item(oApplication.Forms.Item("CME").Items.Item("omatHar").Specific.RowCount).Specific.String(+"")
    '                BubbleEvent = False
    '            End If
    '        End If

    '        If oApplication.Forms.Item(eventInfo.FormUID).Items.Item("MyMatrix").Enabled = True And eventInfo.ItemUID = "MyMatrix" Then
    '            If eventInfo.Row > 4 Then

    '            Else
    '                oApplication.Forms.Item("MyForm").Items.Item("MyMatrix").Specific.Columns.Item("MyColumn-1").Cells.Item(oApplication.Forms.Item("MyForm").Items.Item("MyMatrix").Specific.RowCount).Specific.String = _
    '                oApplication.Forms.Item("MyForm").Items.Item("MyMatrix").Specific.Columns.Item("MyMatrix").Cells.Item(oApplication.Forms.Item("MyForm").Items.Item("MyMatrix").Specific.RowCount).Specific.String(+"")
    '                BubbleEvent = False
    '            End If
    '        End If
    '    End If
    'End Sub
#End Region
#Region " ... 3) FormDataEvent ..."
    Private Sub oApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles oApplication.FormDataEvent
        Try
            Select Case BusinessObjectInfo.FormUID
                'Case EDefEntegratorSecim_FormUID
                '    oEDefEntegratorSecim.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case EDefListeRaporlarDeger_FormUID
                    oEDefListeRaporlarDeger.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case EDefRapor_FormUID
                    oEDefRapor.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    'EDefRapor_FormUID
            End Select

            Select Case BusinessObjectInfo.FormTypeEx
                

            End Select

        Catch ex As Exception
            oApplication.StatusBar.SetText("Project FormDataEvent Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
#End Region

#Region " ... 3) AppEvent ..."
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles oApplication.AppEvent

        Select Case EventType

            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                Dim myProcesses2 As Process() = Process.GetProcessesByName("EDefter")
                Dim myProcess2 As Process
                For Each myProcess2 In myProcesses2
                    If myProcess2.MainWindowTitle.Contains("EDefter") Then myProcess2.Kill()
                Next myProcess2

                Dim myProcesses As Process() = System.Diagnostics.Process.GetProcesses
                Dim myProcess As Process
                For Each myProcess In myProcesses
                    If myProcess.ProcessName.Contains("EDefter") Or myProcess.MainWindowTitle.Contains("EDefter") Then
                        myProcess.Kill()
                    End If
                Next myProcess

                End


            Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition

                Dim myProcesses2 As Process() = Process.GetProcessesByName("EDefter")
                Dim myProcess2 As Process
                For Each myProcess2 In myProcesses2
                    If myProcess2.MainWindowTitle.Contains("EDefter") Then myProcess2.Kill()
                Next myProcess2

                Dim myProcesses As Process() = System.Diagnostics.Process.GetProcesses
                Dim myProcess As Process
                For Each myProcess In myProcesses
                    If myProcess.ProcessName.Contains("EDefter") Or myProcess.MainWindowTitle.Contains("EDefter") Then
                        myProcess.Kill()
                    End If
                Next myProcess
                End


            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown

                Dim myProcesses2 As Process() = Process.GetProcessesByName("EDefter")
                Dim myProcess2 As Process
                For Each myProcess2 In myProcesses2
                    If myProcess2.MainWindowTitle.Contains("EDefter") Then myProcess2.Kill()
                Next myProcess2

                Dim myProcesses As Process() = System.Diagnostics.Process.GetProcesses
                Dim myProcess As Process
                For Each myProcess In myProcesses
                    If myProcess.ProcessName.Contains("EDefter") Or myProcess.MainWindowTitle.Contains("EDefter") Then
                        myProcess.Kill()
                    End If
                Next myProcess
                End


        End Select

    End Sub
#End Region
  
End Module
