Public Class ClsEDefListeRaporlarDeger
    Dim frmEDefListeRaporlarDeger As SAPbouiCOM.Form
    Dim oDBDSHeader, oDBDSDetails As SAPbouiCOM.DBDataSource
    Dim boolFormLoaded As Boolean = False
   
    Sub LoadEDefListeRaporlarDeger(Optional ByVal Docentry As Integer = 0)
        Try
            oGFun.LoadXML(frmEDefListeRaporlarDeger, EDefListeRaporlarDeger_FormUID, EDefListeRaporlarDegerXml)
            frmEDefListeRaporlarDeger = oApplication.Forms.Item(EDefListeRaporlarDeger_FormUID)
            oDBDSHeader = frmEDefListeRaporlarDeger.DataSources.DBDataSources.Item(0)
            oDBDSDetails = frmEDefListeRaporlarDeger.DataSources.DBDataSources.Item(1)

            frmEDefListeRaporlarDeger.EnableMenu("772", False)
            frmEDefListeRaporlarDeger.EnableMenu("773", False)
            frmEDefListeRaporlarDeger.EnableMenu("774", False)
            frmEDefListeRaporlarDeger.EnableMenu("1280", False)
            frmEDefListeRaporlarDeger.EnableMenu("1283", False)
            frmEDefListeRaporlarDeger.EnableMenu("1285", False)
            frmEDefListeRaporlarDeger.EnableMenu("1284", False)
            frmEDefListeRaporlarDeger.EnableMenu("1286", False)

            '**2 Eğer değişiklik yapılacak bir hareket satırı seçili ise açık gelsin diye...
            If Docentry > 0 Then
                frmEDefListeRaporlarDeger.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                frmEDefListeRaporlarDeger.Items.Item("tFolderNo").Specific.Value = Docentry
                frmEDefListeRaporlarDeger.Items.Item("1").Click()
                frmEDefListeRaporlarDeger.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                frmEDefListeRaporlarDeger.Items.Item("tFolderNo").Enabled = False
                Exit Sub
            Else
                frmEDefListeRaporlarDeger.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE 'Form modu sevda
                frmEDefListeRaporlarDeger.Items.Item("tFolderNo").Enabled = False
            End If
            '**2 SON

            Me.InitForm()
            Me.DefineModesForFields()
        Catch ex As Exception
            oGFun.StatusBarErrorMsg(" Load report parameters entry Form Failed : " & ex.Message)
            boolFormLoaded = False
        End Try
    End Sub
    Sub InitForm()
        Try
            frmEDefListeRaporlarDeger.Freeze(True)
            oGFun.LoadComboBoxSeries(frmEDefListeRaporlarDeger.Items.Item("15").Specific, "ELRAP")
            

            frmEDefListeRaporlarDeger.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("InitForm Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            frmEDefListeRaporlarDeger.Freeze(False)
        Finally
            frmEDefListeRaporlarDeger.Freeze(False)
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Try
                        Select Case pVal.ItemUID
                            Case "1"
                                If pVal.BeforeAction And (frmEDefListeRaporlarDeger.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or frmEDefListeRaporlarDeger.Mode = SAPbouiCOM.BoFormMode.fm_EDIT_MODE) Then
                                    If Me.ValidateALL = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        oDBDSDetails.SetValue("U_Guncellendi", oDBDSDetails.Offset, "Y")
                                        BubbleEvent = True
                                    End If
                                ElseIf pVal.BeforeAction And frmEDefListeRaporlarDeger.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    If Me.ValidateALL = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        oDBDSDetails.SetValue("U_Guncellendi", oDBDSDetails.Offset, "Y")
                                        BubbleEvent = True
                                    End If
                                End If
                                If (pVal.ActionSuccess = True And frmEDefListeRaporlarDeger.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                                    frmEDefListeRaporlarDeger.Items.Item("2").Click()
                                    'InitForm()
                                End If

                        End Select
                    Catch ex As Exception
                        oGFun.StatusBarErrorMsg("Item Pressed Event Failed 5: " & ex.Message)
                    Finally
                    End Try

                
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    Select Case pVal.ItemUID
                        Case "15"
                            If pVal.ItemChanged And pVal.BeforeAction = False Then
                                oDBDSHeader.SetValue("DocNum", 0, frmEDefListeRaporlarDeger.BusinessObject.GetNextSerialNumber(frmEDefListeRaporlarDeger.Items.Item("15").Specific.Selected.Value))
                            End If
                    End Select

              
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("ItemEvent Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub DefineModesForFields()
        Try
            frmEDefListeRaporlarDeger.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmEDefListeRaporlarDeger.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("DefineModesForFields Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "1282"
                    Me.InitForm()
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Menu Event Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Function ValidateALL() As Boolean
        Try
            Dim Baslik As String = oDBDSHeader.GetValue("U_RBASLIK", 0).Trim
            

            If Baslik = String.Empty Or Baslik = "" Then
                oGFun.StatusBarErrorMsg("Seçim Ölçüt Adı Girmelisiniz...")
                Return False
            End If
            
            Return True
        Catch ex As Exception
            oGFun.StatusBarErrorMsg("Validate Event Failed" & ex.Message)
            Return False
        End Try
    End Function

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oEDefListeRaporlar.LoadGrid()
        Catch ex As Exception

        End Try
    End Sub
End Class
