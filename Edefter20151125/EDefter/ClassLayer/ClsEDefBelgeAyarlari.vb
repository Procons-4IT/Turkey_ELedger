Public Class ClsEDefBelgeAyarlari
    Dim frmEDefBelgeAyarlari As SAPbouiCOM.Form
    Dim oDBDSHeader, oDBDSDetails As SAPbouiCOM.DBDataSource
    Dim oMatrix As SAPbouiCOM.Matrix
    Dim oIptalButon As SAPbouiCOM.Button
    Dim boolFormLoaded As Boolean = False
    
    Sub LoadEDefBelgeAyarlari(Optional ByVal Docentry As Integer = 0)
        Try
            oGFun.LoadXML(frmEDefBelgeAyarlari, EDefBelgeAyarlari_FormUID, EDefBelgeAyarlariXml)
            frmEDefBelgeAyarlari = oApplication.Forms.Item(EDefBelgeAyarlari_FormUID)
            oDBDSHeader = frmEDefBelgeAyarlari.DataSources.DBDataSources.Item(0)
            oDBDSDetails = frmEDefBelgeAyarlari.DataSources.DBDataSources.Item(1)
            oMatrix = frmEDefBelgeAyarlari.Items.Item("13").Specific
            oIptalButon = frmEDefBelgeAyarlari.Items.Item("2").Specific

            frmEDefBelgeAyarlari.EnableMenu("772", False)
            frmEDefBelgeAyarlari.EnableMenu("773", False)
            frmEDefBelgeAyarlari.EnableMenu("774", False)
            frmEDefBelgeAyarlari.EnableMenu("1280", False)
            frmEDefBelgeAyarlari.EnableMenu("1283", False)
            frmEDefBelgeAyarlari.EnableMenu("1285", False)
            frmEDefBelgeAyarlari.EnableMenu("1284", False)
            frmEDefBelgeAyarlari.EnableMenu("1286", False)

            Dim rsetilk As SAPbobsCOM.Recordset
            Dim sqlquery As String = ""
            
            If blnIsHANA Then
                sqlquery = "SELECT 1 FROM ""@EINTEGSEL"""
                rsetilk = oGFun.ReturnRecordSet(sqlquery)
                If rsetilk.RecordCount > 0 Then
                    sqlquery = "SELECT ""DocNum"" as ""Sonuc"" FROM ""@EINTEGSEL"" "
                    rsetilk = oGFun.ReturnRecordSet(sqlquery)
                Else
                    sqlquery = "SELECT 0 as ""Sonuc"" FROM DUMMY"
                    rsetilk = oGFun.ReturnRecordSet(sqlquery)

                End If
            Else
                sqlquery = "IF EXISTS(SELECT 1 FROM [dbo].[@EINTEGSEL](NOLOCK)) " + vbCrLf
                sqlquery = sqlquery + "SELECT DocNum as Sonuc FROM [dbo].[@EINTEGSEL](NOLOCK) " + vbCrLf
                sqlquery = sqlquery + "ELSE " + vbCrLf
                sqlquery = sqlquery + "SELECT 0 as Sonuc" + vbCrLf
                rsetilk = oGFun.ReturnRecordSet(sqlquery)
            End If

            'MessageBox.Show(rsetilk.Fields.Item("Sonuc").Value)

            If rsetilk.Fields.Item("Sonuc").Value <> 0 Then
                frmEDefBelgeAyarlari.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                frmEDefBelgeAyarlari.Items.Item("tFolderNo").Specific.Value = rsetilk.Fields.Item("Sonuc").Value
                frmEDefBelgeAyarlari.Items.Item("1").Click()
                frmEDefBelgeAyarlari.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                frmEDefBelgeAyarlari.Items.Item("tFolderNo").Enabled = False
                Exit Sub
            Else
                MessageBox.Show("Belge ayarları için öncelikle entegratör seçimi yapmalısınız!..", "Uyarı", MessageBoxButtons.OK)
                oIptalButon.Item.Click()
                'frmEDefBelgeAyarlari.Close()
                'frmEDefBelgeAyarlari.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE 'Form modu sevda
                'frmEDefBelgeAyarlari.Items.Item("tFolderNo").Enabled = False
            End If

            'Me.InitForm()
            'Me.DefineModesForFields()
        Catch ex As Exception
            oGFun.StatusBarErrorMsg(" Load EDefBelgeAyarlari Form Failed : " & ex.Message)
            boolFormLoaded = False
        End Try
    End Sub
    Sub InitForm()
        Try
            frmEDefBelgeAyarlari.Freeze(True)
            oGFun.LoadComboBoxSeries(frmEDefBelgeAyarlari.Items.Item("15").Specific, "EInteg")
            oGFun.SetNewLine(oMatrix, oDBDSDetails)



            frmEDefBelgeAyarlari.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("InitForm Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            frmEDefBelgeAyarlari.Freeze(False)
        Finally
            frmEDefBelgeAyarlari.Freeze(False)
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Try
                        Select Case pVal.ItemUID
                            Case "1"
                                If pVal.BeforeAction And (frmEDefBelgeAyarlari.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or frmEDefBelgeAyarlari.Mode = SAPbouiCOM.BoFormMode.fm_EDIT_MODE Or frmEDefBelgeAyarlari.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If Me.ValidateALL = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        BubbleEvent = True
                                    End If
                                End If
                                If (pVal.ActionSuccess = True And frmEDefBelgeAyarlari.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                                    frmEDefBelgeAyarlari.Items.Item("2").Click()
                                    'InitForm()
                                End If

                        End Select
                    Catch ex As Exception
                        oGFun.StatusBarErrorMsg("Item Pressed Event Failed 3: " & ex.Message)
                    Finally
                    End Try

                  
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    Select Case pVal.ItemUID
                        Case "15"
                            If pVal.ItemChanged And pVal.BeforeAction = False Then
                                oDBDSHeader.SetValue("DocNum", 0, frmEDefBelgeAyarlari.BusinessObject.GetNextSerialNumber(frmEDefBelgeAyarlari.Items.Item("15").Specific.Selected.Value))
                            End If
                    End Select

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Try
                        Select Case pVal.ItemUID
                            Case "13"
                                Select Case pVal.ColUID
                                    Case "V_1"
                                        oGFun.SetNewLine(oMatrix, oDBDSDetails, oMatrix.VisualRowCount, "V_1")

                                End Select
                        End Select
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Lost Focus Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                    End Try
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("ItemEvent Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    'Sub DefineModesForFields()
    '    Try
    '        frmEDefBelgeAyarlari.Items.Item("tFolderNo").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
    '        frmEDefBelgeAyarlari.Items.Item("tFolderNo").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
    '    Catch ex As Exception
    '        oApplication.StatusBar.SetText("DefineModesForFields Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '    Finally
    '    End Try
    'End Sub

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
            'Dim Baslik As String = oDBDSHeader.GetValue("U_PRMBASLIK", 0).Trim
            'Dim ParametreDegeri As String = oDBDSDetails.GetValue("U_PRMADI", 0).Trim
            'Dim ParametreAciklamasi As String = oDBDSDetails.GetValue("U_PRMDEGER", 0).Trim

            'If Baslik = String.Empty Or Baslik = "" Then
            '    oGFun.StatusBarErrorMsg("Grup Adı Girmelisiniz...")
            '    Return False
            'End If
            'If ParametreDegeri = String.Empty Or ParametreDegeri = "" Then
            '    oGFun.StatusBarErrorMsg("Parametre Değeri Girmelisiniz...")
            '    Return False
            'End If
            'If ParametreAciklamasi = String.Empty Or ParametreAciklamasi = "" Then
            '    oGFun.StatusBarErrorMsg("Parametre Açıklaması Girmelisiniz...")
            '    Return False
            'End If
            Return True
        Catch ex As Exception
            oGFun.StatusBarErrorMsg("Validate Event Failed" & ex.Message)
            Return False
        End Try
    End Function

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            'oEDefBelgeAyarlari.LoadGrid()
        Catch ex As Exception

        End Try
    End Sub
End Class
