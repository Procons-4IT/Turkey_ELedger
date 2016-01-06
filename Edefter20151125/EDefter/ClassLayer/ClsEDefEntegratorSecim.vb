Public Class ClsEDefEntegratorSecim
    Dim frmEDefEntegratorSecim As SAPbouiCOM.Form
    Dim oDBDSHeader As SAPbouiCOM.DBDataSource


    Sub LoadEDefEntegratorSecim()
        Try
            oGFun.LoadXML(frmEDefEntegratorSecim, EDefEntegratorSecim_FormUID, EDefEntegratorSecimXml)
            frmEDefEntegratorSecim = oApplication.Forms.Item(EDefEntegratorSecim_FormUID)

            oDBDSHeader = frmEDefEntegratorSecim.DataSources.DBDataSources.Item(0)


            frmEDefEntegratorSecim.EnableMenu("1283", False)
            frmEDefEntegratorSecim.EnableMenu("1284", False)
            frmEDefEntegratorSecim.EnableMenu("1285", False)
            frmEDefEntegratorSecim.EnableMenu("1286", False)
            frmEDefEntegratorSecim.EnableMenu("1287", False)
            '&Remove 1283
            '&Cancel 1284
            'R&estore 1285
            'Cl&ose 1286
            'Duplicate 1287

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
                frmEDefEntegratorSecim.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                frmEDefEntegratorSecim.Items.Item("tFolderNo").Specific.Value = rsetilk.Fields.Item("Sonuc").Value
                frmEDefEntegratorSecim.Items.Item("1").Click()
                frmEDefEntegratorSecim.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                frmEDefEntegratorSecim.Items.Item("tFolderNo").Enabled = False
                Exit Sub
            Else
                frmEDefEntegratorSecim.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE 'Form modu sevda
                frmEDefEntegratorSecim.Items.Item("tFolderNo").Enabled = False
            End If
            '**2 SON


            Me.InitForm()
            Me.DefineModesForFields()
        Catch ex As Exception
            oGFun.StatusBarErrorMsg(" Load Entegrasyon Secim Form Failed : " & ex.Message)

        End Try
    End Sub
    Sub InitForm()
        Try
            frmEDefEntegratorSecim.Freeze(True)

            oGFun.LoadComboBoxSeries(frmEDefEntegratorSecim.Items.Item("c_series").Specific, "EInteg")

            frmEDefEntegratorSecim.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("InitForm Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            frmEDefEntegratorSecim.Freeze(False)
        Finally
            frmEDefEntegratorSecim.Freeze(False)
        End Try
    End Sub
    Sub DefineModesForFields()
        Try

            frmEDefEntegratorSecim.Items.Item("tFolderNo").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmEDefEntegratorSecim.Items.Item("tFolderNo").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("DefineModesForFields Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
               
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Try
                        Select Case pVal.ItemUID
                            Case "1"
                                If pVal.BeforeAction Then
                                    If pVal.BeforeAction And (frmEDefEntegratorSecim.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or frmEDefEntegratorSecim.Mode = SAPbouiCOM.BoFormMode.fm_EDIT_MODE Or frmEDefEntegratorSecim.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                        If Me.ValidateALL = False Then
                                            BubbleEvent = False
                                            Exit Sub
                                        Else
                                            BubbleEvent = True
                                        End If
                                   
                                    End If
                                End If
                                If pVal.ActionSuccess And frmEDefEntegratorSecim.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    frmEDefEntegratorSecim.Items.Item("2").Click()
                                End If

                        End Select
                    Catch ex As Exception
                        oGFun.StatusBarErrorMsg("Item Pressed Event Failed 6: " & ex.Message)
                    Finally
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    Try

                        Select Case pVal.ItemUID
                            Case "c_series"

                                If pVal.BeforeAction = False And pVal.ItemChanged Then
                                    Dim lNum As Integer = frmEDefEntegratorSecim.BusinessObject.GetNextSerialNumber(frmEDefEntegratorSecim.Items.Item(pVal.ItemUID).Specific.Selected.Value)

                                    oDBDSHeader.SetValue("DocNum", 0, lNum)
                                End If

                            Case "12"
                                If pVal.BeforeAction = False And pVal.ItemChanged Then
                                    Dim o As SAPbouiCOM.ComboBox
                                    Dim o2 As SAPbouiCOM.EditText
                                    o = frmEDefEntegratorSecim.Items.Item("12").Specific
                                    o2 = frmEDefEntegratorSecim.Items.Item("14").Specific
                                    o2.Value = DirectCast(o.Selected, SAPbouiCOM.ValidValueClass).Description
                                    
                                End If

                        End Select
                    Catch ex As Exception
                        oGFun.StatusBarErrorMsg(" Combo Select Event Failed : " & ex.Message)
                    Finally
                    End Try

            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("ItemEvent Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
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
            Dim cmbEnteg As SAPbouiCOM.ComboBox
            cmbEnteg = frmEDefEntegratorSecim.Items.Item("12").Specific
            'op2 = frmEDefEntegratorSecim.Items.Item("op2").Specific
            'op3 = frmEDefEntegratorSecim.Items.Item("op3").Specific
            'op4 = frmEDefEntegratorSecim.Items.Item("op4").Specific

            If cmbEnteg.Selected Is Nothing Then
                oGFun.StatusBarErrorMsg("Entegratörlerden birini seçip kaydetmeden ekrandan çıkış yapamazsınız...")
                Return False
            End If

            Return True
        Catch ex As Exception
            oGFun.StatusBarErrorMsg("Validate Event Failed" & ex.Message)
            Return False
        End Try
    End Function

   
End Class


