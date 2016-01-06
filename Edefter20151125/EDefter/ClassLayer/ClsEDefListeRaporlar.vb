Public Class ClsEDefListeRaporlar
    Dim frmEDefListeRaporlar As SAPbouiCOM.Form
    Dim boolFormLoaded As Boolean = False
    Dim oGrid As SAPbouiCOM.Grid
    Dim strSecili As String = ""

    Sub LoadEDefListeRaporlar()
        Try
            oGFun.LoadXML(frmEDefListeRaporlar, EDefListeRaporlar_FormUID, EDefListeRaporlarXml)
            frmEDefListeRaporlar = oApplication.Forms.Item(EDefListeRaporlar_FormUID)

            frmEDefListeRaporlar.EnableMenu("772", False)
            frmEDefListeRaporlar.EnableMenu("773", False)
            frmEDefListeRaporlar.EnableMenu("774", False)
            frmEDefListeRaporlar.EnableMenu("1280", False)
            frmEDefListeRaporlar.EnableMenu("1283", False)
            frmEDefListeRaporlar.EnableMenu("1285", False)
            frmEDefListeRaporlar.EnableMenu("1284", False)
            frmEDefListeRaporlar.EnableMenu("1286", False)

            oGrid = frmEDefListeRaporlar.Items.Item("7").Specific

            InitForm()

        Catch ex As Exception
            oGFun.StatusBarErrorMsg(" Load LoadEDefListeRaporlar form failed : " & ex.Message)
            boolFormLoaded = False
        End Try
    End Sub
    Sub InitForm()
        Try
            frmEDefListeRaporlar.Freeze(True)

            LoadGrid()


            frmEDefListeRaporlar.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("InitForm Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            frmEDefListeRaporlar.Freeze(False)
        Finally
            frmEDefListeRaporlar.Freeze(False)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "1282"
                    Me.InitForm()
                    Me.DefineModesForFields()
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Menu Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Try
                        Select Case pVal.ItemUID
                            Case "7" 'Grid
                                If pVal.ActionSuccess Then
                                    'Grid seçili satır bulma
                                    If DirectCast(pVal, SAPbouiCOM.ItemEventClass).ColUID = "RowsHeader" Then 'GridSort event değilse.. yani girdde kolon değil satır seçildiyse
                                        Dim intDataTableRowIndex As Integer

                                        intDataTableRowIndex = oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder).ToString())
                                        strSecili = oGrid.DataTable.GetValue(0, intDataTableRowIndex).ToString()

                                    End If
                                End If

                        End Select
                    Catch ex As Exception

                    End Try
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    Select Case pVal.ItemUID
                        Case "3"
                            If pVal.ActionSuccess Then
                                oEDefListeRaporlarDeger.LoadEDefListeRaporlarDeger()
                            End If
                        Case "6" 'Düzelt tuşu
                            If pVal.BeforeAction Then
                                If strSecili = "" Then
                                    oApplication.MessageBox("Güncellenecek Kaydı Seçiniz...", 1)
                                    BubbleEvent = False
                                End If
                            End If
                            If pVal.ActionSuccess Then
                                oEDefListeRaporlarDeger.LoadEDefListeRaporlarDeger(strSecili)
                                strSecili = ""
                            End If
                        Case "9" 'Rapor Çağır
                            If pVal.BeforeAction Then
                                If strSecili = "" Then
                                    oApplication.MessageBox("Verisi İstenen Kaydı Seçiniz...", 1)
                                    BubbleEvent = False
                                End If
                            End If
                            If pVal.ActionSuccess Then
                                oEDefRapor.LoadEDefRapor(strSecili)
                                strSecili = ""
                            End If
                    End Select


            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Item  Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub DefineModesForFields()
        Try

        Catch ex As Exception
            oApplication.StatusBar.SetText("DefineModesForFields Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub LoadGrid()
        Dim oDataTable As SAPbouiCOM.DataTable = frmEDefListeRaporlar.DataSources.DataTables.Item("DT_0")
        Dim oQuery As String
        If blnIsHANA Then
            oQuery = "SELECT ""DocEntry"" as ""No"", ""U_RBASLIK"" as ""Başlık"" FROM ""@ELRAPH"""
        Else
            oQuery = "SELECT [DocEntry] as No, [U_RBASLIK] as Başlık FROM [dbo].[@ELRAPH](NOLOCK)"
        End If

        oDataTable.Clear()
        oDataTable.ExecuteQuery(oQuery)

    End Sub

End Class

