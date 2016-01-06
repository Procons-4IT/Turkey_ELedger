Module MainMenu
    Public blnIsHANA As Boolean = False

#Region "... Main ..."
    Sub Main()
        Try
            oGFun.SetApplication() '1)

            'oApplication.SetFilter(New SAPbouiCOM.EventFilter) '2)
            If Not oGFun.CookieConnect() = 0 Then '3)
                oApplication.MessageBox("DI Api Connection Failed")
                End
            End If
            If Not oGFun.ConnectionContext() = 0 Then '4)
                oApplication.MessageBox("Failed to Connect Company")
                End
            End If

            If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                blnIsHANA = True
            End If

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Application Not Found", "Production  Module")
            System.Windows.Forms.Application.ExitThread()
        Finally
        End Try

        Try
            Try

                Dim oTableCreation As New TableCreation     '5)              
                'EventHandler.SetEventFilter()               '6)

                Dim orsOndalik As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                orsOndalik.DoQuery("SELECT ""DecSep"", ""DateSep""  FROM OADM")
                If orsOndalik.RecordCount > 0 Then
                    orsOndalik.MoveFirst()
                    strOndalikAyirac = orsOndalik.Fields.Item(0).Value
                    strTarihAyirac = orsOndalik.Fields.Item(1).Value
                End If


                oGFun.AddXML("Menu.xml")                    '7)
                'Dim oMeniItem As SAPbouiCOM.MenuItem = EventHandler.oApplication.Menus.Item("Purchase")
                'oMeniItem.Image = System.Windows.Forms.Application.StartupPath & "\Icon.jpg"


              

            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
                System.Windows.Forms.Application.ExitThread()
            Finally
            End Try
            oApplication.StatusBar.SetText("Connected.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            oApplication.StatusBar.SetText("EDefter Module Main Method Failed : ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub

    
#End Region
End Module
