Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms

Public Class ClsEDefRapor
    Dim frmEDefRapor As SAPbouiCOM.Form
    Dim oDBDSHeader, oDBDSDetails, oDBDSDetails2, oDBDSDetails3, oDBDSDetails4, oDBDSDetails5 As SAPbouiCOM.DBDataSource
    Dim boolFormLoaded As Boolean = False
    Dim oMatrixUyumsoft As SAPbouiCOM.Matrix
    Dim oMatrixHuber As SAPbouiCOM.Matrix
    Dim oMatrixIzibiz As SAPbouiCOM.Matrix
    Dim oMatrixBimsa As SAPbouiCOM.Matrix

    Dim oDataTable As SAPbouiCOM.DataTable
    Dim oDataTable2 As SAPbouiCOM.DataTable
    Dim oDataTable3 As SAPbouiCOM.DataTable 'Huber
    Dim oDataTable4 As SAPbouiCOM.DataTable 'izibiz
    Dim oDataTable5 As SAPbouiCOM.DataTable 'Bimsa
    Dim oDataTableBaslik As SAPbouiCOM.DataTable
    'Dim oProgBar As SAPbouiCOM.ProgressBar
    Dim strEntegrator As String = ""
    Dim strEntegratorKod As String = ""
    Dim oedtYmyNo As SAPbouiCOM.EditText
    Dim oedtSatirNo As SAPbouiCOM.EditText
    Dim strDoc As Integer
    Dim strSelectedFilepath As String = ""

    Sub LoadEDefRapor(Optional ByVal Docentry As Integer = 0)
        Try
            oGFun.LoadXML(frmEDefRapor, EDefRapor_FormUID, EDefRaporXml)
            frmEDefRapor = oApplication.Forms.Item(EDefRapor_FormUID)
            oDBDSHeader = frmEDefRapor.DataSources.DBDataSources.Item("@ELRAPH")
            oDBDSDetails = frmEDefRapor.DataSources.DBDataSources.Item("@ELRAPS")
            oDBDSDetails2 = frmEDefRapor.DataSources.DBDataSources.Item("@ELRAPV")
            oDBDSDetails3 = frmEDefRapor.DataSources.DBDataSources.Item("@ELRAPVH")
            oDBDSDetails4 = frmEDefRapor.DataSources.DBDataSources.Item("@ELRAPVI")
            oDBDSDetails5 = frmEDefRapor.DataSources.DBDataSources.Item("@ELRAPVB")

            strDoc = Docentry

            Dim orsEntegrator As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If blnIsHANA Then
                orsEntegrator.DoQuery("SELECT ""U_IntegratorCode"", ""U_IntegratorName""  FROM ""@EINTEGSEL""")
            Else
                orsEntegrator.DoQuery("SELECT [U_IntegratorCode], [U_IntegratorName]  FROM [dbo].[@EINTEGSEL](NOLOCK)")
            End If
            If orsEntegrator.RecordCount > 0 Then
                orsEntegrator.MoveFirst()
                strEntegratorKod = orsEntegrator.Fields.Item(0).Value
                strEntegrator = orsEntegrator.Fields.Item(1).Value
            Else
                oGFun.StatusBarErrorMsg(" Entegratör seçimini yapmalısınız... ")
                boolFormLoaded = False
            End If

            If strEntegratorKod = "1" Then 'uyumsoft
                'Uyumsoft
                frmEDefRapor.Items.Item("13").Visible = True
                'huber
                frmEDefRapor.Items.Item("14").Visible = False
                'İzibiz
                frmEDefRapor.Items.Item("25").Visible = False
                'Bimsa
                frmEDefRapor.Items.Item("26").Visible = False

                frmEDefRapor.Items.Item("edtYmyNo").Visible = False
                frmEDefRapor.Items.Item("lblYmyNo").Visible = False

                frmEDefRapor.Items.Item("edtSatirNo").Visible = False
                frmEDefRapor.Items.Item("lblSatirNo").Visible = False

                'frmEDefRapor.Items.Item("33").Visible = False

                oMatrixUyumsoft = frmEDefRapor.Items.Item("13").Specific
            ElseIf strEntegratorKod = "2" Then 'İzibiz
                'Uyumsoft
                frmEDefRapor.Items.Item("13").Visible = False
                'İzibiz
                frmEDefRapor.Items.Item("14").Visible = False
                'Huber
                frmEDefRapor.Items.Item("25").Visible = True
                'Bimsa
                frmEDefRapor.Items.Item("26").Visible = False

                frmEDefRapor.Items.Item("edtYmyNo").Visible = True
                frmEDefRapor.Items.Item("lblYmyNo").Visible = True

                frmEDefRapor.Items.Item("edtSatirNo").Visible = True
                frmEDefRapor.Items.Item("lblSatirNo").Visible = True

                'frmEDefRapor.Items.Item("33").Visible = True

                oMatrixIzibiz = frmEDefRapor.Items.Item("25").Specific
                oedtYmyNo = frmEDefRapor.Items.Item("edtYmyNo").Specific
                oedtSatirNo = frmEDefRapor.Items.Item("edtSatirNo").Specific

            ElseIf strEntegratorKod = "4" Then '4 Huber - izibiz2
                'Uyumsoft
                frmEDefRapor.Items.Item("13").Visible = False
                'İzibiz
                frmEDefRapor.Items.Item("25").Visible = False
                'Huber
                frmEDefRapor.Items.Item("14").Visible = True
                'Bimsa
                frmEDefRapor.Items.Item("26").Visible = False

                frmEDefRapor.Items.Item("edtYmyNo").Visible = True
                frmEDefRapor.Items.Item("lblYmyNo").Visible = True

                frmEDefRapor.Items.Item("edtSatirNo").Visible = True
                frmEDefRapor.Items.Item("lblSatirNo").Visible = True

                'frmEDefRapor.Items.Item("33").Visible = True

                oMatrixHuber = frmEDefRapor.Items.Item("14").Specific
                oedtYmyNo = frmEDefRapor.Items.Item("edtYmyNo").Specific
                oedtSatirNo = frmEDefRapor.Items.Item("edtSatirNo").Specific
            Else '3 Bimsa
                'Uyumsoft
                frmEDefRapor.Items.Item("13").Visible = False
                'İzibiz
                frmEDefRapor.Items.Item("25").Visible = False
                'Huber
                frmEDefRapor.Items.Item("14").Visible = False
                'Bimsa
                frmEDefRapor.Items.Item("26").Visible = True

                frmEDefRapor.Items.Item("edtYmyNo").Visible = True
                frmEDefRapor.Items.Item("lblYmyNo").Visible = True

                frmEDefRapor.Items.Item("edtSatirNo").Visible = True
                frmEDefRapor.Items.Item("lblSatirNo").Visible = True

                'frmEDefRapor.Items.Item("33").Visible = True

                oMatrixBimsa = frmEDefRapor.Items.Item("26").Specific
                oedtYmyNo = frmEDefRapor.Items.Item("edtYmyNo").Specific
                oedtSatirNo = frmEDefRapor.Items.Item("edtSatirNo").Specific
            End If

            oDataTable = frmEDefRapor.DataSources.DataTables.Item("DT_0")
            oDataTable2 = frmEDefRapor.DataSources.DataTables.Item("DT_1")
            oDataTable3 = frmEDefRapor.DataSources.DataTables.Item("DT_2")
            oDataTableBaslik = frmEDefRapor.DataSources.DataTables.Item("DT_3")
            oDataTable4 = frmEDefRapor.DataSources.DataTables.Item("DT_4")
            oDataTable5 = frmEDefRapor.DataSources.DataTables.Item("DT_5")

            frmEDefRapor.EnableMenu("772", False)
            frmEDefRapor.EnableMenu("773", False)
            frmEDefRapor.EnableMenu("774", False)
            frmEDefRapor.EnableMenu("1280", False)
            frmEDefRapor.EnableMenu("1283", False)
            frmEDefRapor.EnableMenu("1285", False)
            frmEDefRapor.EnableMenu("1284", False)
            frmEDefRapor.EnableMenu("1286", False)

            '**2 Eğer değişiklik yapılacak bir hareket satırı seçili ise açık gelsin diye...
            If Docentry > 0 Then
                frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                frmEDefRapor.Items.Item("tFolderNo").Specific.Value = Docentry
                frmEDefRapor.Items.Item("1").Click()
                frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                frmEDefRapor.Items.Item("tFolderNo").Enabled = False

                If strEntegratorKod = "1" Then 'uyumsoft
                    'Uyumsoft
                    If oMatrixUyumsoft.VisualRowCount <= 1 Then

                        LoadGrid()
                    ElseIf oDBDSDetails.GetValue("U_Guncellendi", 0) = "Y" Then

                        MessageBox.Show("Seçim Kriterleri Yeni Ya da Mevcut Kriterler Değiştirildi. Verileriniz Kriterlere Uygun Olarak Yüklenmelidir!...", "UYARI", MessageBoxButtons.OK)

                        'LoadGrid()
                    End If

                ElseIf strEntegratorKod = "2" Then 'izibiz
                    'Huber
                    If oMatrixIzibiz.VisualRowCount < 1 Then
                        Dim orsYevmiyeMaddeNo As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim OQuery As String
                        If blnIsHANA Then

                            OQuery = "SELECT IFNULL(MAX(""U_linenumbercounter""),0) + 1  FROM ""@ELRAPVI"" WHERE extract(Year From NOW())=extract(Year From ""U_entereddate"")  AND ""DocEntry""<>" & frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                            orsYevmiyeMaddeNo.DoQuery(OQuery)
                            If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                oedtYmyNo.Value = orsYevmiyeMaddeNo.Fields.Item(0).Value
                                OQuery = "SELECT IFNULL(MAX(""U_linenumber""),0) + 1 FROM ""@ELRAPVI"" WHERE ""U_linenumbercounter"" =" & oedtYmyNo.Value
                                If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                    oedtSatirNo.Value = orsYevmiyeMaddeNo.Fields.Item(0).Value
                                End If
                            Else
                                oGFun.StatusBarErrorMsg("Yevmiye_Madde_No okunamadı... ")
                                boolFormLoaded = False
                            End If
                        Else
                            orsYevmiyeMaddeNo.DoQuery("DECLARE @BY as int, @BS as int SELECT @BY = ISNULL(MAX(U_linenumbercounter),0)+1 FROM [dbo].[@ELRAPVI](NOLOCK) WHERE Datepart(YYYY,getdate()) = Datepart(YYYY,[U_entereddate]) AND DocEntry<>" + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + " SELECT @BS = ISNULL(MAX(U_linenumber),0)+1 FROM [dbo].[@ELRAPVI](NOLOCK) WHERE U_linenumbercounter = (@BY-1) SELECT @BY as BaslangicNo, @BS as BaslangicNo2")
                            If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                orsYevmiyeMaddeNo.MoveFirst()
                                oedtYmyNo.Value = orsYevmiyeMaddeNo.Fields.Item("BaslangicNo").Value
                                oedtSatirNo.Value = orsYevmiyeMaddeNo.Fields.Item("BaslangicNo2").Value
                            Else

                                oGFun.StatusBarErrorMsg("Yevmiye_Madde_No okunamadı... ")
                                boolFormLoaded = False
                            End If
                        End If


                        LoadGrid()
                    ElseIf oDBDSDetails.GetValue("U_Guncellendi", 0) = "Y" Then
                        Dim orsYevmiyeMaddeNo As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim OQuery As String
                        If blnIsHANA Then

                            OQuery = "SELECT IFNULL(MAX(""U_linenumbercounter""),0) + 1  FROM ""@ELRAPVI"" WHERE extract(Year From NOW())=extract(Year From ""U_entereddate"")  AND ""DocEntry""<>" & frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                            orsYevmiyeMaddeNo.DoQuery(OQuery)
                            If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                oedtYmyNo.Value = orsYevmiyeMaddeNo.Fields.Item(0).Value
                                OQuery = "SELECT IFNULL(MAX(""U_linenumber""),0) + 1 FROM ""@ELRAPVI"" WHERE ""U_linenumbercounter"" =" & oedtYmyNo.Value
                                If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                    oedtSatirNo.Value = orsYevmiyeMaddeNo.Fields.Item(0).Value
                                End If
                            Else
                                oGFun.StatusBarErrorMsg("Yevmiye_Madde_No okunamadı... ")
                                boolFormLoaded = False
                            End If
                        Else
                            orsYevmiyeMaddeNo.DoQuery("DECLARE @BY as int, @BS as int SELECT @BY = ISNULL(MAX(U_linenumbercounter),0)+1 FROM [dbo].[@ELRAPVI](NOLOCK) WHERE Datepart(YYYY,getdate()) = Datepart(YYYY,[U_entereddate]) AND DocEntry<>" + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + " SELECT @BS = ISNULL(MAX(U_linenumber),0)+1 FROM [dbo].[@ELRAPVI](NOLOCK) WHERE U_linenumbercounter = (@BY-1) SELECT @BY as BaslangicNo, @BS as BaslangicNo2")
                            If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                orsYevmiyeMaddeNo.MoveFirst()
                                oedtYmyNo.Value = orsYevmiyeMaddeNo.Fields.Item("BaslangicNo").Value
                                oedtSatirNo.Value = orsYevmiyeMaddeNo.Fields.Item("BaslangicNo2").Value
                            Else
                                oGFun.StatusBarErrorMsg("Yevmiye_Madde_No okunamadı... ")
                                boolFormLoaded = False
                            End If

                        End If
                       
                        MessageBox.Show("Seçim Kriterleri Yeni Ya da Mevcut Kriterler Değiştirildi. Verileriniz Kriterlere Uygun Olarak Tekrar Yüklenmelidir!...", "UYARI", MessageBoxButtons.OK)

                        'LoadGrid()
                    Else
                        frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                    End If
                ElseIf strEntegratorKod = "4" Then '4 Huber - izibiz2
                    'Huber
                    If oMatrixHuber.VisualRowCount < 1 Then
                        Dim orsYevmiyeMaddeNo As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim OQuery As String
                        If blnIsHANA Then
                            OQuery = "SELECT IFNULL(MAX(""U_linenumbercounter""),0) + 1  FROM ""@ELRAPVH"" WHERE extract(Year From NOW())=extract(Year From ""U_entereddate"")  AND ""DocEntry""<>" & frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                            orsYevmiyeMaddeNo.DoQuery(OQuery)
                            If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                oedtYmyNo.Value = orsYevmiyeMaddeNo.Fields.Item(0).Value
                                OQuery = "SELECT IFNULL(MAX(""U_linenumber""),0) + 1 FROM ""@ELRAPVH"" WHERE ""U_linenumbercounter"" =" & oedtYmyNo.Value
                                If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                    oedtSatirNo.Value = orsYevmiyeMaddeNo.Fields.Item(0).Value
                                End If
                            Else
                                oGFun.StatusBarErrorMsg("Yevmiye_Madde_No okunamadı... ")
                                boolFormLoaded = False
                            End If
                        Else
                            orsYevmiyeMaddeNo.DoQuery("DECLARE @BY as int, @BS as int SELECT @BY = ISNULL(MAX(U_Yevmiye_Madde_No),0)+1  FROM [dbo].[@ELRAPVH](NOLOCK) WHERE Datepart(YYYY,getdate()) = Datepart(YYYY,[U_Yevmiye_Tarihi]) AND DocEntry<>" + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + " SELECT @BS = ISNULL(MAX(U_Satir_Madde_No),0)+1 FROM [dbo].[@ELRAPVH](NOLOCK) WHERE U_Yevmiye_Madde_No = (@BY-1) SELECT @BY as BaslangicNo, @BS as BaslangicNo2")
                            If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                orsYevmiyeMaddeNo.MoveFirst()
                                oedtYmyNo.Value = orsYevmiyeMaddeNo.Fields.Item("BaslangicNo").Value
                                oedtSatirNo.Value = orsYevmiyeMaddeNo.Fields.Item("BaslangicNo2").Value
                            Else
                                oGFun.StatusBarErrorMsg("Yevmiye_Madde_No okunamadı... ")
                                boolFormLoaded = False
                            End If

                        End If
                       
                        LoadGrid()
                    ElseIf oDBDSDetails.GetValue("U_Guncellendi", 0) = "Y" Then
                        Dim orsYevmiyeMaddeNo As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim OQuery As String
                        If blnIsHANA Then
                            OQuery = "SELECT IFNULL(MAX(""U_linenumbercounter""),0) + 1  FROM ""@ELRAPVH"" WHERE extract(Year From NOW())=extract(Year From ""U_entereddate"")  AND ""DocEntry""<>" & frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                            orsYevmiyeMaddeNo.DoQuery(OQuery)
                            If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                oedtYmyNo.Value = orsYevmiyeMaddeNo.Fields.Item(0).Value
                                OQuery = "SELECT IFNULL(MAX(""U_linenumber""),0) + 1 FROM ""@ELRAPVH"" WHERE ""U_linenumbercounter"" =" & oedtYmyNo.Value
                                If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                    oedtSatirNo.Value = orsYevmiyeMaddeNo.Fields.Item(0).Value
                                End If
                            Else
                                oGFun.StatusBarErrorMsg("Yevmiye_Madde_No okunamadı... ")
                                boolFormLoaded = False
                            End If
                        Else
                            orsYevmiyeMaddeNo.DoQuery("DECLARE @BY as int, @BS as int SELECT @BY = ISNULL(MAX(U_Yevmiye_Madde_No),0)+1  FROM [dbo].[@ELRAPVH](NOLOCK) WHERE Datepart(YYYY,getdate()) = Datepart(YYYY,[U_Yevmiye_Tarihi]) AND DocEntry<>" + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + "  SELECT @BS = ISNULL(MAX(U_Satir_Madde_No),0)+1 FROM [dbo].[@ELRAPVH](NOLOCK) WHERE U_Yevmiye_Madde_No = (@BY-1) SELECT @BY as BaslangicNo, @BS as BaslangicNo2")
                            If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                orsYevmiyeMaddeNo.MoveFirst()
                                oedtYmyNo.Value = orsYevmiyeMaddeNo.Fields.Item("BaslangicNo").Value
                                oedtSatirNo.Value = orsYevmiyeMaddeNo.Fields.Item("BaslangicNo2").Value
                            Else
                                oGFun.StatusBarErrorMsg("Yevmiye_Madde_No okunamadı... ")
                                boolFormLoaded = False
                            End If

                        End If
                        
                        MessageBox.Show("Seçim Kriterleri Yeni Ya da Mevcut Kriterler Değiştirildi. Verileriniz Kriterlere Uygun Olarak Tekrar Yüklenmelidir!...", "UYARI", MessageBoxButtons.OK)

                        'LoadGrid()
                    Else
                        frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                    End If

                ElseIf strEntegratorKod = "3" Then '3 Bimsa
                    'Bimsa
                    If oMatrixBimsa.VisualRowCount < 1 Then
                        Dim orsYevmiyeMaddeNo As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim OQuery As String
                        If blnIsHANA Then
                            OQuery = "SELECT IFNULL(MAX(""U_linenumbercounter""),0) + 1  FROM ""@ELRAPVB"" WHERE extract(Year From NOW())=extract(Year From ""U_entereddate"")  AND ""DocEntry""<>" & frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                            orsYevmiyeMaddeNo.DoQuery(OQuery)
                            If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                oedtYmyNo.Value = orsYevmiyeMaddeNo.Fields.Item(0).Value
                                OQuery = "SELECT IFNULL(MAX(""U_linenumber""),0) + 1 FROM ""@ELRAPVB"" WHERE ""U_linenumbercounter"" =" & oedtYmyNo.Value
                                If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                    oedtSatirNo.Value = orsYevmiyeMaddeNo.Fields.Item(0).Value
                                End If
                            Else
                                oGFun.StatusBarErrorMsg("Yevmiye_Madde_No okunamadı... ")
                                boolFormLoaded = False
                            End If
                        Else
                            orsYevmiyeMaddeNo.DoQuery("DECLARE @BY as int, @BS as int SELECT @BY = ISNULL(MAX(U_linenumbercounter),0)+1 FROM [dbo].[@ELRAPVB](NOLOCK) WHERE Datepart(YYYY,getdate()) = Datepart(YYYY,[U_entereddate]) AND DocEntry<>" + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + " SELECT @BS = ISNULL(MAX(U_linenumber),0)+1 FROM [dbo].[@ELRAPVB](NOLOCK) WHERE U_linenumbercounter = (@BY-1) SELECT @BY as BaslangicNo, @BS as BaslangicNo2")
                            If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                orsYevmiyeMaddeNo.MoveFirst()
                                oedtYmyNo.Value = orsYevmiyeMaddeNo.Fields.Item("BaslangicNo").Value
                                oedtSatirNo.Value = orsYevmiyeMaddeNo.Fields.Item("BaslangicNo2").Value
                            Else
                                oGFun.StatusBarErrorMsg("Yevmiye_Madde_No okunamadı... ")
                                boolFormLoaded = False
                            End If
                        End If
                       LoadGrid()
                    ElseIf oDBDSDetails.GetValue("U_Guncellendi", 0) = "Y" Then
                        Dim orsYevmiyeMaddeNo As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim OQuery As String
                        If blnIsHANA Then
                            OQuery = "SELECT IFNULL(MAX(""U_linenumbercounter""),0) + 1  FROM ""@ELRAPVB"" WHERE extract(Year From NOW())=extract(Year From ""U_entereddate"")  AND ""DocEntry""<>" & frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                            orsYevmiyeMaddeNo.DoQuery(OQuery)
                            If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                oedtYmyNo.Value = orsYevmiyeMaddeNo.Fields.Item(0).Value
                                OQuery = "SELECT IFNULL(MAX(""U_linenumber""),0) + 1 FROM ""@ELRAPVB"" WHERE ""U_linenumbercounter"" =" & oedtYmyNo.Value
                                If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                    oedtSatirNo.Value = orsYevmiyeMaddeNo.Fields.Item(0).Value
                                End If
                            Else
                                oGFun.StatusBarErrorMsg("Yevmiye_Madde_No okunamadı... ")
                                boolFormLoaded = False
                            End If
                        Else
                            orsYevmiyeMaddeNo.DoQuery("DECLARE @BY as int, @BS as int SELECT @BY = ISNULL(MAX(U_linenumbercounter),0)+1 FROM [dbo].[@ELRAPVB](NOLOCK) WHERE Datepart(YYYY,getdate()) = Datepart(YYYY,[U_entereddate]) AND DocEntry<>" + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + " SELECT @BS = ISNULL(MAX(U_linenumber),0)+1 FROM [dbo].[@ELRAPVB](NOLOCK) WHERE U_linenumbercounter = (@BY-1) SELECT @BY as BaslangicNo, @BS as BaslangicNo2")
                            If orsYevmiyeMaddeNo.RecordCount > 0 Then
                                orsYevmiyeMaddeNo.MoveFirst()
                                oedtYmyNo.Value = orsYevmiyeMaddeNo.Fields.Item("BaslangicNo").Value
                                oedtSatirNo.Value = orsYevmiyeMaddeNo.Fields.Item("BaslangicNo2").Value
                            Else
                                oGFun.StatusBarErrorMsg("Yevmiye_Madde_No okunamadı... ")
                                boolFormLoaded = False
                            End If

                        End If
                       
                        MessageBox.Show("Seçim Kriterleri Yeni Ya da Mevcut Kriterler Değiştirildi. Verileriniz Kriterlere Uygun Olarak Tekrar Yüklenmelidir!...", "UYARI", MessageBoxButtons.OK)

                        'LoadGrid()
                    Else
                        frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                    End If
                End If

                Exit Sub
            Else
                frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE 'Form modu sevda
                frmEDefRapor.Items.Item("tFolderNo").Enabled = False
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
            frmEDefRapor.Freeze(True)
            oGFun.LoadComboBoxSeries(frmEDefRapor.Items.Item("15").Specific, "ELRAP")


            frmEDefRapor.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText("InitForm Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            frmEDefRapor.Freeze(False)
        Finally
            frmEDefRapor.Freeze(False)
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    Try
                        Select Case pVal.ItemUID
                            Case "tFolderNo"
                                If pVal.BeforeAction Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                        End Select
                    Catch ex As Exception

                    End Try

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Try
                        Select Case pVal.ItemUID
                            Case "1"
                                If pVal.BeforeAction And (frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_EDIT_MODE) Then
                                    If Me.ValidateALL = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        BubbleEvent = True
                                    End If
                                ElseIf pVal.BeforeAction And (frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                                    If Me.ValidateALL = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        oDBDSDetails.SetValue("U_Guncellendi", oDBDSDetails.Offset, "N")
                                        BubbleEvent = True
                                    End If

                                End If
                                If (pVal.ActionSuccess = True And frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                                    frmEDefRapor.Items.Item("2").Click()
                                    'InitForm()
                                End If

                            Case "33"
                                If (pVal.ActionSuccess = True) Then
                                    LoadGrid()
                                End If


                            Case "21" 'Excele Aktar
                                If pVal.BeforeAction Then
                                    If Me.ValidateALL = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        Dim orsEntegrator As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        If blnIsHANA Then
                                            orsEntegrator.DoQuery("SELECT ""U_IntegratorCode"", ""U_IntegratorName""  FROM ""@EINTEGSEL""")
                                        Else
                                            orsEntegrator.DoQuery("SELECT [U_IntegratorCode], [U_IntegratorName]  FROM [dbo].[@EINTEGSEL](NOLOCK)")
                                        End If

                                        If orsEntegrator.RecordCount > 0 Then
                                            orsEntegrator.MoveFirst()
                                            strEntegratorKod = orsEntegrator.Fields.Item(0).Value
                                            strEntegrator = orsEntegrator.Fields.Item(1).Value
                                        Else
                                            oGFun.StatusBarErrorMsg(" Entegratör seçimini yapmalısınız... ")
                                            boolFormLoaded = False
                                            Exit Sub
                                        End If

                                        If strEntegratorKod = "1" Then
                                            ExceleYazUyumsoft()
                                            BubbleEvent = True
                                        ElseIf strEntegratorKod = "2" Then 'İzibiz
                                            ExceleYazIzibiz()
                                            BubbleEvent = True
                                        ElseIf strEntegratorKod = "4" Then '"4" huber

                                            ExceleYazHuber()
                                            BubbleEvent = True
                                        ElseIf strEntegratorKod = "3" Then 'Bimsa
                                            TextYazBimsa()
                                            BubbleEvent = True
                                        End If

                                    End If

                                End If


                        End Select
                    Catch ex As Exception
                        oGFun.StatusBarErrorMsg("Item Pressed Event Failed 2: " & ex.Message)
                    Finally
                    End Try


                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    Select Case pVal.ItemUID
                        Case "15"
                            If pVal.ItemChanged And pVal.BeforeAction = False Then
                                oDBDSHeader.SetValue("DocNum", 0, frmEDefRapor.BusinessObject.GetNextSerialNumber(frmEDefRapor.Items.Item("15").Specific.Selected.Value))
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
            frmEDefRapor.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmEDefRapor.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
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
            'Dim Baslik As String = oDBDSHeader.GetValue("U_RBASLIK", 0).Trim


            'If Baslik = String.Empty Or Baslik = "" Then
            '    oGFun.StatusBarErrorMsg("Seçim Ölçüt Adı Girmelisiniz...")
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
            oEDefListeRaporlar.LoadGrid()
        Catch ex As Exception

        End Try
    End Sub

    Sub LoadGrid()
        Try
            If strEntegratorKod = "1" Then ''Uyumsoft
                Dim oQuery As String
                If blnIsHANA Then
                    'sp_EDefterGrid1
                    oQuery = "Call SP_EDEFTERGRID1(" + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + ", '" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "')"
                Else
                    oQuery = "EXECUTE dbo.sp_EDefterGrid1 " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + ", '" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "'"
                End If


                oDataTable.Clear()
                oDataTable.ExecuteQuery(oQuery)

                'MatrisDoldur(oDataTable)

                frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                'frmEDefRapor.Items.Item("tFolderNo").Specific.Value = strDoc.ToString() 'frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                frmEDefRapor.Items.Item("tFolderNo").Specific.Value = frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                frmEDefRapor.Items.Item("1").Click()
                frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                frmEDefRapor.Items.Item("tFolderNo").Enabled = False
                frmEDefRapor.Items.Item("1").Click()

            ElseIf strEntegratorKod = "2" Then 'İzibiz

                Dim oQuery2 As String
                'Huber
                If frmEDefRapor.Items.Item("edtYmyNo").Specific.Value.ToString().Trim = "" Then

                    If blnIsHANA Then
                        'sp_EDefterGrid3
                        oQuery2 = "Call SP_EDEFTERGRID3 ('" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "', " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + ",1,1)"
                    Else
                        oQuery2 = "EXECUTE dbo.sp_EDefterGrid3 '" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "', " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + ",1,1"
                    End If


                Else
                    If blnIsHANA Then
                        'sp_EDefterGrid3
                        oQuery2 = "Call SP_EDEFTERGRID3 ('" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "', " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + "," + frmEDefRapor.Items.Item("edtYmyNo").Specific.Value.ToString() + "," + frmEDefRapor.Items.Item("edtSatirNo").Specific.Value.ToString() & ")"
                    Else
                        oQuery2 = "EXECUTE dbo.sp_EDefterGrid3 '" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "', " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + "," + frmEDefRapor.Items.Item("edtYmyNo").Specific.Value.ToString() + "," + frmEDefRapor.Items.Item("edtSatirNo").Specific.Value.ToString()
                    End If


                End If

                oDataTable4.Clear()
                oDataTable4.ExecuteQuery(oQuery2)

                frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                'frmEDefRapor.Items.Item("tFolderNo").Specific.Value = strDoc.ToString() 'frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                frmEDefRapor.Items.Item("tFolderNo").Specific.Value = frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                frmEDefRapor.Items.Item("1").Click()
                frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                frmEDefRapor.Items.Item("tFolderNo").Enabled = False
                frmEDefRapor.Items.Item("1").Click()

            ElseIf strEntegratorKod = "4" Then 'Huber

                Dim oQuery2 As String
                'Huber
                If frmEDefRapor.Items.Item("edtYmyNo").Specific.Value.ToString().Trim = "" Then
                    If blnIsHANA Then
                        'oQuery2 = "Call sp_EDefterGrid2 '" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "', " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + ",1,1"
                    Else
                        oQuery2 = "EXECUTE dbo.sp_EDefterGrid2 '" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "', " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + ",1,1"
                    End If


                Else
                    If blnIsHANA Then
                        'oQuery2 = "Call sp_EDefterGrid2 '" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "', " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + "," + frmEDefRapor.Items.Item("edtYmyNo").Specific.Value.ToString() + "," + frmEDefRapor.Items.Item("edtSatirNo").Specific.Value.ToString()
                    Else
                        oQuery2 = "EXECUTE dbo.sp_EDefterGrid2 '" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "', " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + "," + frmEDefRapor.Items.Item("edtYmyNo").Specific.Value.ToString() + "," + frmEDefRapor.Items.Item("edtSatirNo").Specific.Value.ToString()
                    End If


                End If

                oDataTable3.Clear()
                oDataTable3.ExecuteQuery(oQuery2)

                frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                'frmEDefRapor.Items.Item("tFolderNo").Specific.Value = strDoc.ToString() 'frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                frmEDefRapor.Items.Item("tFolderNo").Specific.Value = frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                frmEDefRapor.Items.Item("1").Click()
                frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                frmEDefRapor.Items.Item("tFolderNo").Enabled = False
                frmEDefRapor.Items.Item("1").Click()

            ElseIf strEntegratorKod = "3" Then 'Bimsa

                Dim oQuery2 As String

                If frmEDefRapor.Items.Item("edtYmyNo").Specific.Value.ToString().Trim = "" Then
                    If blnIsHANA Then
                        'sp_EDefterGrid4
                        oQuery2 = "Call SP_EDEFTERGRID4 ('" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "', " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + ",1,1 )"
                    Else
                        oQuery2 = "EXECUTE dbo.sp_EDefterGrid4 '" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "', " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + ",1,1"
                    End If


                Else
                    If blnIsHANA Then
                        oQuery2 = "Call SP_EDEFTERGRID4 ('" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "', " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + "," + frmEDefRapor.Items.Item("edtYmyNo").Specific.Value.ToString() + "," + frmEDefRapor.Items.Item("edtSatirNo").Specific.Value.ToString() & ")"
                    Else
                        oQuery2 = "EXECUTE dbo.sp_EDefterGrid4 '" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "', " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + "," + frmEDefRapor.Items.Item("edtYmyNo").Specific.Value.ToString() + "," + frmEDefRapor.Items.Item("edtSatirNo").Specific.Value.ToString()
                    End If
                End If

                oDataTable5.Clear()
                oDataTable5.ExecuteQuery(oQuery2)

                frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                'frmEDefRapor.Items.Item("tFolderNo").Specific.Value = strDoc.ToString() 'frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                frmEDefRapor.Items.Item("tFolderNo").Specific.Value = frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                frmEDefRapor.Items.Item("1").Click()
                frmEDefRapor.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                frmEDefRapor.Items.Item("tFolderNo").Enabled = False
                frmEDefRapor.Items.Item("1").Click()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Sub MatrisDoldur(ByVal oTable As SAPbouiCOM.DataTable)
        Dim deger As Int16 = 0
        'Dim progress As SAPbouiCOM.ProgressBar

        Try

            'progress = oApplication.StatusBar.CreateProgressBar("Veri Gride Aktarılıyor", oTable.Rows.Count, True)

            'progress.Value = 0

            oMatrixUyumsoft.Clear()

            If oTable.IsEmpty = False Then
                If oTable.GetValue("StartDate", 0).ToString() = "0" Then
                    oApplication.StatusBar.SetText("Uyarı: Sectiginiz Kriterlere Uygun Kayit Bulunamadi... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    'frmEDefRapor.Items.Item("edtCount").Specific.Value = 0
                    Exit Sub
                End If

                oMatrixUyumsoft.FlushToDataSource()

                Dim strDocType As String = ""

                For i As Integer = 0 To oTable.Rows.Count - 1
                    oMatrixUyumsoft.AddRow()

                    oMatrixUyumsoft.SetLineData(oMatrixUyumsoft.VisualRowCount)

                    strDocType = oTable.GetValue("EntryHeader.DocumentType", i).ToString().Trim

                    'MessageBox.Show(oTable.GetValue("EntryHeader.EntryDetail.Amount", i).ToString())
                    oMatrixUyumsoft.Columns.Item("#").Cells.Item(i + 1).Specific.Value = i + 1
                    oMatrixUyumsoft.Columns.Item("Col_0").Cells.Item(i + 1).Specific.Value = oTable.GetValue("StartDate", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_1").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EndDate", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_2").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.EnteredBy", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_3").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.EnteredDate", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_4").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.EntryNumber", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_5").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.EntryComment", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_6").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.BatchID", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_7").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.BatchDescription", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_8").Cells.Item(i + 1).Specific.Value = Replace(oTable.GetValue("EntryHeader.TotalDebit", i).ToString(), strOndalikAyirac, ".")
                    'oMatrixUyumsoft.Columns.Item("Col_8").Cells.Item(i + 1).Specific.Value = Replace(oTable.GetValue("EntryHeader.TotalDebit", i).ToString(), ".", strOndalikAyirac)
                    oMatrixUyumsoft.Columns.Item("Col_9").Cells.Item(i + 1).Specific.Value = Replace(oTable.GetValue("EntryHeader.TotalCredit", i).ToString(), strOndalikAyirac, ".")
                    oMatrixUyumsoft.Columns.Item("Col_10").Cells.Item(i + 1).Specific.Value = strDocType 'oTable.GetValue("EntryHeader.DocumentType", i).ToString()
                    If strDocType = "" Then
                        oMatrixUyumsoft.Columns.Item("Col_11").Cells.Item(i + 1).Specific.Value = ""
                        oMatrixUyumsoft.Columns.Item("Col_12").Cells.Item(i + 1).Specific.Value = ""
                        oMatrixUyumsoft.Columns.Item("Col_13").Cells.Item(i + 1).Specific.Value = ""
                    Else
                        oMatrixUyumsoft.Columns.Item("Col_11").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.DocumentTypeDescription", i).ToString()
                        oMatrixUyumsoft.Columns.Item("Col_12").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.DocumentNumber", i).ToString()
                        oMatrixUyumsoft.Columns.Item("Col_13").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.DocumentDate", i).ToString()
                    End If

                    oMatrixUyumsoft.Columns.Item("Col_14").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.PaymentMethod", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_15").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.EntryDetail.Account.AccountMainID", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_16").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.EntryDetail.Account.AccountMainDescription", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_17").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.EntryDetail.Account.AccountSubDescription", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_18").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.EntryDetail.Account.AccountSubID", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_19").Cells.Item(i + 1).Specific.Value = Replace(oTable.GetValue("EntryHeader.EntryDetail.Amount", i).ToString(), strOndalikAyirac, ".")
                    oMatrixUyumsoft.Columns.Item("Col_20").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.EntryDetail.DebitCreditCode", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_21").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.EntryDetail.PostingDate", i).ToString()
                    oMatrixUyumsoft.Columns.Item("Col_22").Cells.Item(i + 1).Specific.Value = oTable.GetValue("EntryHeader.EntryDetail.DetailComment", i).ToString()

                    'progress.Value = progress.Value +1
                    'deger = progress.Value
                Next
                'frmEDefRapor.Items.Item("edtCount").Specific.Value = oDataTable.Rows.Count.ToString
            Else
                'frmEDefRapor.Items.Item("edtCount").Specific.Value = 0
                'progress.Stop()
                'progress = Nothing
                oGFun.StatusBarWarningMsg("Uygun kayıt bulunamamıştır...")

                Exit Sub
            End If

            'progress.Stop()
            'progress = Nothing

        Catch ex As Exception
            MessageBox.Show(deger.ToString())
            'progress.Stop()
            'progress = Nothing
            oApplication.StatusBar.SetText("Hata: Metod MatrisDoldur() ..  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub

    Sub ExceleYazUyumsoft()
        Dim systemCultureInfo As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Try

            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            Dim ExcelUygulama As Microsoft.Office.Interop.Excel.Application
            Dim ExcelProje As Microsoft.Office.Interop.Excel.Workbook
            Dim ExcelSayfa As Microsoft.Office.Interop.Excel.Worksheet
            Dim Missing As Object = System.Reflection.Missing.Value
            Dim ExcelRange As Microsoft.Office.Interop.Excel.Range

            Dim rowCnt As Integer = 0
            Dim columnCnt As Integer = 0

            Dim s_dosyaadi As String = ""
            Dim s_veri As String = ""
            Dim strDocType As String = ""


            ExcelUygulama = New Microsoft.Office.Interop.Excel.Application()
            ExcelProje = ExcelUygulama.Workbooks.Add(Missing)
            ExcelSayfa = CType(ExcelProje.Worksheets.Item(1), Microsoft.Office.Interop.Excel.Worksheet)
            ExcelRange = ExcelSayfa.UsedRange
            ExcelSayfa = CType(ExcelUygulama.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

            ExcelUygulama.Visible = False
            ExcelUygulama.AlertBeforeOverwriting = False

            Dim bolge As Microsoft.Office.Interop.Excel.Range

            'Dim oQuery2 As String = "EXECUTE dbo.sp_EDefterExcel1 '" + frmEDefRapor.Items.Item("Item_29").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_30").Specific.Value.ToString() + "', '" + frmEDefRapor.Items.Item("Item_33").Specific.Value.ToString() + "','" + frmEDefRapor.Items.Item("Item_39").Specific.Value.ToString() + "'"
            Dim oQuery2 As String
            If blnIsHANA Then
                'sp_EDefterExcel1
                oQuery2 = "Call SP_EDEFTEREXCEL1 (" + frmEDefRapor.Items.Item("17").Specific.Value.ToString() & ")"
            Else
                oQuery2 = "EXECUTE dbo.sp_EDefterExcel1 " + frmEDefRapor.Items.Item("17").Specific.Value.ToString()
            End If

            oDataTable2.Clear()
            oDataTable2.ExecuteQuery(oQuery2)
            If oDataTable2.Rows.Count > 0 Then
                MessageBox.Show("Excele Veri Aktarımına Başlanıyor...")
                For rowCnt = 1 To oDataTable2.Rows.Count + 1
                    For columnCnt = 1 To oDataTable2.Columns.Count
                        If rowCnt = 1 Then
                            bolge = CType(ExcelSayfa.Cells(1, columnCnt), Microsoft.Office.Interop.Excel.Range)
                            If columnCnt = 1 Then
                                bolge.Value2 = "StartDate"
                            ElseIf columnCnt = 2 Then
                                bolge.Value2 = "EndDate"
                            ElseIf columnCnt = 3 Then
                                bolge.Value2 = "EntryHeader.EnteredBy"
                            ElseIf columnCnt = 4 Then
                                bolge.Value2 = "EntryHeader.EnteredDate"
                            ElseIf columnCnt = 5 Then
                                bolge.Value2 = "EntryHeader.EntryNumber"
                            ElseIf columnCnt = 6 Then
                                bolge.Value2 = "EntryHeader.EntryComment"
                            ElseIf columnCnt = 7 Then
                                bolge.Value2 = "EntryHeader.BatchID"
                            ElseIf columnCnt = 8 Then
                                bolge.Value2 = "EntryHeader.BatchDescription"
                            ElseIf columnCnt = 9 Then
                                bolge.Value2 = "EntryHeader.TotalDebit"
                            ElseIf columnCnt = 10 Then
                                bolge.Value2 = "EntryHeader.TotalCredit"
                            ElseIf columnCnt = 11 Then
                                bolge.Value2 = "EntryHeader.DocumentType"
                            ElseIf columnCnt = 12 Then
                                bolge.Value2 = "EntryHeader.DocumentTypeDescription"
                            ElseIf columnCnt = 13 Then
                                bolge.Value2 = "EntryHeader.DocumentNumber"
                            ElseIf columnCnt = 14 Then
                                bolge.Value2 = "EntryHeader.DocumentDate"
                            ElseIf columnCnt = 15 Then
                                bolge.Value2 = "EntryHeader.PaymentMethod"
                            ElseIf columnCnt = 16 Then
                                bolge.Value2 = "EntryHeader.EntryDetail.Account.AccountMainID"
                            ElseIf columnCnt = 17 Then
                                bolge.Value2 = "EntryHeader.EntryDetail.Account.AccountMainDescription"
                            ElseIf columnCnt = 18 Then
                                bolge.Value2 = "EntryHeader.EntryDetail.Account.AccountSubDescription"
                            ElseIf columnCnt = 19 Then
                                bolge.Value2 = "EntryHeader.EntryDetail.Account.AccountSubID"
                            ElseIf columnCnt = 20 Then
                                bolge.Value2 = "EntryHeader.EntryDetail.Amount"
                            ElseIf columnCnt = 21 Then
                                bolge.Value2 = "EntryHeader.EntryDetail.DebitCreditCode"
                            ElseIf columnCnt = 22 Then
                                bolge.Value2 = "EntryHeader.EntryDetail.PostingDate"
                            ElseIf columnCnt = 23 Then
                                bolge.Value2 = "EntryHeader.EntryDetail.DetailComment"
                            End If
                        ElseIf columnCnt > 1 And columnCnt < 7 Then
                            bolge = CType(ExcelSayfa.Cells(rowCnt, columnCnt - 1), Microsoft.Office.Interop.Excel.Range)
                            bolge.Value2 = oDataTable2.GetValue(columnCnt - 2, rowCnt - 2).ToString()
                        ElseIf columnCnt = 7 Then 'prosedürden dönen SatirNo excele yazılmasın diye geçiyoruz

                        ElseIf columnCnt > 7 And columnCnt < 13 Then
                            bolge = CType(ExcelSayfa.Cells(rowCnt, columnCnt - 2), Microsoft.Office.Interop.Excel.Range)
                            bolge.Value2 = oDataTable2.GetValue(columnCnt - 2, rowCnt - 2).ToString()
                        ElseIf columnCnt = 13 Then
                            bolge = CType(ExcelSayfa.Cells(rowCnt, columnCnt - 2), Microsoft.Office.Interop.Excel.Range)
                            strDocType = oDataTable2.GetValue(columnCnt - 2, rowCnt - 2).ToString().Trim
                            bolge.Value2 = strDocType

                        ElseIf columnCnt = 14 Or columnCnt = 15 Or columnCnt = 16 Then
                            bolge = CType(ExcelSayfa.Cells(rowCnt, columnCnt - 2), Microsoft.Office.Interop.Excel.Range)
                            If strDocType = "" Then
                                bolge.Value2 = ""
                            Else
                                bolge.Value2 = oDataTable2.GetValue(columnCnt - 2, rowCnt - 2).ToString()
                            End If
                        ElseIf columnCnt > 16 Then
                            bolge = CType(ExcelSayfa.Cells(rowCnt, columnCnt - 2), Microsoft.Office.Interop.Excel.Range)
                            bolge.Value2 = oDataTable2.GetValue(columnCnt - 2, rowCnt - 2).ToString()
                        End If

                    Next
                Next
                MessageBox.Show("Aktarım tamamlandı, dosya kaydedilecek..")
            
            End If


            'MessageBox.Show(Environment.GetFolderPath(Environment.SpecialFolder.Desktop).ToString())

            'Kaydetme
            s_dosyaadi = frmEDefRapor.Items.Item("Item_5").Specific.Value.ToString() 'Item_5
            If s_dosyaadi <> "" Then
                'MessageBox.Show("Dosyanın kaydedilecegi dizin ve dosya: " + System.Windows.Forms.Application.StartupPath + "\" + s_dosyaadi + ".xlsx")
                'ExcelProje.SaveAs(System.Windows.Forms.Application.StartupPath + "\" + s_dosyaadi + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Missing, Missing, False, Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
                ExcelProje.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.Desktop).ToString() + "\" + s_dosyaadi + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Missing, Missing, False, Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
                ExcelProje.Close(True, Missing, Missing)
                ExcelUygulama.Quit()

                MessageBox.Show("Dosya Başarıyla Kaydedildi.")
            Else
                MessageBox.Show("Lütfen Bir Dosya Adı Giriniz.")
            End If

        Catch ex As Exception
            oGFun.StatusBarErrorMsg("1 Excele yazma hatası..." & ex.Message)
        Finally
            System.Threading.Thread.CurrentThread.CurrentCulture = systemCultureInfo
        End Try
        
    End Sub
    Sub ExceleYazHuber()
        Dim systemCultureInfo As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Try

            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            Dim ExcelUygulama As Microsoft.Office.Interop.Excel.Application
            Dim ExcelProje As Microsoft.Office.Interop.Excel.Workbook
            Dim ExcelSayfa As Microsoft.Office.Interop.Excel.Worksheet
            Dim ExcelSayfa2 As Microsoft.Office.Interop.Excel.Worksheet
            Dim Missing As Object = System.Reflection.Missing.Value
            Dim ExcelRange As Microsoft.Office.Interop.Excel.Range
            'Dim ExcelRange2 As Microsoft.Office.Interop.Excel.Range

            Dim rowCnt As Integer = 0
            Dim columnCnt As Integer = 0

            Dim s_dosyaadi As String = ""
            Dim s_veri As String = ""
            Dim strDocType As String = ""


            ExcelUygulama = New Microsoft.Office.Interop.Excel.Application()
            ExcelProje = ExcelUygulama.Workbooks.Add(Missing)

            ExcelSayfa = CType(ExcelProje.Worksheets.Item(1), Microsoft.Office.Interop.Excel.Worksheet)
            ExcelSayfa.Name = "Defterdar_Csv_Baslik_Import_Tem"
            ExcelRange = ExcelSayfa.UsedRange

            ExcelSayfa = CType(ExcelUygulama.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

            ExcelUygulama.Visible = False
            ExcelUygulama.AlertBeforeOverwriting = False

            Dim bolge As Microsoft.Office.Interop.Excel.Range
            Dim bolge2 As Microsoft.Office.Interop.Excel.Range
            Dim oQueryBaslik As String
            If blnIsHANA Then
                oQueryBaslik = "SELECT DISTINCT ""U_Yevmiye_Madde_No"", ""U_Fis_Numarasi"", CAST(""U_Yevmiye_Tarihi"" AS varchar(10)), ""U_Fisi_Kaydeden"", ""U_Kayit_Acik"", ""U_Toplam_Borc"", ""U_Toplam_Alacak"", ""U_Firma_No"", ""U_Sube_No"", ""U_Kaynak_Referansi"" FROM ""@ELRAPVH"" WHERE ""DocEntry"" = " + frmEDefRapor.Items.Item("17").Specific.Value.ToString()
            Else
                oQueryBaslik = "SELECT DISTINCT [U_Yevmiye_Madde_No],[U_Fis_Numarasi],CONVERT(VARCHAR(10), [U_Yevmiye_Tarihi], 120) ,[U_Fisi_Kaydeden],[U_Kayit_Acik],[U_Toplam_Borc],[U_Toplam_Alacak],[U_Firma_No],[U_Sube_No],[U_Kaynak_Referansi] FROM [dbo].[@ELRAPVH](NOLOCK) WHERE DocEntry = " + frmEDefRapor.Items.Item("17").Specific.Value.ToString()
            End If

            oDataTableBaslik.Clear()
            oDataTableBaslik.ExecuteQuery(oQueryBaslik)

            If oDataTableBaslik.Rows.Count > 0 Then
                MessageBox.Show("Excele Veri Aktarımına Başlanıyor...")
                For rowCnt = 1 To oDataTableBaslik.Rows.Count + 1
                    For columnCnt = 1 To 10
                        If rowCnt = 1 Then
                            bolge = CType(ExcelSayfa.Cells(1, columnCnt), Microsoft.Office.Interop.Excel.Range)
                            If columnCnt = 1 Then
                                bolge.Value2 = "Yevmiye_Madde_Numarası"
                            ElseIf columnCnt = 2 Then
                                bolge.Value2 = "Fiş_Numarası"
                            ElseIf columnCnt = 3 Then
                                bolge.Value2 = "Yevmiye_Tarihi"
                            ElseIf columnCnt = 4 Then
                                bolge.Value2 = "Fişi_Kaydeden"
                            ElseIf columnCnt = 5 Then
                                bolge.Value2 = "Kayıt_Açıklaması"
                            ElseIf columnCnt = 6 Then
                                bolge.Value2 = "Toplam_Borç"
                            ElseIf columnCnt = 7 Then
                                bolge.Value2 = "Toplam_Alacak"
                            ElseIf columnCnt = 8 Then
                                bolge.Value2 = "Firma_No"
                            ElseIf columnCnt = 9 Then
                                bolge.Value2 = "Şube_No"
                            ElseIf columnCnt = 10 Then
                                bolge.Value2 = "Kaynak_Referansı"
                            End If

                        ElseIf rowCnt > 1 Then
                            bolge = CType(ExcelSayfa.Cells(rowCnt, columnCnt), Microsoft.Office.Interop.Excel.Range)
                            bolge.Value2 = oDataTableBaslik.GetValue(columnCnt - 1, rowCnt - 2).ToString()
                        End If
                    Next
                Next

                Dim oQuery2 As String
                If blnIsHANA Then
                    oQuery2 = "SELECT ""U_Yevmiye_Madde_No"", ""U_Fis_Numarasi"", CAST(""U_Yevmiye_Tarihi"" AS varchar(10)), ""U_Fisi_Kaydeden"", ""U_Kayit_Acik"", ""U_Toplam_Borc"", ""U_Toplam_Alacak"", ""U_Firma_No"", ""U_Sube_No"", ""U_Kaynak_Referansi"", ""U_Fis_Satir_No"", ""U_Satir_Madde_No"", ""U_Ana_Hesap_Kodu"", ""U_Ana_Hesap_Acik"", ""U_Alt_Hesap_Kodu"", ""U_Alt_Hesap_Acik"", ""U_Borc"", ""U_Alacak"", ""U_Dokuman_No"", ""U_Dokuman_Tipi"", ""U_Dokuman_Tipi_Acik"", CAST(""U_Dokuman_Tarihi"" AS varchar(10)), ""U_Odeme_Yontemi"", ""U_Fis_Detay_Acik"" FROM ""@ELRAPVH"" WHERE ""DocEntry"" = " + frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                Else
                    oQuery2 = "SELECT [U_Yevmiye_Madde_No],[U_Fis_Numarasi],CONVERT(VARCHAR(10), [U_Yevmiye_Tarihi], 120) ,[U_Fisi_Kaydeden],[U_Kayit_Acik],[U_Toplam_Borc],[U_Toplam_Alacak],[U_Firma_No],[U_Sube_No],[U_Kaynak_Referansi],[U_Fis_Satir_No],[U_Satir_Madde_No],[U_Ana_Hesap_Kodu],[U_Ana_Hesap_Acik],[U_Alt_Hesap_Kodu],[U_Alt_Hesap_Acik],[U_Borc],[U_Alacak],[U_Dokuman_No],[U_Dokuman_Tipi],[U_Dokuman_Tipi_Acik],CONVERT(VARCHAR(10), [U_Dokuman_Tarihi], 120),[U_Odeme_Yontemi],[U_Fis_Detay_Acik] FROM [dbo].[@ELRAPVH](NOLOCK) WHERE DocEntry = " + frmEDefRapor.Items.Item("17").Specific.Value.ToString()
                End If

                oDataTable2.Clear()
                oDataTable2.ExecuteQuery(oQuery2)
                'İkinci Sayfa
                ExcelSayfa2 = ExcelProje.Worksheets(2)
                'ExcelSayfa2 = CType(ExcelProje.Worksheets.Item(2), Microsoft.Office.Interop.Excel.Worksheet)
                ExcelSayfa2.Name = "Defterdar_Csv_Detay_Import_Temp"

                ' ExcelSayfa2 = CType(ExcelUygulama.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                bolge2 = ExcelSayfa2.UsedRange

                rowCnt = 1
                columnCnt = 1

                For rowCnt = 1 To oDataTable2.Rows.Count + 1
                    For columnCnt = 1 To 18
                        If rowCnt = 1 Then
                            bolge2 = CType(ExcelSayfa2.Cells(1, columnCnt), Microsoft.Office.Interop.Excel.Range)
                            If columnCnt = 1 Then
                                bolge2.Value2 = "Satır_Madde_No"
                            ElseIf columnCnt = 2 Then
                                bolge2.Value2 = "Fiş_Numarası"
                            ElseIf columnCnt = 3 Then
                                bolge2.Value2 = "Ana_Hesap_Kodu"
                            ElseIf columnCnt = 4 Then
                                bolge2.Value2 = "Ana_Hesap_Açıklaması"
                            ElseIf columnCnt = 5 Then
                                bolge2.Value2 = "Alt_Hesap_Kodu"
                            ElseIf columnCnt = 6 Then
                                bolge2.Value2 = "Alt_Hesap_Açıklaması"
                            ElseIf columnCnt = 7 Then
                                bolge2.Value2 = "Borç"

                            ElseIf columnCnt = 8 Then
                                bolge2.Value2 = "Alacak"
                            ElseIf columnCnt = 9 Then
                                bolge2.Value2 = "Fiş_Satır_No"
                            ElseIf columnCnt = 10 Then
                                bolge2.Value2 = "Doküman_No"
                            ElseIf columnCnt = 11 Then
                                bolge2.Value2 = "Doküman_Tipi"
                            ElseIf columnCnt = 12 Then
                                bolge2.Value2 = "Doküman_Tipi_Açıklaması"
                            ElseIf columnCnt = 13 Then
                                bolge2.Value2 = "Doküman_Tarihi"
                            ElseIf columnCnt = 14 Then
                                bolge2.Value2 = "Ödeme_Yöntemi"
                            ElseIf columnCnt = 15 Then
                                bolge2.Value2 = "Fiş_Detay_Açıklaması"
                            ElseIf columnCnt = 16 Then
                                bolge2.Value2 = "Firma_No"
                            ElseIf columnCnt = 17 Then
                                bolge2.Value2 = "Şube_No"
                            ElseIf columnCnt = 18 Then
                                bolge2.Value2 = "Kaynak_Referansı"
                            End If
                        Else
                            bolge2 = CType(ExcelSayfa2.Cells(rowCnt, columnCnt), Microsoft.Office.Interop.Excel.Range)
                            If columnCnt = 1 Then
                                bolge2.Value2 = oDataTable2.GetValue(11, rowCnt - 2).ToString()
                            ElseIf columnCnt = 2 Then
                                bolge2.Value2 = oDataTable2.GetValue(1, rowCnt - 2).ToString()
                            ElseIf columnCnt = 3 Then
                                bolge2.Value2 = oDataTable2.GetValue(12, rowCnt - 2).ToString()
                            ElseIf columnCnt = 4 Then
                                bolge2.Value2 = oDataTable2.GetValue(13, rowCnt - 2).ToString()
                            ElseIf columnCnt = 5 Then
                                bolge2.Value2 = oDataTable2.GetValue(14, rowCnt - 2).ToString()
                            ElseIf columnCnt = 6 Then
                                bolge2.Value2 = oDataTable2.GetValue(15, rowCnt - 2).ToString()
                            ElseIf columnCnt = 7 Then
                                bolge2.Value2 = oDataTable2.GetValue(16, rowCnt - 2).ToString()
                            ElseIf columnCnt = 8 Then
                                bolge2.Value2 = oDataTable2.GetValue(17, rowCnt - 2).ToString()
                            ElseIf columnCnt = 9 Then
                                bolge2.Value2 = oDataTable2.GetValue(10, rowCnt - 2).ToString()
                            ElseIf columnCnt = 10 Then
                                bolge2.Value2 = oDataTable2.GetValue(18, rowCnt - 2).ToString()
                            ElseIf columnCnt = 11 Then
                                bolge2.Value2 = oDataTable2.GetValue(19, rowCnt - 2).ToString()
                            ElseIf columnCnt = 12 Then
                                bolge2.Value2 = oDataTable2.GetValue(20, rowCnt - 2).ToString()
                            ElseIf columnCnt = 13 Then
                                If oDataTable2.GetValue(18, rowCnt - 2).ToString() = "" Then
                                    bolge2.Value2 = ""
                                Else
                                    bolge2.Value2 = oDataTable2.GetValue(21, rowCnt - 2).ToString()
                                End If
                            ElseIf columnCnt = 14 Then
                                bolge2.Value2 = oDataTable2.GetValue(22, rowCnt - 2).ToString()
                            ElseIf columnCnt = 15 Then
                                bolge2.Value2 = oDataTable2.GetValue(23, rowCnt - 2).ToString()
                            ElseIf columnCnt = 16 Then
                                bolge2.Value2 = oDataTable2.GetValue(7, rowCnt - 2).ToString()
                            ElseIf columnCnt = 17 Then
                                bolge2.Value2 = oDataTable2.GetValue(8, rowCnt - 2).ToString()
                            ElseIf columnCnt = 18 Then
                                bolge2.Value2 = oDataTable2.GetValue(9, rowCnt - 2).ToString()
                            End If

                        End If

                    Next
                Next
                MessageBox.Show("Aktarım tamamlandı, dosya kaydedilecek..")
            
            End If


            'MessageBox.Show(Environment.GetFolderPath(Environment.SpecialFolder.Desktop).ToString())

            'Kaydetme
            s_dosyaadi = frmEDefRapor.Items.Item("Item_5").Specific.Value.ToString() 'Item_5
            If s_dosyaadi <> "" Then
                'MessageBox.Show("Dosyanın kaydedilecegi dizin ve dosya: " + System.Windows.Forms.Application.StartupPath + "\" + s_dosyaadi + ".xlsx")
                'ExcelProje.SaveAs(System.Windows.Forms.Application.StartupPath + "\" + s_dosyaadi + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Missing, Missing, False, Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
                ExcelProje.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.Desktop).ToString() + "\" + s_dosyaadi + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Missing, Missing, False, Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
                ExcelProje.Close(True, Missing, Missing)
                ExcelUygulama.Quit()

                MessageBox.Show("Dosya Başarıyla Kaydedildi.")
            Else
                MessageBox.Show("Lütfen Bir Dosya Adı Giriniz.")
            End If


        Catch ex As Exception
            oGFun.StatusBarErrorMsg("4 Excele yazma hatası..." & ex.Message)
        Finally
            System.Threading.Thread.CurrentThread.CurrentCulture = systemCultureInfo
        End Try

    End Sub

    Sub ExceleYazIzibiz()
        Dim ss As Int16
        Dim systemCultureInfo As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Try

            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            Dim ExcelUygulama As Microsoft.Office.Interop.Excel.Application
            Dim ExcelProje As Microsoft.Office.Interop.Excel.Workbook
            Dim ExcelSayfa As Microsoft.Office.Interop.Excel.Worksheet
            Dim Missing As Object = System.Reflection.Missing.Value
            Dim ExcelRange As Microsoft.Office.Interop.Excel.Range

            Dim rowCnt As Integer = 0
            Dim columnCnt As Integer = 0

            Dim s_dosyaadi As String = ""
            Dim s_veri As String = ""
            Dim strDocType As String = ""


            ExcelUygulama = New Microsoft.Office.Interop.Excel.Application()
            ExcelProje = ExcelUygulama.Workbooks.Add(Missing)
            ExcelSayfa = CType(ExcelProje.Worksheets.Item(1), Microsoft.Office.Interop.Excel.Worksheet)
            ExcelRange = ExcelSayfa.UsedRange
            ExcelSayfa = CType(ExcelUygulama.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

            ExcelUygulama.Visible = False
            ExcelUygulama.AlertBeforeOverwriting = False

            Dim bolge As Microsoft.Office.Interop.Excel.Range

            ''İzibiz
            Dim oQuery2 As String
            If blnIsHANA Then
                'sp_EDefterExcel3
                oQuery2 = "Call SP_EDEFTEREXCEL3 (" + frmEDefRapor.Items.Item("17").Specific.Value.ToString() & ")"
            Else
                oQuery2 = "EXECUTE dbo.sp_EDefterExcel3 " + frmEDefRapor.Items.Item("17").Specific.Value.ToString()
            End If


            oDataTable4.Clear()
            oDataTable4.ExecuteQuery(oQuery2)
            If oDataTable4.Rows.Count > 0 Then
                MessageBox.Show("Excele Veri Aktarımına Başlanıyor...")
                For rowCnt = 1 To oDataTable4.Rows.Count + 1
                    For columnCnt = 1 To oDataTable4.Columns.Count
                        If rowCnt = 1 Then
                            bolge = CType(ExcelSayfa.Cells(1, columnCnt), Microsoft.Office.Interop.Excel.Range)
                            If columnCnt = 1 Then
                                bolge.Value2 = "detailref"
                            ElseIf columnCnt = 2 Then
                                bolge.Value2 = "entryref"
                            ElseIf columnCnt = 3 Then
                                bolge.Value2 = "linenumber"
                            ElseIf columnCnt = 4 Then
                                bolge.Value2 = "linenumbercounter"
                            ElseIf columnCnt = 5 Then
                                bolge.Value2 = "accmainid"
                            ElseIf columnCnt = 6 Then
                                bolge.Value2 = "accmainiddesc"
                            ElseIf columnCnt = 7 Then
                                bolge.Value2 = "accsubid"
                            ElseIf columnCnt = 8 Then
                                bolge.Value2 = "accsubdesc"
                            ElseIf columnCnt = 9 Then
                                bolge.Value2 = "amount"
                            ElseIf columnCnt = 10 Then
                                bolge.Value2 = "debitcreditcode"
                            ElseIf columnCnt = 11 Then
                                bolge.Value2 = "postingdate"
                            ElseIf columnCnt = 12 Then
                                bolge.Value2 = "documenttype"
                            ElseIf columnCnt = 13 Then
                                bolge.Value2 = "doctypedesc"
                            ElseIf columnCnt = 14 Then
                                bolge.Value2 = "documentnumber"
                            ElseIf columnCnt = 15 Then
                                bolge.Value2 = "documentreference"
                            ElseIf columnCnt = 16 Then
                                bolge.Value2 = "entrynumbercounter"
                            ElseIf columnCnt = 17 Then
                                bolge.Value2 = "documentdate"
                            ElseIf columnCnt = 18 Then
                                bolge.Value2 = "paymentmethod"
                            ElseIf columnCnt = 19 Then
                                bolge.Value2 = "detailcomment"
                            ElseIf columnCnt = 20 Then
                                bolge.Value2 = "erpno"
                            ElseIf columnCnt = 21 Then
                                bolge.Value2 = "divisionno"
                            ElseIf columnCnt = 22 Then
                                bolge.Value2 = "enteredby"
                            ElseIf columnCnt = 23 Then
                                bolge.Value2 = "entereddate"
                            ElseIf columnCnt = 24 Then
                                bolge.Value2 = "entrynumber"
                            ElseIf columnCnt = 25 Then
                                bolge.Value2 = "entrycomment"
                            End If
                        ElseIf rowCnt > 1 Then
                            'ss = "row:" + rowCnt.ToString + ",col:" + columnCnt.ToString + "-" + oDataTable4.GetValue(columnCnt - 1, rowCnt - 2).ToString() 'rowCnt.ToString
                            bolge = CType(ExcelSayfa.Cells(rowCnt, columnCnt), Microsoft.Office.Interop.Excel.Range)
                            bolge.Value2 = oDataTable4.GetValue(columnCnt - 1, rowCnt - 2).ToString()

                        End If

                    Next
                Next
                MessageBox.Show("Aktarım tamamlandı, dosya kaydedilecek..")

            End If

            'Kaydetme
            s_dosyaadi = frmEDefRapor.Items.Item("Item_5").Specific.Value.ToString() 'Item_5
            If s_dosyaadi <> "" Then
                'MessageBox.Show("Dosyanın kaydedilecegi dizin ve dosya: " + System.Windows.Forms.Application.StartupPath + "\" + s_dosyaadi + ".xlsx")
                'ExcelProje.SaveAs(System.Windows.Forms.Application.StartupPath + "\" + s_dosyaadi + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Missing, Missing, False, Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
                ExcelProje.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.Desktop).ToString() + "\" + s_dosyaadi + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Missing, Missing, False, Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange)
                ExcelProje.Close(True, Missing, Missing)
                ExcelUygulama.Quit()

                MessageBox.Show("Dosya Başarıyla Kaydedildi.")
            Else
                MessageBox.Show("Lütfen Bir Dosya Adı Giriniz.")
            End If

        Catch ex As Exception
            oGFun.StatusBarErrorMsg("2 Excele yazma hatası..." & ss & "    ---- " & ex.Message)
        Finally
            System.Threading.Thread.CurrentThread.CurrentCulture = systemCultureInfo
        End Try

    End Sub

    Sub TextYazBimsa()
        Dim s_dosyaadi As String = ""
        Dim strYazilacakVeri As String = ""

        Using oGetFileName As New CoreFrieght_Intraspeed.GetFileNameClass(1)
            'oGetFileName.Filter = "Excel files (*.csv)|*.csv"
            oGetFileName.Filter = "Text files (*.txt)|*.txt"
            oGetFileName.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal)
            Dim threadGetExcelFile As New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf oGetFileName.GetFileName))
            threadGetExcelFile.SetApartmentState(Threading.ApartmentState.STA)

            Try
                threadGetExcelFile.Start()
                While Not threadGetExcelFile.IsAlive
                    'oGFun.StatusBarErrorMsg("3 Text yazma thread hatası oluştu!...")
                    Exit Sub
                End While
                ' Wait for thread to get started
                System.Threading.Thread.Sleep(1)
                ' Wait a sec more
                threadGetExcelFile.Join()
                ' Wait for thread to end
                Dim fileName = String.Empty
                fileName = oGetFileName.FileName

                If fileName <> String.Empty Then
                    s_dosyaadi = fileName
                Else
                    Exit Sub
                End If


                Dim fs As IO.FileStream = New IO.FileStream(s_dosyaadi, IO.FileMode.OpenOrCreate, IO.FileAccess.Write, IO.FileShare.None)
                Dim sw As IO.StreamWriter = New IO.StreamWriter(fs)
                Try
                    Dim inty As Int16 = 0
                    Dim oQueryBaslik As String
                    If blnIsHANA Then
                        'sp_EDefterText4
                        oQueryBaslik = "Call SP_EDEFTERTEXT4 (" + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + ",'" + SirketAdi + "')"
                    Else
                        oQueryBaslik = "Execute [dbo].[sp_EDefterText4] " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + ",'" + SirketAdi + "'"
                    End If

                    oDataTable5.Clear()
                    oDataTable5.ExecuteQuery(oQueryBaslik)

                    If oDataTable5.Rows.Count > 0 Then
                        MessageBox.Show("Text Veri Aktarımına Başlanıyor...")
                        'MessageBox.Show(oDataTable5.GetValue(0, 1).ToString())

                        For inty = 1 To oDataTable5.Rows.Count
                            strYazilacakVeri = oDataTable5.GetValue(0, inty - 1).ToString()
                            sw.WriteLine(strYazilacakVeri)
                        Next

                    End If

                    MessageBox.Show("Dosya Başarıyla Kaydedildi.")

                Catch ex As Exception
                    oGFun.StatusBarErrorMsg("3 Text yazma hatası..." & ex.Message)
                Finally
                    sw.Close()
                    fs.Close()
                End Try
            Catch
            End Try
        End Using
    End Sub

    'Sub TextYazBimsa()
    '    Dim s_dosyaadi As String = ""
    '    Dim strYazilacakVeri As String = ""
    '    GetFileHeader()
    '    's_dosyaadi = frmEDefRapor.Items.Item("Item_5").Specific.Value.ToString()
    '    's_dosyaadi = Environment.GetFolderPath(Environment.SpecialFolder.Desktop).ToString() + "\" + s_dosyaadi + ".txt"
    '    fillopen()
    '    If strSelectedFilepath <> "" Then
    '        s_dosyaadi = strSelectedFilepath
    '    Else
    '        Exit Sub
    '    End If


    '    Dim fs As IO.FileStream = New IO.FileStream(s_dosyaadi, IO.FileMode.Open, IO.FileAccess.Write, IO.FileShare.None)
    '    Dim sw As IO.StreamWriter = New IO.StreamWriter(fs)
    '    Try
    '        Dim inty As Int16 = 0
    '        Dim oQueryBaslik As String = "Execute [dbo].[sp_EDefterText4] " + frmEDefRapor.Items.Item("17").Specific.Value.ToString() + ",'" + SirketAdi + "'"
    '        oDataTable5.Clear()
    '        oDataTable5.ExecuteQuery(oQueryBaslik)

    '        If oDataTable5.Rows.Count > 0 Then
    '            MessageBox.Show("Text Veri Aktarımına Başlanıyor...")
    '            'MessageBox.Show(oDataTable5.GetValue(0, 1).ToString())

    '            For inty = 1 To oDataTable5.Rows.Count
    '                strYazilacakVeri = oDataTable5.GetValue(0, inty - 1).ToString()
    '                'If inty >= 158 Then
    '                '    MessageBox.Show(strYazilacakVeri)
    '                'End If
    '                sw.WriteLine(strYazilacakVeri)
    '            Next

    '        End If

    '        MessageBox.Show("Dosya Başarıyla Kaydedildi.")

    '    Catch ex As Exception
    '        oGFun.StatusBarErrorMsg("3 Text yazma hatası..." & ex.Message)
    '    Finally
    '        sw.Close()
    '        fs.Close()
    '    End Try
    'End Sub


    'Kodun İlk Hali
    'Private Sub GetFileHeader()
    '    Using oGetFileName As New CoreFrieght_Intraspeed.GetFileNameClass(1)
    '        'oGetFileName.Filter = "Excel files (*.csv)|*.csv"
    '        oGetFileName.Filter = "Text files (*.txt)|*.txt"
    '        oGetFileName.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal)
    '        Dim threadGetExcelFile As New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf oGetFileName.GetFileName))
    '        threadGetExcelFile.SetApartmentState(Threading.ApartmentState.STA)

    '        Try
    '            threadGetExcelFile.Start()
    '            While Not threadGetExcelFile.IsAlive


    '            End While
    '            ' Wait for thread to get started
    '            System.Threading.Thread.Sleep(1)
    '            ' Wait a sec more
    '            threadGetExcelFile.Join()
    '            ' Wait for thread to end
    '            Dim fileName = String.Empty
    '            fileName = oGetFileName.FileName

    '            If fileName <> String.Empty Then
    '            End If
    '        Catch
    '        End Try
    '    End Using
    'End Sub

End Class
