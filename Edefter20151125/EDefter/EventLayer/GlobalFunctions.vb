Imports System.Reflection
Imports System.IO
Imports SAPbouiCOM
Imports System.Xml

''' <summary>
''' Globally whatever Function and method do you want define here 
''' We can use any class and module from here  
''' </summary>
''' <remarks></remarks>
Public Class GlobalFunctions

#Region " ...  Common For Company ..."

    Sub AddXML(ByVal pathstr As String)
        Try
            Dim xmldoc As New Xml.XmlDocument
            Dim stream As System.IO.Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("EDefter." + pathstr)
            Dim streamreader As New System.IO.StreamReader(stream, True)
            xmldoc.LoadXml(streamreader.ReadToEnd())
            streamreader.Close()
            oApplication.LoadBatchActions(xmldoc.InnerXml)
        Catch ex As Exception
            oApplication.StatusBar.SetText("AddXML Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Public Function FormExist(ByVal FormID As String) As Boolean
        FormExist = False
        For Each uid As SAPbouiCOM.Form In oApplication.Forms
            If uid.UniqueID = FormID Then
                FormExist = True
                Exit Function
            End If
        Next
        If FormExist Then
            oApplication.Forms.Item(FormID).Visible = True
            oApplication.Forms.Item(FormID).Select()
        End If
    End Function

    Public Function ConnectionContext() As Integer
        Try

            Dim strErrorCode As String
            If oCompany.Connected = True Then oCompany.Disconnect()

            oApplication.StatusBar.SetText("EDefter Modülü ADDON Bağlantısı Yapılıyor..........      Lütfen Bekleyiniz ..........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            strErrorCode = oCompany.Connect
            ConnectionContext = strErrorCode
            If strErrorCode = 0 Then
                SirketAdi = oCompany.CompanyName
                oApplication.StatusBar.SetText("EDefter Modülü ADDON - Bağlantı Başarılı... ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                System.Media.SystemSounds.Asterisk.Play()

            Else
                oApplication.StatusBar.SetText("EDefter Modülü ADDON - Bağlantı Başarısız, Hata Açıklaması: " & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Function

    Public Function CookieConnect() As Integer
        Try
            Dim strCkie, strContext As String
            oCompany = New SAPbobsCOM.Company
            Debug.Print(oCompany.CompanyDB)
            strCkie = oCompany.GetContextCookie()
            strContext = oApplication.Company.GetConnectionContext(strCkie)
            CookieConnect = oCompany.SetSboLoginContext(strContext)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Function

    Public Sub SetApplication()
        Try
            Dim oGUI As New SAPbouiCOM.SboGuiApi
            oGUI.AddonIdentifier = ""
            oGUI.Connect(Environment.GetCommandLineArgs.GetValue(1).ToString())
            oApplication = oGUI.GetApplication()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Sub
    Sub Msg(ByVal strMsg As String, Optional ByVal msgTime As String = "S", Optional ByVal errType As String = "W")
        Dim time As SAPbouiCOM.BoMessageTime
        Dim msgType As SAPbouiCOM.BoStatusBarMessageType
        Select Case errType.ToUpper()
            Case "E"
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Error
            Case "W"
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Warning
            Case "N"
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_None
            Case "S"
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Success
            Case Else
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Warning
        End Select
        Select Case msgTime.ToUpper()
            Case "M"
                time = SAPbouiCOM.BoMessageTime.bmt_Medium
            Case "S"
                time = SAPbouiCOM.BoMessageTime.bmt_Short
            Case "L"
                time = SAPbouiCOM.BoMessageTime.bmt_Long
            Case Else
                time = SAPbouiCOM.BoMessageTime.bmt_Medium
        End Select
        oApplication.StatusBar.SetText(strMsg, time, msgType)
    End Sub
#End Region

#Region " ... Menu Creation ..."

    Sub LoadXML(ByVal Form As SAPbouiCOM.Form, ByVal FormId As String, ByVal FormXML As String)
        Try
            AddXML(FormXML)
            Form = oApplication.Forms.Item(FormId)
            Form.Select()
        Catch ex As Exception
            oApplication.StatusBar.SetText("LoadXML Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

#End Region

#Region "       Common For Data Base Creation ...   "

    Public Function UDOExists(ByVal code As String) As Boolean
        GC.Collect()
        Dim v_UDOMD As SAPbobsCOM.UserObjectsMD
        Dim v_ReturnCode As Boolean
        v_UDOMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        v_ReturnCode = v_UDOMD.GetByKey(code)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UDOMD)
        v_UDOMD = Nothing
        Return v_ReturnCode
    End Function

    Function DataExists(ByVal TableName As String) As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oFlag As Boolean = True
            If blnIsHANA Then
                rs.DoQuery("Select 1 from """ & Trim(TableName) & """ ")
            Else
                rs.DoQuery("Select 1 from [" & Trim(TableName) & "] ")
            End If

            If rs.EoF Then oFlag = False
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            rs = Nothing
            GC.Collect()
            Return oFlag
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function
    Function InsertMasterData(ByVal TableName As String, ByVal sql As String) As Boolean
        InsertMasterData = False
        Dim v_RetVal As Long
        Dim v_ErrCode As Long
        Dim v_ErrMsg As String = ""
        Try
            If Not Me.DataExists(TableName) Then
                oApplication.StatusBar.SetText(TableName + " Master Data Kaydediliyor...................", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rs.DoQuery(sql)

                If v_RetVal <> 0 Then
                    oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                    oApplication.StatusBar.SetText("Failed to Insert Master Data for " & TableName & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else
                    oApplication.StatusBar.SetText("[" & TableName & "] -  Inserted Master Data!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Return True
                End If
            Else
                GC.Collect()
                Return False
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(addonName & ":> " & ex.Message & " @ " & ex.Source)
        End Try
    End Function

    Function CreateTable(ByVal TableName As String, ByVal TableDesc As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
        CreateTable = False
        Dim v_RetVal As Long
        Dim v_ErrCode As Long
        Dim v_ErrMsg As String = ""
        Try
            If Not Me.TableExists(TableName) Then
                Dim v_UserTableMD As SAPbobsCOM.UserTablesMD
                oApplication.StatusBar.SetText("Creating Table " + TableName + " ...................", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                v_UserTableMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                v_UserTableMD.TableName = TableName
                v_UserTableMD.TableDescription = TableDesc
                v_UserTableMD.TableType = TableType
                v_RetVal = v_UserTableMD.Add()
                If v_RetVal <> 0 Then
                    oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                    oApplication.StatusBar.SetText("Failed to Create Table " & TableDesc & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD)
                    v_UserTableMD = Nothing
                    Return False
                Else
                    oApplication.StatusBar.SetText("[" & TableName & "] - " & TableDesc & " Created Successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD)
                    v_UserTableMD = Nothing
                    Return True
                End If
            Else
                GC.Collect()
                Return False
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(addonName & ":> " & ex.Message & " @ " & ex.Source)
        End Try
    End Function
    Public Function removeValidValues(ByRef _combo As SAPbouiCOM.ComboBox) As Boolean
        Try

            While _combo.ValidValues.Count > 0
                _combo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
            End While
            Return True
        Catch ex As Exception
            '   B1Connections.theAppl.SetStatusBarMessage(ex.Message)
            Return False
        End Try

    End Function
    Function ColumnExists(ByVal TableName As String, ByVal FieldID As String) As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oFlag As Boolean = True
            Dim strQuery As String = "Select 1 from ""CUFD"" Where ""TableID""='" & Trim(TableName.ToUpper) & "' and ""AliasID""='" & Trim(FieldID) & "'"
            rs.DoQuery(strQuery)
            If rs.EoF Then oFlag = False
            rs = Nothing
            GC.Collect()
            Return oFlag
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function UDFExists(ByVal TableName As String, ByVal FieldID As String) As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oFlag As Boolean = True
            rs.DoQuery("Select 1 from ""CUFD"" Where ""TableID""='" & Trim(TableName) & "' and ""AliasID""='" & Trim(FieldID) & "'")
            If rs.EoF Then oFlag = False
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            rs = Nothing
            GC.Collect()
            Return oFlag
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function
    Function TriggerExists(ByVal ProcedureName As String) As Boolean
        Dim oFlag As Boolean
        Dim oObjectExists As String = "if exists (select * from dbo.sysobjects where name = '" + ProcedureName + "'  and OBJECTPROPERTY(id, 'IsTrigger') = 1 )  SELECT 'True' ELSE SELECT 'False'"
        Dim oRsObjectExists As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsObjectExists.DoQuery(oObjectExists)
        oFlag = oRsObjectExists.Fields.Item(0).Value
        Return oFlag
    End Function
    Function CreateTrigger(ByVal TriggerName As String, ByVal TriggerText As String) As Boolean
        CreateTrigger = False
        Dim v_RetVal As Long
        Dim v_ErrCode As Long
        Dim v_ErrMsg As String = ""
        Try
            If Not Me.TriggerExists(TriggerName) Then
                oApplication.StatusBar.SetText("Creating Trigger " + TriggerName + " ...................", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                If v_RetVal <> 0 Then
                    oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                    oApplication.StatusBar.SetText("Failed to Create Trigger " & TriggerName & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    Return False
                Else
                    'Dim oFlag As Boolean
                    Dim oRsObjectExists As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsObjectExists.DoQuery(TriggerText)
                    'oFlag = oRsObjectExists.Fields.Item(0).Value
                    oApplication.StatusBar.SetText("[" & TriggerName & "] - Created Successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    Return True
                End If
            Else
                GC.Collect()
                Return False
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(addonName & ":> " & ex.Message & " @ " & ex.Source)
        End Try
    End Function
    Function ProcedureExists(ByVal ProcedureName As String) As Boolean
        Dim oFlag As Boolean
        Dim oRsObjectExists As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oObjectExists As String

        If blnIsHANA Then
            oObjectExists = "SELECT * FROM ""PROCEDURES"" WHERE ""PROCEDURE_NAME"" = '" + ProcedureName + "'"
            oRsObjectExists.DoQuery(oObjectExists)
            If oRsObjectExists.RecordCount > 0 Then
                oFlag = True
            Else
                oFlag = False
            End If
        Else
            oObjectExists = "IF OBJECT_ID('" + ProcedureName + "', 'P') IS NOT NULL SELECT 'True' ELSE SELECT 'False'"
            oRsObjectExists.DoQuery(oObjectExists)
            oFlag = oRsObjectExists.Fields.Item(0).Value
        End If
        Return oFlag
    End Function
    Function CreateProcedure(ByVal ProcedureName As String, ByVal ProcedureText As String) As Boolean
        CreateProcedure = False
        Dim v_RetVal As Long
        Dim v_ErrCode As Long
        Dim v_ErrMsg As String = ""
        Try
            If Not Me.ProcedureExists(ProcedureName) Then
                oApplication.StatusBar.SetText("Creating Procedure " + ProcedureName + " ...................", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                If v_RetVal <> 0 Then
                    oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                    oApplication.StatusBar.SetText("Failed to Create Procedure " & ProcedureName & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    Return False
                Else
                    Dim oFlag As Boolean
                    Dim oRsObjectExists As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsObjectExists.DoQuery(ProcedureText)
                    If Not blnIsHANA Then
                        oFlag = oRsObjectExists.Fields.Item(0).Value
                    End If
                    oApplication.StatusBar.SetText("[" & ProcedureName & "] - Created Successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    Return True
                End If
            Else
                GC.Collect()
                Return False
            End If
        Catch ex As Exception
            If ex.Message.Contains("Invalid index") = False Then
                oApplication.StatusBar.SetText(addonName & ":> " & ex.Message & " @ " & ex.Source, BoMessageTime.bmt_Medium)
            End If

        End Try
    End Function

    Function TableExists(ByVal TableName As String) As Boolean
        Try
            Dim oTables As SAPbobsCOM.UserTablesMD
            Dim oFlag As Boolean
            oTables = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            oFlag = oTables.GetByKey(TableName)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oTables)
            Return oFlag
        Catch ex As Exception
            oApplication.StatusBar.SetText(addonName & ":> " & ex.Message & " @ " & ex.Source)
        End Try

    End Function


    Function CreateUserFieldsComboBox(ByVal TableName As String, ByVal FieldName As String, ByVal FieldDescription As String, ByVal type As SAPbobsCOM.BoFieldTypes, Optional ByVal size As Long = 0, Optional ByVal subType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal LinkedTable As String = "", Optional ByVal ComboValidValues As String(,) = Nothing, Optional ByVal DefaultValidValues As String = "") As Boolean
        Try
            'If TableName.StartsWith("@") = False Then
            If Not Me.UDFExists(TableName, FieldName) Then
                Dim v_UserField As SAPbobsCOM.UserFieldsMD
                v_UserField = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                v_UserField.TableName = TableName
                v_UserField.Name = FieldName
                v_UserField.Description = FieldDescription
                v_UserField.Type = type
                If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                    If size <> 0 Then
                        v_UserField.Size = size
                    End If
                End If
                If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                    v_UserField.SubType = subType
                End If

                For i As Int16 = 0 To ComboValidValues.GetLength(0) - 1
                    If i > 0 Then v_UserField.ValidValues.Add()
                    v_UserField.ValidValues.Value = ComboValidValues(i, 0)
                    v_UserField.ValidValues.Description = ComboValidValues(i, 1)
                Next
                If DefaultValidValues <> "" Then v_UserField.DefaultValue = DefaultValidValues

                If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                v_RetVal = v_UserField.Add()
                If v_RetVal <> 0 Then
                    oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                    oApplication.StatusBar.SetText("Failed to add UserField " & FieldDescription & " - " & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                    v_UserField = Nothing
                    Return False
                Else
                    oApplication.StatusBar.SetText(" & TableName & - " & FieldDescription & " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                    v_UserField = Nothing
                    Return True
                End If

            Else
                Return False
            End If
            ' End If
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try
    End Function

    Function CreateUserFields(ByVal TableName As String, ByVal FieldName As String, ByVal FieldDescription As String, ByVal type As SAPbobsCOM.BoFieldTypes, Optional ByVal size As Long = 0, Optional ByVal subType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal LinkedTable As String = "", Optional ByVal DefaultValue As String = "") As Boolean
        Try
            If TableName.StartsWith("@") = True Then
                If Not Me.ColumnExists(TableName, FieldName) Then
                    Dim v_UserField As SAPbobsCOM.UserFieldsMD
                    v_UserField = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    v_UserField.TableName = TableName
                    v_UserField.Name = FieldName
                    v_UserField.Description = FieldDescription
                    v_UserField.Type = type
                    If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                        If size <> 0 Then
                            If type = SAPbobsCOM.BoFieldTypes.db_Numeric Then
                                v_UserField.EditSize = 11
                            Else
                                v_UserField.Size = size
                            End If


                        End If
                    End If
                    If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                        v_UserField.SubType = subType
                    End If
                    If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                    If DefaultValue <> "" Then v_UserField.DefaultValue = DefaultValue

                    v_RetVal = v_UserField.Add()
                    If v_RetVal <> 0 Then
                        oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                        oApplication.StatusBar.SetText("Failed to add UserField masterid" & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return False
                    Else
                        oApplication.StatusBar.SetText("[" & TableName & "] - " & FieldDescription & " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return True
                    End If
                Else
                    Return False
                End If
            End If

            If TableName.StartsWith("@") = False Then
                If Not Me.UDFExists(TableName, FieldName) Then
                    Dim v_UserField As SAPbobsCOM.UserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    v_UserField.TableName = TableName
                    v_UserField.Name = FieldName
                    v_UserField.Description = FieldDescription
                    v_UserField.Type = type
                    If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                        If size <> 0 Then
                            If type = SAPbobsCOM.BoFieldTypes.db_Numeric Then
                                v_UserField.EditSize = 11
                            Else
                                v_UserField.Size = size
                            End If

                        End If
                    End If
                    If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                        v_UserField.SubType = subType
                    End If
                    If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                    v_RetVal = v_UserField.Add()
                    If v_RetVal <> 0 Then
                        oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                        oApplication.StatusBar.SetText("Failed to add UserField " & FieldDescription & " - " & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return False
                    Else
                        oApplication.StatusBar.SetText(" & TableName & - " & FieldDescription & " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return True
                    End If

                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try
    End Function

    Function RegisterUDO(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal FindField As String(,), ByVal UDOHTableName As String, Optional ByVal UDODTableName As String = "", Optional ByVal ChildTable As String = "", Optional ByVal ChildTable1 As String = "", _
    Optional ByVal ChildTable2 As String = "", Optional ByVal ChildTable3 As String = "", Optional ByVal ChildTable4 As String = "", Optional ByVal ChildTable5 As String = "", _
    Optional ByVal ChildTable6 As String = "", Optional ByVal ChildTable7 As String = "", Optional ByVal ChildTable8 As String = "", Optional ByVal ChildTable9 As String = "", _
    Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim ActionSuccess As Boolean = False
        Try
            RegisterUDO = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = UDOHTableName

            If UDODTableName <> "" Then
                v_udoMD.ChildTables.TableName = UDODTableName
                v_udoMD.ChildTables.Add()
            End If

            If ChildTable <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable1 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable1
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable2 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable2
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable3 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable3
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable4 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable4
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable5 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable5
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable6 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable6
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable7 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable7
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable8 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable8
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable9 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable9
                v_udoMD.ChildTables.Add()
            End If

            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                v_udoMD.LogTableName = "A" & UDOHTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To FindField.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = FindField(i, 0)
                v_udoMD.FindColumns.ColumnDescription = FindField(i, 1)
            Next

            If v_udoMD.Add() = 0 Then
                RegisterUDO = True
                If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                oApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                oApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                MessageBox.Show(oCompany.GetLastErrorDescription)
                RegisterUDO = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
            If ActionSuccess = False And oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        Catch ex As Exception
            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End Try
    End Function

    Function RegisterUDOForDefaultForm(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal FindField As String(,), ByVal UDOHTableName As String, Optional ByVal UDODTableName As String = "", Optional ByVal ChildTable As String = "", _
   Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim ActionSuccess As Boolean = False
        Try
            RegisterUDOForDefaultForm = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = UDOHTableName

            If UDODTableName <> "" Then
                v_udoMD.ChildTables.TableName = UDODTableName
                v_udoMD.ChildTables.Add()
            End If

            If ChildTable <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable
                v_udoMD.ChildTables.Add()
            End If
            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                v_udoMD.LogTableName = "A" & UDOHTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To FindField.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = FindField(i, 0)
                v_udoMD.FindColumns.ColumnDescription = FindField(i, 1)
            Next
            For i As Int16 = 0 To FindField.GetLength(0) - 1
                If i > 0 Then v_udoMD.FormColumns.Add()
                v_udoMD.FormColumns.FormColumnAlias = FindField(i, 0)
                v_udoMD.FormColumns.FormColumnDescription = FindField(i, 1)
            Next

            If v_udoMD.Add() = 0 Then
                RegisterUDOForDefaultForm = True
                If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                oApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                oApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                'MessageBox.Show(oCompany.GetLastErrorDescription)
                RegisterUDOForDefaultForm = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
            If ActionSuccess = False And oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        Catch ex As Exception
            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End Try
    End Function

#End Region

#Region "       Functions  & Methods        "

    Function GetServerDate() As String
        Try
            Dim rsetBob As SAPbobsCOM.SBObob = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            Dim rsetServerDate As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            rsetServerDate = rsetBob.Format_StringToDate(oApplication.Company.ServerDate())

            Return CDate(rsetServerDate.Fields.Item(0).Value).ToString("yyyyMMdd")

        Catch ex As Exception
            oGFun.StatusBarErrorMsg("Get Server Date Function Failed : " & ex.Message)
            Return ""
        Finally
        End Try
    End Function

    Function DoQuery(ByVal strSql As String) As SAPbobsCOM.Recordset
        Try
            Dim rsetCode As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsetCode.DoQuery(strSql)
            Return rsetCode
        Catch ex As Exception
            oApplication.StatusBar.SetText("Execute Query Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return Nothing
        Finally
        End Try
    End Function

    Function SetDateFormate(ByVal StrDate As String) As String
        Try

            Dim strsql
            If blnIsHANA Then
                strsql = "SELECT CAST(CAST('" & StrDate & "' AS timestamp) AS char(10)) AS ""HRS"" FROM DUMMY"
            Else
                strsql = "SELECT CONVERT(char(10),CONVERT(DATETIME,'" & _
                                                            StrDate & "'), 103) HRS"
            End If
             Dim rsetPayMethod As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsetPayMethod.DoQuery(strsql)
            Return rsetPayMethod.Fields.Item("HRS").Value
        Catch ex As Exception
            oApplication.StatusBar.SetText("Add Date Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function
    Function SetDateFormat2(ByVal StrDate As String) As String '20141112 formatında gelen veriyi 2014-11-12 şeklinde geri döndürür
        Try
            Dim strsql
            If blnIsHANA Then
                strsql = "SELECT CAST(CAST('" & StrDate & "' AS timestamp) AS char(10)) AS ""HRS"" FROM DUMMY"

            Else
                strsql = "SELECT CONVERT(char(10),CONVERT(DATETIME,'" & _
                                                           StrDate & "'), 120) HRS"
            End If

            Dim rsetPayMethod As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsetPayMethod.DoQuery(strsql)
            Return rsetPayMethod.Fields.Item("HRS").Value
        Catch ex As Exception
            oApplication.StatusBar.SetText("Add Date Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function
    Function isDuplicate(ByVal oEditText As SAPbouiCOM.EditText, ByVal strTableName As String, ByVal strFildName As String, ByVal strMessage As String) As Boolean
        Try
            Dim rsetPayMethod As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim blReturnVal As Boolean = False
            Dim strQuery As String
            If oEditText.Value.Equals("") Then

                oApplication.StatusBar.SetText(strMessage & " : Should Not Be left Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If blnIsHANA Then
                strQuery = "SELECT * FROM """ & strTableName & """ WHERE UPPER(""" & strFildName & """)=UPPER('" & oEditText.Value & "')"
            Else
                strQuery = "SELECT * FROM " & strTableName & " WHERE UPPER(" & strFildName & ")=UPPER('" & oEditText.Value & "')"
            End If

            rsetPayMethod.DoQuery(strQuery)

            If rsetPayMethod.RecordCount > 0 Then
                oEditText.Active = True
                oApplication.StatusBar.SetText(strMessage & " [ " & oEditText.Value & " ] : Already Exist in Table...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True

        Catch ex As Exception
            oApplication.StatusBar.SetText(" isDuplicate Function Failed : ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

#Region " ...  Set ComboBox Null ... "

    Function isMultiDuplicate(ByVal strTableName As String, ByVal FldNa1 As String, ByVal FldVa1 As String, _
                                 Optional ByVal FldNa2 As String = "", Optional ByVal FldVa2 As String = "", _
                                 Optional ByVal FldNa3 As String = "", Optional ByVal FldVa3 As String = "", _
                                 Optional ByVal FldNa4 As String = "", Optional ByVal FldVa4 As String = "", _
                                 Optional ByVal FldNa5 As String = "", Optional ByVal FldVa5 As String = "") As Boolean
        Try
            Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQuery As String
            If blnIsHANA Then
                strQuery = "SELECT COUNT(*) FROM """ & strTableName & """ WHERE """ & FldNa1 & """='" & FldVa1 & "' "
                strQuery = strQuery + IIf(FldNa2 = "", "", " AND """ & FldNa2 & """='" & FldVa2 & "' ")
                strQuery = strQuery + IIf(FldNa3 = "", "", " AND """ & FldNa3 & """='" & FldVa3 & "' ")
                strQuery = strQuery + IIf(FldNa4 = "", "", " AND """ & FldNa4 & """='" & FldVa4 & "' ")
                strQuery = strQuery + IIf(FldNa5 = "", "", " AND """ & FldNa5 & """='" & FldVa5 & "' ")

            Else
                strQuery = "SELECT COUNT(*) FROM " & strTableName & " WHERE " & FldNa1 & "='" & FldVa1 & "' "
                strQuery = strQuery + IIf(FldNa2 = "", "", " AND " & FldNa2 & "='" & FldVa2 & "' ")
                strQuery = strQuery + IIf(FldNa3 = "", "", " AND " & FldNa3 & "='" & FldVa3 & "' ")
                strQuery = strQuery + IIf(FldNa4 = "", "", " AND " & FldNa4 & "='" & FldVa4 & "' ")
                strQuery = strQuery + IIf(FldNa5 = "", "", " AND " & FldNa5 & "='" & FldVa5 & "' ")

            End If
            rset.DoQuery(strQuery)
            rset.MoveFirst()
            Return IIf(rset.Fields.Item(0).Value = "0", False, True)

        Catch ex As Exception
            oApplication.StatusBar.SetText(" isDuplicate Function Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function
#End Region

    Function isYear(ByVal oEditText As SAPbouiCOM.EditText) As Boolean
        Try
            Dim intYear As Integer = 0
            If oEditText.Value.Equals("") = False Then
                Try
                    intYear = CInt(oEditText.Value)
                Catch ex As Exception
                    oApplication.StatusBar.SetText(" Invalid Year Value ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End Try
                ' Check the Year Length ...
                If intYear.ToString().Length <> 4 Then
                    oApplication.StatusBar.SetText(" Invalid Year Value ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("LoadXML Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function

    Function isValiedPrecentage(ByVal dblValue As Double) As Boolean
        Try
            If dblValue > 100 Then
                oApplication.StatusBar.SetText(" Invalid Percentage Value ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("LoadXML Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function

    Function isValiedQuantity(ByVal oEditText As SAPbouiCOM.EditText, ByVal ErrFieldDesc As String) As Boolean
        Try
            If oEditText.Value.Equals("") Then
                oGFun.StatusBarErrorMsg(ErrFieldDesc & " should not be left empty")
                Return False
            ElseIf CDbl(oEditText.Value) <= 0 Then
                oGFun.StatusBarErrorMsg(ErrFieldDesc & " must be greater than zero [message 131-12]")
                Return False
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("is Valied Quantity Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function

    Private Function String_Parser(ByVal string_veri As String, ByVal ayrac As String)
        Dim parsed_strings(0) As String
        Do
            Dim ayrac_yeri As Integer 'HER DEFASINDA YENİ DEĞER ALACAK, DÖNGÜ İÇİNDE TANIMLIYORUM
            ayrac_yeri = Microsoft.VisualBasic.InStr(string_veri, ayrac, CompareMethod.Text) 'VERİ İÇİNDE AYRAÇIN YERİ SAPTANIYOR
            If ayrac_yeri = 0 Then
                parsed_strings(parsed_strings.Length - 1) = string_veri 'AYRAÇ HİÇ YOKSA VERİNİN TAMAMI DİZİYE ATILIYOR
                string_veri = ""
            Else
                ReDim Preserve parsed_strings(parsed_strings.Length) 'AYRAÇ VARSA, DİZİ İÇİNDE YENİ BİR ALAN AÇILIYOR, ESKİ VERİLER SAKLANIYOR
                parsed_strings(parsed_strings.Length - 2) = Microsoft.VisualBasic.Left(string_veri, ayrac_yeri - 1) 'AYRAÇ YERİNE KADAR OLAN KISIM DİZİDE AÇILAN ALANA ALINIYOR
                string_veri = Microsoft.VisualBasic.Right(string_veri, CInt(string_veri.Length) - ayrac_yeri) ' ALINAN KISIM STRING İÇİNDEN SİLİNİYOR
            End If
        Loop Until (string_veri = "") 'STRİNG İÇİNDE OLAN VERİ BİTENE KADAR DÖNÜYOR
        Return parsed_strings
    End Function

    Function DateParse(ByVal veriler As String) As String '30/09/14 şeklinde gelen tarih bilgisini 20140930 şeklinde geri gönderir

        'Dim e_postalar() As String
        'Dim e_postalar() As String = String_Parser(veriler, strTarihAyirac)
        If InStr(veriler, ".") > 0 Then
            Dim e_postalar() As String = String_Parser(veriler, ".")
            Dim gun As String = e_postalar(0)
            If gun.Length = 1 Then gun = "0" + gun
            Dim ay As String = e_postalar(1)
            If ay.Length = 1 Then ay = "0" + ay
            Return "20" + e_postalar(2).Trim + ay.Trim + gun.Trim
        ElseIf InStr(veriler, "/") > 0 Then
            Dim e_postalar() As String = String_Parser(veriler, "/")
            Dim gun As String = e_postalar(0)
            If gun.Length = 1 Then gun = "0" + gun
            Dim ay As String = e_postalar(1)
            If ay.Length = 1 Then ay = "0" + ay
            Return "20" + e_postalar(2).Trim + ay.Trim + gun.Trim
        ElseIf InStr(veriler, "-") > 0 Then
            Dim e_postalar() As String = String_Parser(veriler, "-")
            Dim gun As String = e_postalar(0)
            If gun.Length = 1 Then gun = "0" + gun
            Dim ay As String = e_postalar(1)
            If ay.Length = 1 Then ay = "0" + ay
            Return "20" + e_postalar(2).Trim + ay.Trim + gun.Trim
        End If



    End Function
    Function DateParse2(ByVal veriler As String) As String '30/09/2014 şeklinde gelen tarih bilgisini 20140930 şeklinde geri gönderir

        'Dim e_postalar() As String
        'Dim e_postalar() As String = String_Parser(veriler, strTarihAyirac)
        If InStr(veriler, ".") > 0 Then
            Dim e_postalar() As String = String_Parser(veriler, ".")
            Dim gun As String = e_postalar(0)
            If gun.Length = 1 Then gun = "0" + gun
            Dim ay As String = e_postalar(1)
            If ay.Length = 1 Then ay = "0" + ay
            Return e_postalar(2).Trim + ay.Trim + gun.Trim
        ElseIf InStr(veriler, "/") > 0 Then
            Dim e_postalar() As String = String_Parser(veriler, "/")
            Dim gun As String = e_postalar(0)
            If gun.Length = 1 Then gun = "0" + gun
            Dim ay As String = e_postalar(1)
            If ay.Length = 1 Then ay = "0" + ay
            Return e_postalar(2).Trim + ay.Trim + gun.Trim
        ElseIf InStr(veriler, "-") > 0 Then
            Dim e_postalar() As String = String_Parser(veriler, "-")
            Dim gun As String = e_postalar(0)
            If gun.Length = 1 Then gun = "0" + gun
            Dim ay As String = e_postalar(1)
            If ay.Length = 1 Then ay = "0" + ay
            Return e_postalar(2).Trim + ay.Trim + gun.Trim
        End If
    End Function

    Function DateParseUzun(ByVal veriler As String) As String '17/10/2014 00:00:00 şeklinde gelen tarih bilgisini 20141017 şeklinde geri gönderir
        'veriler = Microsoft.VisualBasic.Left(veriler, 10)
        Dim e_veri() As String = String_Parser(veriler, " ")
        veriler = e_veri(0)

        If InStr(veriler, ".") > 0 Then
            Dim e_postalar() As String = String_Parser(veriler, ".")
            Dim gun As String = e_postalar(0)
            If gun.Length = 1 Then gun = "0" + gun
            Dim ay As String = e_postalar(1)
            If ay.Length = 1 Then ay = "0" + ay
            Return e_postalar(2).Trim + ay.Trim + gun.Trim
        ElseIf InStr(veriler, "/") > 0 Then
            Dim e_postalar() As String = String_Parser(veriler, "/")
            Dim gun As String = e_postalar(0)
            If gun.Length = 1 Then gun = "0" + gun
            Dim ay As String = e_postalar(1)
            If ay.Length = 1 Then ay = "0" + ay
            Return e_postalar(2).Trim + ay.Trim + gun.Trim
        ElseIf InStr(veriler, "-") > 0 Then
            Dim e_postalar() As String = String_Parser(veriler, "-")
            Dim gun As String = e_postalar(0)
            If gun.Length = 1 Then gun = "0" + gun
            Dim ay As String = e_postalar(1)
            If ay.Length = 1 Then ay = "0" + ay
            Return e_postalar(2).Trim + ay.Trim + gun.Trim

        End If


    End Function

    Function DateParseUzun2(ByVal veriler As String) As String '17/10/2014 00:00:00 şeklinde gelen tarih bilgisini 17/10/2014 şeklinde geri gönderir
        'veriler = Microsoft.VisualBasic.Left(veriler, 10)
        'Dim e_postalar() As String = String_Parser(veriler, "/")'strTarihAyirac
        'Dim e_postalar() As String = String_Parser(veriler, strTarihAyirac) '
        Dim e_veri() As String = String_Parser(veriler, " ")
        veriler = e_veri(0)

        If InStr(veriler, ".") > 0 Then
            Dim e_postalar() As String = String_Parser(veriler, ".")
            Dim gun As String = e_postalar(0)
            If gun.Length = 1 Then gun = "0" + gun
            Dim ay As String = e_postalar(1)
            If ay.Length = 1 Then ay = "0" + ay
            Return gun.Trim + strTarihAyirac + ay.Trim + strTarihAyirac + e_postalar(2).Trim
        ElseIf InStr(veriler, "/") > 0 Then
            Dim e_postalar() As String = String_Parser(veriler, "/")
            Dim gun As String = e_postalar(0)
            If gun.Length = 1 Then gun = "0" + gun
            Dim ay As String = e_postalar(1)
            If ay.Length = 1 Then ay = "0" + ay
            Return gun.Trim + strTarihAyirac + ay.Trim + strTarihAyirac + e_postalar(2).Trim
        ElseIf InStr(veriler, "-") > 0 Then
            Dim e_postalar() As String = String_Parser(veriler, "-")
            Dim gun As String = e_postalar(0)
            If gun.Length = 1 Then gun = "0" + gun
            Dim ay As String = e_postalar(1)
            If ay.Length = 1 Then ay = "0" + ay
            Return gun.Trim + strTarihAyirac + ay.Trim + strTarihAyirac + e_postalar(2).Trim
        End If


    End Function

    Function isDateCompare(ByVal oEditFromDate As SAPbouiCOM.EditText, ByVal oEditToDate As SAPbouiCOM.EditText, ByVal ErrorMsg As String) As Boolean
        Try
            If oEditFromDate.Value.Equals("") = False And oEditToDate.Value.Equals("") = False Then
                Dim dtFromDate As Date = DateTime.ParseExact(oEditFromDate.Value, "yyyyMMdd", Nothing)
                Dim dtToDate As Date = DateTime.ParseExact(oEditToDate.Value, "yyyyMMdd", Nothing)
                If dtFromDate > dtToDate Then
                    oApplication.StatusBar.SetText(ErrorMsg & " ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("DateValidate Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function

    Function AddDate(ByVal SDATE As String, ByVal Days As Integer) As String
        Try
            Dim strsql
            If blnIsHANA Then
                strsql = "SELECT CAST('" & SDATE & "' AS timestamp) || " & Days & " AS ""DATE"" FROM DUMMY"
            Else
                strsql = "SELECT CONVERT(DATETIME,'" & SDATE & "')+" & Days & " DATE"
            End If
            Dim rsetPayMethod As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsetPayMethod.DoQuery(strsql)
            Return CDate(rsetPayMethod.Fields.Item("DATE").Value).ToString("yyyyMMdd")
        Catch ex As Exception
            oApplication.StatusBar.SetText("Add Date Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function


    Function ComputeDateDAYDiff(ByVal SDATE As String, ByVal SDATE2 As String) As Integer
        Try

            Dim strsql
            If blnIsHANA Then
                strsql = "SELECT DAYS_BETWEEN(CAST('" & SDATE & "' AS timestamp), CAST('" & SDATE2 & "' AS timestamp)) AS ""GUNFARK"" FROM DUMMY"
            Else
                strsql = "SELECT DATEDIFF(d,CONVERT(DATETIME,'" & SDATE & "'),CONVERT(DATETIME,'" & SDATE2 & "')) GUNFARK"
            End If
            Dim rsetPayMethod As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsetPayMethod.DoQuery(strsql)
            Return CInt(rsetPayMethod.Fields.Item("GUNFARK").Value)
        Catch ex As Exception
            oApplication.StatusBar.SetText("Datediff Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function

    Function isVaiedTimeHHSSMM(ByVal oEdit As SAPbouiCOM.EditText) As Boolean
        Try
            Dim StrTime As String = oEdit.Value
            If StrTime.Trim.Equals("") Then Return True
            Dim Hours = 0, Minutes = 0, Seconds As Double = 0
            Dim dblTime As Double = 0
            Try
                dblTime = CDbl(StrTime.Replace(":", ""))
            Catch ex As Exception
                oGFun.StatusBarErrorMsg(" In-Valied time value ")
                Return False
            End Try
            Dim FormatedTime = dblTime.ToString("00:00:00")
            Dim strArray As String() = FormatedTime.Split(":")
            Try

                If strArray.Length > 0 Then
                    If CDbl(strArray(0)) < 0 Then
                        oGFun.StatusBarErrorMsg(strArray(0) & " is not a valid hours")
                        Return False
                    ElseIf CDbl(strArray(1)) > 59 Or CDbl(strArray(1)) < 0 Then
                        oGFun.StatusBarErrorMsg(strArray(1) & " is not a valid minutes")
                        Return False
                    ElseIf CDbl(strArray(2)) > 59 Or CDbl(strArray(1)) < 0 Then
                        oGFun.StatusBarErrorMsg(strArray(2) & " is not a valid seconds")
                        Return False
                    End If
                End If
                oEdit.Value = FormatedTime
                Return True
            Catch ex As Exception
                oGFun.StatusBarErrorMsg("is not a valid time")
                Return False
            Finally
            End Try
        Catch ex As Exception

        End Try
    End Function

    Function LoadComboBoxSeries(ByVal oComboBox As SAPbouiCOM.ComboBox, ByVal UDOID As String) As Boolean
        Try
            oComboBox.ValidValues.LoadSeries(UDOID, SAPbouiCOM.BoSeriesMode.sf_Add)
            oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        Catch ex As Exception
            oApplication.StatusBar.SetText("LoadComboBoxSeries Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try

    End Function

    Function LoadDocumentDate(ByVal oEditText As SAPbouiCOM.EditText) As Boolean
        Try
            oEditText.Active = True
            oEditText.String = "A"
        Catch ex As Exception
            oApplication.StatusBar.SetText("LoadDocumentDate Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try

    End Function

    Sub SetComboBoxValueRefresh(ByVal oComboBox As SAPbouiCOM.ComboBox, ByVal strQry As String)
        Try
            Dim rsetValidValue As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim intCount As Integer = oComboBox.ValidValues.Count
            ' Remove the Combo Box Value Based On Count ...
            If intCount > 0 Then
                While intCount > 0
                    oComboBox.ValidValues.Remove(intCount - 1, SAPbouiCOM.BoSearchKey.psk_Index)
                    intCount = intCount - 1
                End While
            End If

            rsetValidValue.DoQuery(strQry)
            rsetValidValue.MoveFirst()
            For j As Integer = 0 To rsetValidValue.RecordCount - 1
                oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value)
                rsetValidValue.MoveNext()
            Next


        Catch ex As Exception
            oGFun.StatusBarErrorMsg("SetComboBoxValueRefresh Method Faild : " & ex.Message)
        Finally
        End Try
    End Sub


    Function setComboBoxValueBosDegerYok(ByVal oComboBox As SAPbouiCOM.ComboBox, ByVal strQry As String) As Boolean
        Try
            If oComboBox.ValidValues.Count = 0 Then
                Dim rsetValidValue As SAPbobsCOM.Recordset = oGFun.DoQuery(strQry)
                rsetValidValue.MoveFirst()
                For j As Integer = 0 To rsetValidValue.RecordCount - 1
                    oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value)
                    rsetValidValue.MoveNext()
                Next

            End If
            ' If oComboBox.ValidValues.Count > 0 Then oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        Catch ex As Exception
            oApplication.StatusBar.SetText("setComboBoxValue Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        Finally
        End Try

    End Function

    Function setComboBoxValue(ByVal oComboBox As SAPbouiCOM.ComboBox, ByVal strQry As String) As Boolean
        Try
            If oComboBox.ValidValues.Count = 0 Then
                Dim rsetValidValue As SAPbobsCOM.Recordset = oGFun.DoQuery(strQry)
                rsetValidValue.MoveFirst()
                For j As Integer = 0 To rsetValidValue.RecordCount - 1
                    oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value)
                    rsetValidValue.MoveNext()
                Next
                If oComboBox.ValidValues.Count > 0 Then
                    oComboBox.ValidValues.Add("", "Boş")
                End If
            End If
            ' If oComboBox.ValidValues.Count > 0 Then oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        Catch ex As Exception
            oApplication.StatusBar.SetText("setComboBoxValue Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        Finally
        End Try

    End Function

    'oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value)
    Function addComboBoxBosSatir(ByVal oComboBox As SAPbouiCOM.ComboBox) As Boolean
        Try
            If oComboBox.ValidValues.Count > 0 Then
                oComboBox.ValidValues.Add("", "Boş")
            End If
            ' If oComboBox.ValidValues.Count > 0 Then oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        Catch ex As Exception
            oApplication.StatusBar.SetText("setComboBoxValue Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        Finally
        End Try

    End Function

    Function setComboBoxValueUclu(ByVal oComboBox As SAPbouiCOM.ComboBox, ByVal strQry As String) As Boolean
        Try
            If oComboBox.ValidValues.Count = 0 Then
                Dim rsetValidValue As SAPbobsCOM.Recordset = oGFun.DoQuery(strQry)
                rsetValidValue.MoveFirst()
                For j As Integer = 0 To rsetValidValue.RecordCount - 1
                    oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value.ToString() + " - " + rsetValidValue.Fields.Item(2).Value)
                    rsetValidValue.MoveNext()
                Next
            End If
            ' If oComboBox.ValidValues.Count > 0 Then oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        Catch ex As Exception
            oApplication.StatusBar.SetText("setComboBoxValue Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        Finally
        End Try

    End Function

    'Combodan Seçli Değere Ait Satırından Herhangi Bir Veriyi Bir Başka Alana Taşımak İçin ---SEVDA
    Function setComboBoxValueANDReturnRecordSet(ByVal oComboBox As SAPbouiCOM.ComboBox, ByVal strQry As String) As SAPbobsCOM.Recordset
        Try
            Dim rsetValidValue As SAPbobsCOM.Recordset = oGFun.DoQuery(strQry)
            If oComboBox.ValidValues.Count = 0 Then

                rsetValidValue.MoveFirst()
                For j As Integer = 0 To rsetValidValue.RecordCount - 1
                    oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value)
                    rsetValidValue.MoveNext()
                Next
            End If

            Return rsetValidValue
        Catch ex As Exception
            oApplication.StatusBar.SetText("setComboBoxValue Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return Nothing
        Finally
        End Try

    End Function

    Function ReturnRecordSet(ByVal strQry As String) As SAPbobsCOM.Recordset
        Try
            Dim rsetValidValue As SAPbobsCOM.Recordset = oGFun.DoQuery(strQry)

            Return rsetValidValue
        Catch ex As Exception
            oApplication.StatusBar.SetText("ReturnRecordSet Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return Nothing
        Finally
        End Try

    End Function

    Function setComboBoxValueUcluANDReturnRecordSet(ByVal oComboBox As SAPbouiCOM.ComboBox, ByVal strQry As String) As SAPbobsCOM.Recordset
        Try
            Dim rsetValidValue As SAPbobsCOM.Recordset = oGFun.DoQuery(strQry)
            If oComboBox.ValidValues.Count = 0 Then

                rsetValidValue.MoveFirst()
                For j As Integer = 0 To rsetValidValue.RecordCount - 1
                    oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value.ToString() + " - " + rsetValidValue.Fields.Item(2).Value)
                    rsetValidValue.MoveNext()
                Next
            End If

            Return rsetValidValue
        Catch ex As Exception
            oApplication.StatusBar.SetText("setComboBoxValue Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return Nothing
        Finally
        End Try

    End Function

    Function setdocvalue(ByVal odoctextbox As SAPbouiCOM.EditText, ByVal strQry As String) As Boolean
        Try
            Dim rsetValue As SAPbobsCOM.Recordset = oGFun.DoQuery(strQry)
            odoctextbox.Active = True
            odoctextbox.Value = rsetValue.Fields.Item(0).Value
            odoctextbox.Active = False
        Catch ex As Exception
            oApplication.StatusBar.SetText("setDocnumber Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        End Try
    End Function

    Sub LoadDepartmentComboBox(ByVal oComboBox As SAPbouiCOM.ComboBox)
        Try
            Dim strQry As String

            If oComboBox.ValidValues.Count = 0 Then
                Dim rsetValidValue As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If blnIsHANA Then
                    strQry = "SELECT ""Code"", ""Name"" FROM OUDP"
                Else
                    strQry = "SELECT Code , Name FROM OUDP"
                End If

                rsetValidValue.DoQuery(strQry)
                rsetValidValue.MoveFirst()
                For j As Integer = 0 To rsetValidValue.RecordCount - 1
                    oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value)
                    rsetValidValue.MoveNext()
                Next
            End If
        Catch ex As Exception
            oGFun.StatusBarErrorMsg("SetComboBoxValueRefresh Method Faild : " & ex.Message)
        Finally
        End Try
    End Sub

    Sub LoadLocationComboBox(ByVal oComboBox As SAPbouiCOM.ComboBox)
        Try
            If oComboBox.ValidValues.Count = 0 Then
                Dim rsetValidValue As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim strQry As String
                If blnIsHANA Then
                    strQry = "SELECT ""Code"", ""Location"" FROM OLCT ORDER BY CAST(""Code"" AS integer)"
                Else
                    strQry = "SELECT Code , Location FROM OLCT ORDER BY CONVERT(INT ,Code) "
                End If


                rsetValidValue.DoQuery(strQry)
                rsetValidValue.MoveFirst()
                For j As Integer = 0 To rsetValidValue.RecordCount - 1
                    oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value)
                    rsetValidValue.MoveNext()
                Next
            End If

        Catch ex As Exception
            oGFun.StatusBarErrorMsg("SetComboBoxValueRefresh Method Faild : " & ex.Message)
        Finally
        End Try
    End Sub

    Function GetCodeGeneration(ByVal TableName As String) As Integer
        Try
            Dim rsetCode As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strCode As String
            If blnIsHANA Then
                strCode = "SELECT IFNULL(MAX(IFNULL(""DocEntry"", 0)), 0) + 1 AS ""Code"" FROM """ & Trim(TableName) & """"
            Else
                strCode = "Select ISNULL(Max(ISNULL(DocEntry,0)),0) + 1 Code From " & Trim(TableName) & ""
            End If

            rsetCode.DoQuery(strCode)
            'MessageBox.Show(CInt(rsetCode.Fields.Item("Code").Value))
            Return CInt(rsetCode.Fields.Item("Code").Value)
        Catch ex As Exception
            oApplication.StatusBar.SetText("GetCodeGeneration Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        Finally
        End Try
    End Function

    Function GetProsIdGeneration() As String
        Try
            Dim rsetCode As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strCode As String
            If blnIsHANA Then
                strCode = "SELECT ""ItemCode"" FROM OITM WHERE ""DocEntry"" = (SELECT MAX(""DocEntry"") FROM OITM)"
            Else
                strCode = "select ITEMCODE from oitm where docentry =(select max(docentry) from oitm)"
            End If

            rsetCode.DoQuery(strCode)
            Return rsetCode.Fields.Item("ITEMCODE").Value
        Catch ex As Exception
            oApplication.StatusBar.SetText("GetCodeGeneration Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        Finally
        End Try
    End Function

    Sub SetNewLine(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSDetail As SAPbouiCOM.DBDataSource, Optional ByVal RowID As Integer = 1, Optional ByVal ColumnUID As String = "")
        Try
            If ColumnUID.Equals("") = False Then
                If oMatrix.VisualRowCount > 0 Then
                    If oMatrix.Columns.Item(ColumnUID).Cells.Item(RowID).Specific.Value.Equals("") = False And RowID = oMatrix.VisualRowCount Then
                        oMatrix.FlushToDataSource()
                        oMatrix.AddRow()
                        oDBDSDetail.InsertRecord(oDBDSDetail.Size)
                        oDBDSDetail.Offset = oMatrix.VisualRowCount - 1
                        oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, oMatrix.VisualRowCount)

                        oMatrix.SetLineData(oMatrix.VisualRowCount)
                        oMatrix.FlushToDataSource()
                    End If
                Else
                    oMatrix.FlushToDataSource()
                    oMatrix.AddRow()
                    oDBDSDetail.InsertRecord(oDBDSDetail.Size)
                    oDBDSDetail.Offset = oMatrix.VisualRowCount - 1
                    oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, oMatrix.VisualRowCount)

                    oMatrix.SetLineData(oMatrix.VisualRowCount)
                    oMatrix.FlushToDataSource()
                End If

            Else
                oMatrix.FlushToDataSource()
                oMatrix.AddRow()
                oDBDSDetail.InsertRecord(oDBDSDetail.Size)
                oDBDSDetail.Offset = oMatrix.VisualRowCount - 1
                oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, oMatrix.VisualRowCount)

                oMatrix.SetLineData(oMatrix.VisualRowCount)
                oMatrix.FlushToDataSource()
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("SetNewLine Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub DeleteRow(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSDetail As SAPbouiCOM.DBDataSource)
        Try
            oMatrix.FlushToDataSource()

            For i As Integer = 1 To oMatrix.VisualRowCount
                oMatrix.GetLineData(i)
                oDBDSDetail.Offset = i - 1
                oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, i)
                oMatrix.SetLineData(i)
                oMatrix.FlushToDataSource()
            Next
            oDBDSDetail.RemoveRecord(oDBDSDetail.Size - 1)
            oMatrix.LoadFromDataSource()

        Catch ex As Exception
            oApplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Function isBatchNoApplicable(ByVal ItemCode As String) As Boolean
        Try
            Dim strsql
            If blnIsHANA Then
                strsql = "select ""ManBtchNum"" from OITM where ""ItemCode""='" & ItemCode & "'"
            Else
                strsql = "select ManBtchNum from oitm where Itemcode='" & ItemCode & "'"
            End If

            Dim RS As SAPbobsCOM.Recordset
            RS = oGFun.DoQuery(strsql)
            If RS.RecordCount > 0 Then
                If RS.Fields.Item(0).Value.ToString = "Y" Then Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
            oApplication.StatusBar.SetText("isBatchNoApplicable Fun. Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Function

    Sub ChooseFromListFilteration(ByVal oForm As SAPbouiCOM.Form, ByVal strCFL_ID As String, ByVal strCFL_Alies As String, ByVal strQuery As String)
        Try
            Dim oCFL As SAPbouiCOM.ChooseFromList = oForm.ChooseFromLists.Item(strCFL_ID)
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim rsetCFL As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()

            rsetCFL.DoQuery(strQuery)
            rsetCFL.MoveFirst()
            For i As Integer = 1 To rsetCFL.RecordCount
                If i = (rsetCFL.RecordCount) Then
                    oCond = oConds.Add()
                    oCond.Alias = strCFL_Alies
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                Else
                    oCond = oConds.Add()
                    oCond.Alias = strCFL_Alies
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                rsetCFL.MoveNext()
            Next
            If rsetCFL.RecordCount = 0 Then
                oCond = oConds.Add()
                oCond.Alias = strCFL_Alies
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE
                oCond.CondVal = "-1"
            End If
            oCFL.SetConditions(oConds)
        Catch ex As Exception
            oApplication.StatusBar.SetText("Choose FromList Filter Global Fun. Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub DoOpenLinkedObjectForm(ByVal FormUniqueID As String, ByVal ActivateMenuItem As String, ByVal FindItemUID As String, ByVal FindItemUIDValue As String)
        Try
            Dim oForm As SAPbouiCOM.Form
            Dim Bool As Boolean = False

            For frm As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(frm).UniqueID = FormUniqueID Then
                    oForm = oApplication.Forms.Item(FormUniqueID)
                    oForm.Close()
                    Exit For
                End If
            Next
            If Bool = False Then
                oApplication.ActivateMenuItem(ActivateMenuItem)
                oForm = oApplication.Forms.Item(ActivateMenuItem)
                oForm.Select()
                oForm.Freeze(True)
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                oForm.Items.Item(FindItemUID).Enabled = True
                oForm.Items.Item(FindItemUID).Specific.Value = Trim(FindItemUIDValue)
                oForm.Items.Item("1").Click()
                oForm.Freeze(False)
            End If
        Catch ex As Exception
            oGFun.StatusBarErrorMsg("" & ex.Message)
        Finally
        End Try
    End Sub

    Sub DeleteEmptyRowInFormDataEvent(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal ColumnUID As String, ByVal oDBDSDetail As SAPbouiCOM.DBDataSource)
        Try
            If oMatrix.VisualRowCount > 0 Then
                If oMatrix.Columns.Item(ColumnUID).Cells.Item(oMatrix.VisualRowCount).Specific.Value.Equals("") Then
                    oMatrix.DeleteRow(oMatrix.VisualRowCount)
                    oDBDSDetail.RemoveRecord(oDBDSDetail.Size - 1)
                    oMatrix.FlushToDataSource()
                End If
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("Delete Empty RowIn Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub SubMenuAddEx(ByVal strMenuUID As String, ByVal strMenuName As String)
        Try
            Dim MenuItem As SAPbouiCOM.MenuItem = oApplication.Menus.Item("1280") 'Data'
            Dim Menu As SAPbouiCOM.Menus = MenuItem.SubMenus
            Dim MenuParam As SAPbouiCOM.MenuCreationParams = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

            MenuParam.Type = SAPbouiCOM.BoMenuType.mt_STRING
            MenuParam.UniqueID = strMenuUID
            MenuParam.String = strMenuName
            MenuParam.Enabled = True
            If MenuItem.SubMenus.Exists(strMenuUID) = False Then Menu.AddEx(MenuParam)

        Catch ex As Exception
            oApplication.StatusBar.SetText("SubMenuAddEx Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub SubMenusRemoveEx(ByVal strMenuID As String)
        Try

            If oApplication.Menus.Item("1280").SubMenus.Exists(strMenuID) Then oApplication.Menus.Item("1280").SubMenus.RemoveEx(strMenuID)

        Catch ex As Exception
            oApplication.StatusBar.SetText("SubMenusRemoveEx Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub StatusBarErrorMsg(ByVal ErrorMsg As String)
        Try
            oApplication.StatusBar.SetText(ErrorMsg, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Catch ex As Exception
            oApplication.StatusBar.SetText("StatusBar ErrorMsg Method Failed" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub StatusBarErrorMsgShort(ByVal ErrorMsg As String)
        Try
            oApplication.StatusBar.SetText(ErrorMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Catch ex As Exception
            oApplication.StatusBar.SetText("StatusBar ErrorMsg Method Failed" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub StatusBarWarningMsg(ByVal WarningMsg As String)
        Try
            oApplication.StatusBar.SetText(WarningMsg, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Catch ex As Exception
            oApplication.StatusBar.SetText("StatusBar WarningMsg Method Failed" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Function getSingleValue(ByVal TblName As String, ByVal ValFldNa As String, ByVal Conditions As String) As String
        Try
            Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strReturnVal As String = ""

            Dim strQuery
            If blnIsHANA Then
                strQuery = "SELECT """ & ValFldNa & """ FROM """ & TblName & """ IIf(" & Conditions.Trim() & " = "", "",  WHERE )" & Conditions
            Else
                strQuery = "SELECT " & ValFldNa & " FROM " & TblName & IIf(Conditions.Trim() = "", "", " WHERE ") & Conditions
            End If

            rset.DoQuery(strQuery)
            Return IIf(rset.RecordCount > 0, rset.Fields.Item(0).Value.ToString(), "")
        Catch ex As Exception
            oApplication.StatusBar.SetText(" Get Single Value Function Failed : ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return ""
        End Try
    End Function

    Function isValidFrAndToDate(ByVal FrDate As String, ByVal ToDate As String) As Boolean
        Try
            Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim strQuery
            If blnIsHANA Then
                strQuery = "SELECT CASE WHEN CAST('" & FrDate & "' AS timestamp) <= CAST('" & ToDate & "' AS timestamp) THEN 'True' Else 'False' End FROM DUMMY"
            Else
                strQuery = "Select case when convert(datetime,'" & FrDate & "') <= convert(datetime,'" & ToDate & "') Then 'True' else 'False' End"
            End If

            rset.DoQuery(strQuery)
            Return Convert.ToBoolean(rset.Fields.Item(0).Value)
        Catch ex As Exception
            oApplication.StatusBar.SetText(" IS valid From Date and To Date Function Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Function isValidFrAndToTime(ByVal FrTime As String, ByVal ToTime As String) As Boolean
        Try
            Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQuery
            If blnIsHANA Then
                strQuery = "SELECT CASE WHEN CAST('" & FrTime & "' AS timestamp) <= CAST('" & ToTime & "' AS timestamp) THEN 'True' Else 'False' End FROM DUMMY"
            Else
                strQuery = "Select case when convert(TimeStamp,'" & FrTime & "') <= convert(TimeStamp,'" & ToTime & "') Then 'True' else 'False' End"
            End If

            rset.DoQuery(strQuery)
            Return Convert.ToBoolean(rset.Fields.Item(0).Value)
        Catch ex As Exception
            oApplication.StatusBar.SetText(" IS valid From Date and To Date Function Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Function GetMinValueArray(ByVal Array() As String) As Integer
        Try
            Dim MinValue As Integer = Array(0)
            For i As Integer = 1 To Array.Length - 1
                If Not Array(i) Is Nothing Then
                    If Array(i) < MinValue Then MinValue = Array(i)
                End If
            Next
            Return MinValue
        Catch ex As Exception
            oGFun.StatusBarErrorMsg("Get MinValue Array : " & ex.Message)
        Finally
        End Try
    End Function

    Function GetWareHouseStock(ByVal sItemCode As String, ByVal sWareHouse As String) As Double
        Try
            Dim sQuery As String
            If blnIsHANA Then
                sQuery = "SELECT w.""OnHand"" FROM OITM i, OITW w WHERE i.""ItemCode"" = w.""ItemCode"" AND i.""ItemCode"" = '" & sItemCode.Trim & "' AND w.""WhsCode"" = '" & sWareHouse.Trim & "'"
            Else
                sQuery = "SELECT w.OnHand FROM OITM i ,OITW w WHERE i.ItemCode = w.ItemCode AND i.ItemCode = '" & sItemCode.Trim & "' AND w.WhsCode = '" & sWareHouse.Trim & "'"
            End If

            Dim rsetOnHand As SAPbobsCOM.Recordset = oGFun.DoQuery(sQuery)
            If rsetOnHand.RecordCount > 0 Then Return CDbl(rsetOnHand.Fields.Item("OnHand").Value)
            Return 0
        Catch ex As Exception
            oGFun.StatusBarErrorMsg("Get MinValue Array : " & ex.Message)
            Return 0
        Finally
        End Try
    End Function

#End Region

#Region "       Attachment Functions     "

    Public Sub ShowFolderBrowser()
        Dim MyProcs() As System.Diagnostics.Process
        BankFileName = ""
        Dim OpenFile As New OpenFileDialog
        Try
            OpenFile.Multiselect = False
            OpenFile.Filter = "All files(*.)|*.*" '   "|*.*"
            Dim filterindex As Integer = 0
            Try
                filterindex = 0
            Catch ex As Exception
            End Try
            OpenFile.FilterIndex = filterindex
            OpenFile.RestoreDirectory = True
            MyProcs = Process.GetProcessesByName("SAP Business One")
            If MyProcs.Length = 1 Then
                For i As Integer = 0 To MyProcs.Length - 1
                    Dim MyWindow As New WindowWrapper(MyProcs(i).MainWindowHandle)
                    Dim ret As DialogResult = OpenFile.ShowDialog(MyWindow)
                    If ret = DialogResult.OK Then
                        BankFileName = OpenFile.FileName
                        OpenFile.Dispose()
                    Else
                        System.Windows.Forms.Application.ExitThread()
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            BankFileName = ""
        Finally
            OpenFile.Dispose()
        End Try
    End Sub

    Public Function FindFile() As String
        Dim ShowFolderBrowserThread As Threading.Thread
        Try
            ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)
            If ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Unstarted Then
                ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA)
                ShowFolderBrowserThread.Start()
            ElseIf ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Stopped Then
                ShowFolderBrowserThread.Start()
                ShowFolderBrowserThread.Join()
            End If
            While ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Running
                System.Windows.Forms.Application.DoEvents()
            End While
            If BankFileName <> "" Then
                Return BankFileName
            End If
        Catch ex As Exception
            oApplication.MessageBox("FileFile Method Failed : " & ex.Message)
        End Try
        Return ""
    End Function

    Public Sub OpenFile(ByVal ServerPath As String, ByVal ClientPath As String)
        Try
            Dim oProcess As System.Diagnostics.Process = New System.Diagnostics.Process
            Try
                oProcess.StartInfo.FileName = ServerPath
                oProcess.Start()
            Catch ex1 As Exception
                Try
                    oProcess.StartInfo.FileName = ClientPath
                    oProcess.Start()
                Catch ex2 As Exception
                    oApplication.StatusBar.SetText("" & ex2.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Finally
                End Try
            Finally
            End Try
        Catch ex As Exception
            oApplication.StatusBar.SetText("" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Public Class WindowWrapper

        Implements System.Windows.Forms.IWin32Window
        Private _hwnd As IntPtr

        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

    End Class


#Region "          Attachment Option          "

    Sub AddAttachment(ByVal oMatAttach As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource, ByVal oDBDSHeader As SAPbouiCOM.DBDataSource)
        Try
            If oMatAttach.VisualRowCount > 0 Then
                Dim rsetAttCount As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oAttachment As SAPbobsCOM.Attachments2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2)
                Dim oAttchLines As SAPbobsCOM.Attachments2_Lines
                oAttchLines = oAttachment.Lines
                oMatAttach.FlushToDataSource()
                If blnIsHANA Then
                    rsetAttCount.DoQuery("Select Count(*) From ATC1 Where ""AbsEntry"" = '" & Trim(oDBDSHeader.GetValue("U_AtcEntry", 0)) & "'")
                Else
                    rsetAttCount.DoQuery("Select Count(*) From ATC1 Where AbsEntry = '" & Trim(oDBDSHeader.GetValue("U_AtcEntry", 0)) & "'")
                End If
                If Trim(rsetAttCount.Fields.Item(0).Value).Equals("0") Then
                    For i As Integer = 1 To oMatAttach.VisualRowCount
                        If i > 1 Then oAttchLines.Add()
                        oDBDSAttch.Offset = i - 1
                        oAttchLines.SourcePath = Trim(oDBDSAttch.GetValue("U_ScrPath", oDBDSAttch.Offset))
                        oAttchLines.FileName = Trim(oDBDSAttch.GetValue("U_FileName", oDBDSAttch.Offset))
                        oAttchLines.FileExtension = Trim(oDBDSAttch.GetValue("U_FileExt", oDBDSAttch.Offset))
                        oAttchLines.Override = SAPbobsCOM.BoYesNoEnum.tYES
                    Next
                    oAttachment.Add()
                    Dim rsetAttch As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If blnIsHANA Then
                        rsetAttch.DoQuery("SELECT CASE WHEN COUNT(*) > 0 THEN MAX(""AbsEntry"") ELSE 0 END AS ""AbsEntry"" FROM ATC1")
                    Else
                        rsetAttch.DoQuery("Select  Case When Count(*) > 0 Then  Max(AbsEntry) Else 0 End AbsEntry  From ATC1")
                    End If

                    oDBDSHeader.SetValue("U_AtcEntry", 0, rsetAttch.Fields.Item(0).Value)
                Else
                    oAttachment.GetByKey(Trim(oDBDSHeader.GetValue("U_AtcEntry", 0)))
                    For i As Integer = 1 To oMatAttach.VisualRowCount
                        If oAttchLines.Count < i Then oAttchLines.Add()
                        oDBDSAttch.Offset = i - 1
                        oAttchLines.SetCurrentLine(i - 1)
                        oAttchLines.SourcePath = Trim(oDBDSAttch.GetValue("U_ScrPath", oDBDSAttch.Offset))
                        oAttchLines.FileName = Trim(oDBDSAttch.GetValue("U_FileName", oDBDSAttch.Offset))
                        oAttchLines.FileExtension = Trim(oDBDSAttch.GetValue("U_FileExt", oDBDSAttch.Offset))
                        oAttchLines.Override = SAPbobsCOM.BoYesNoEnum.tYES
                    Next
                    oAttachment.Update()
                End If
            End If
            'Delete the Attachment Rows...
            Dim rsetDelete As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If blnIsHANA Then
                rsetDelete.DoQuery("Delete From ATC1 Where ""AbsEntry"" = '" & Trim(oDBDSHeader.GetValue("U_AtcEntry", 0)) & "' And ""Line"" >'" & oMatAttach.VisualRowCount & "' ")
            Else
                rsetDelete.DoQuery("Delete From ATC1 Where AbsEntry = '" & Trim(oDBDSHeader.GetValue("U_AtcEntry", 0)) & "' And Line >'" & oMatAttach.VisualRowCount & "' ")
            End If


        Catch ex As Exception
            oApplication.StatusBar.SetText("AddAttachment Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub DeleteRowAttachment(ByVal oForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource, ByVal SelectedRowID As Integer)
        Try
            oDBDSAttch.RemoveRecord(SelectedRowID - 1)
            oMatrix.DeleteRow(SelectedRowID)
            oMatrix.FlushToDataSource()

            For i As Integer = 1 To oMatrix.VisualRowCount
                oMatrix.GetLineData(i)
                oDBDSAttch.Offset = i - 1

                oDBDSAttch.SetValue("LineID", oDBDSAttch.Offset, i)
                oDBDSAttch.SetValue("U_TrgtPath", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("trgtpath").Cells.Item(i).Specific.Value))
                oDBDSAttch.SetValue("U_ScrPath", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("scrpath").Cells.Item(i).Specific.Value))
                oDBDSAttch.SetValue("U_FileName", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("filename").Cells.Item(i).Specific.Value))
                oDBDSAttch.SetValue("U_FileExt", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("fileext").Cells.Item(i).Specific.Value))
                oDBDSAttch.SetValue("U_Date", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("date").Cells.Item(i).Specific.Value))
                oMatrix.SetLineData(i)
                oMatrix.FlushToDataSource()
            Next
            'oDBDSAttch.RemoveRecord(oDBDSAttch.Size - 1)
            oMatrix.LoadFromDataSource()

            oForm.Items.Item("b_display").Enabled = False
            oForm.Items.Item("b_delete").Enabled = False

            If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

        Catch ex As Exception
            oApplication.StatusBar.SetText("DeleteRowAttachment Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Function SetAttachMentFile(ByVal oForm As SAPbouiCOM.Form, ByVal oDBDSHeader As SAPbouiCOM.DBDataSource, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource) As Boolean
        Try
            If oCompany.AttachMentPath.Length <= 0 Then
                oGFun.StatusBarErrorMsg("Attchment folder not defined, or Attchment folder has been changed or removed. [Message 131-102]")
                Return False
            End If

            Dim strFileName As String = FindFile()
            If strFileName.Equals("") = False Then
                Dim FileExist() As String = strFileName.Split("\")
                Dim FileDestPath As String = oCompany.AttachMentPath & FileExist(FileExist.Length - 1)

                If File.Exists(FileDestPath) Then
                    Dim LngRetVal As Long = oApplication.MessageBox("A file with this name already exists,would you like to replace this?  " & FileDestPath & " will be replaced.", 1, "Yes", "No")
                    If LngRetVal <> 1 Then Return False
                End If
                Dim fileNameExt() As String = FileExist(FileExist.Length - 1).Split(".")
                Dim ScrPath As String = oCompany.AttachMentPath
                ScrPath = ScrPath.Substring(0, ScrPath.Length - 1)
                Dim TrgtPath As String = strFileName.Substring(0, strFileName.LastIndexOf("\"))

                oMatrix.AddRow()
                oMatrix.FlushToDataSource()
                oDBDSAttch.Offset = oDBDSAttch.Size - 1
                oDBDSAttch.SetValue("LineID", oDBDSAttch.Offset, oMatrix.VisualRowCount)
                oDBDSAttch.SetValue("U_TrgtPath", oDBDSAttch.Offset, ScrPath)
                oDBDSAttch.SetValue("U_ScrPath", oDBDSAttch.Offset, TrgtPath)
                oDBDSAttch.SetValue("U_FileName", oDBDSAttch.Offset, fileNameExt(0))
                oDBDSAttch.SetValue("U_FileExt", oDBDSAttch.Offset, fileNameExt(1))
                oDBDSAttch.SetValue("U_Date", oDBDSAttch.Offset, oGFun.GetServerDate())
                oMatrix.SetLineData(oDBDSAttch.Size)
                oMatrix.FlushToDataSource()
                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Set AttachMent File Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function

    Sub OpenAttachment(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource, ByVal PvalRow As Integer)
        Try
            If PvalRow <= oMatrix.VisualRowCount And PvalRow <> 0 Then
                ' Dim RowIndex As Integer = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder) - 1
                Dim strServerPath, strClientPath As String

                strServerPath = Trim(oDBDSAttch.GetValue("U_TrgtPath", PvalRow - 1)) + "\" + Trim(oDBDSAttch.GetValue("U_FileName", PvalRow - 1)) + "." + Trim(oDBDSAttch.GetValue("U_FileExt", PvalRow - 1))
                strClientPath = Trim(oDBDSAttch.GetValue("U_ScrPath", PvalRow - 1)) + "\" + Trim(oDBDSAttch.GetValue("U_FileName", PvalRow - 1)) + "." + Trim(oDBDSAttch.GetValue("U_FileExt", PvalRow - 1))
                'Open Attachment File
                Me.OpenFile(strServerPath, strClientPath)
            End If

        Catch ex As Exception
            oApplication.StatusBar.SetText("OpenAttachment Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub AttchButtonEnable(ByVal oForm As SAPbouiCOM.Form, ByVal Matrix As SAPbouiCOM.Matrix, ByVal PvalRow As Integer)
        Try
            If PvalRow <= Matrix.VisualRowCount And PvalRow <> 0 Then
                Matrix.SelectRow(PvalRow, True, False)
                If Matrix.IsRowSelected(PvalRow) = True Then
                    oForm.Items.Item("b_display").Enabled = True
                    oForm.Items.Item("b_delete").Enabled = True
                Else
                    oForm.Items.Item("b_display").Enabled = False
                    oForm.Items.Item("b_delete").Enabled = False
                End If
            End If
        Catch ex As Exception
            oGFun.StatusBarErrorMsg("Attach Button Enble Function...")
        End Try
    End Sub

#End Region

#End Region

#Region "       Sub Grid Functions ...          "

    Sub SetNewLineSubGrid(ByVal UniqID As Integer, ByVal oMatSubGrid As SAPbouiCOM.Matrix, ByVal oDBDSSubGrid As SAPbouiCOM.DBDataSource, Optional ByVal RowID As Integer = 1, Optional ByVal ColumnUID As String = "", Optional ByVal DefaulFields As String(,) = Nothing)
        Try
            If ColumnUID.Equals("") = False Then
                If oMatSubGrid.Columns.Item(ColumnUID).Cells.Item(RowID).Specific.Value.Equals("") = False And RowID = oMatSubGrid.VisualRowCount Then
                    oMatSubGrid.FlushToDataSource()
                    oMatSubGrid.AddRow()
                    oDBDSSubGrid.InsertRecord(oDBDSSubGrid.Size)
                    oDBDSSubGrid.Offset = oMatSubGrid.VisualRowCount - 1
                    oDBDSSubGrid.SetValue("LineID", oDBDSSubGrid.Offset, oMatSubGrid.VisualRowCount)
                    oDBDSSubGrid.SetValue("U_UniqID", oDBDSSubGrid.Offset, UniqID)
                    If Not DefaulFields Is Nothing Then
                        For f As Int16 = 0 To DefaulFields.GetLength(0) - 1
                            oDBDSSubGrid.SetValue(DefaulFields(f, 0), oDBDSSubGrid.Offset, DefaulFields(f, 1))
                        Next
                    End If
                    oMatSubGrid.SetLineData(oMatSubGrid.VisualRowCount)
                    oMatSubGrid.FlushToDataSource()
                    oMatSubGrid.LoadFromDataSource()
                End If
            Else
                oMatSubGrid.FlushToDataSource()
                oMatSubGrid.AddRow()
                oDBDSSubGrid.InsertRecord(oDBDSSubGrid.Size)
                oDBDSSubGrid.Offset = oMatSubGrid.VisualRowCount - 1
                oDBDSSubGrid.SetValue("LineID", oDBDSSubGrid.Offset, oMatSubGrid.VisualRowCount)
                oDBDSSubGrid.SetValue("U_UniqID", oDBDSSubGrid.Offset, UniqID)
                If Not DefaulFields Is Nothing Then
                    For f As Int16 = 0 To DefaulFields.GetLength(0) - 1
                        oDBDSSubGrid.SetValue(DefaulFields(f, 0), oDBDSSubGrid.Offset, DefaulFields(f, 1))
                    Next
                End If
                oMatSubGrid.SetLineData(oMatSubGrid.VisualRowCount)
                oMatSubGrid.FlushToDataSource()
                oMatSubGrid.LoadFromDataSource()
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("SetNewLine Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub DeleteRowSubGrid(ByVal oMatMainSubGrid As SAPbouiCOM.Matrix, ByVal oDBDSMainSubGrid As SAPbouiCOM.DBDataSource, ByVal SubRowID As Integer)
        Try
            oMatMainSubGrid.FlushToDataSource()
            For i As Integer = 1 To oMatMainSubGrid.VisualRowCount

                If i <= oMatMainSubGrid.VisualRowCount Then

                    oMatMainSubGrid.GetLineData(i)
                    oDBDSMainSubGrid.Offset = i - 1
                    oDBDSMainSubGrid.SetValue("LineID", oDBDSMainSubGrid.Offset, i)

                    If Trim(oMatMainSubGrid.Columns.Item("uniqid").Cells.Item(i).Specific.Value).Equals(SubRowID.ToString()) = True Then
                        oMatMainSubGrid.DeleteRow(i)
                        i -= 1
                        GoTo CmdLine
                    ElseIf CDbl(Trim(oMatMainSubGrid.Columns.Item("uniqid").Cells.Item(i).Specific.Value)) > SubRowID Then
                        oDBDSMainSubGrid.SetValue("U_UniqID", oDBDSMainSubGrid.Offset, CDbl(Trim(oMatMainSubGrid.Columns.Item("uniqid").Cells.Item(i).Specific.Value)) - 1)
                    Else
                        oDBDSMainSubGrid.SetValue("U_UniqID", oDBDSMainSubGrid.Offset, CDbl(Trim(oMatMainSubGrid.Columns.Item("uniqid").Cells.Item(i).Specific.Value)))
                    End If
                    'oDBDSMainSubGrid.SetValue("U_ItemCode", oDBDSMainSubGrid.Offset, Trim(oMatMainSubGrid.Columns.Item("itemcode").Cells.Item(i).Specific.Value))
                    'oDBDSMainSubGrid.SetValue("U_Size", oDBDSMainSubGrid.Offset, Trim(oMatMainSubGrid.Columns.Item("size").Cells.Item(i).Specific.Value))
                    'oDBDSMainSubGrid.SetValue("U_Quantity", oDBDSMainSubGrid.Offset, Trim(oMatMainSubGrid.Columns.Item("quantity").Cells.Item(i).Specific.Value))
                    'oDBDSMainSubGrid.SetValue("U_Sounds", oDBDSMainSubGrid.Offset, Trim(oMatMainSubGrid.Columns.Item("sounds").Cells.Item(i).Specific.Value))
                    'oDBDSMainSubGrid.SetValue("U_Seconds", oDBDSMainSubGrid.Offset, Trim(oMatMainSubGrid.Columns.Item("seconds").Cells.Item(i).Specific.Value))
                    'oDBDSMainSubGrid.SetValue("U_Reject", oDBDSMainSubGrid.Offset, Trim(oMatMainSubGrid.Columns.Item("rejection").Cells.Item(i).Specific.Value))
                    'oDBDSMainSubGrid.SetValue("U_Remarks", oDBDSMainSubGrid.Offset, Trim(oMatMainSubGrid.Columns.Item("remarks").Cells.Item(i).Specific.Value))

                    'oMatMainSubGrid.SetLineData(i)
CmdLine:
                    oMatMainSubGrid.FlushToDataSource()
                End If
            Next
        Catch ex As Exception
            oApplication.StatusBar.SetText("Delete Row SubGrid Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Function LoadSubGrid(ByVal oMatSubGrid As SAPbouiCOM.Matrix, ByVal oDBDSSubGrid As SAPbouiCOM.DBDataSource, ByVal oDBDSMainSubGrid As SAPbouiCOM.DBDataSource, ByVal UniqID As Integer, Optional ByVal DefaulFields As String(,) = Nothing) As Boolean
        Try
            oMatSubGrid.Clear()
            oDBDSSubGrid.Clear()

            Dim strGetColUID As String
            If blnIsHANA Then
                strGetColUID = "Select ""AliasID"" From CUFD Where  ""AliasID"" <> 'UniqID' AND ""TableID"" ='" & oDBDSSubGrid.TableName & "' "
            Else
                strGetColUID = "Select AliasID From CUFD Where  AliasID <> 'UniqID' AND TableID ='" & oDBDSSubGrid.TableName & "' "
            End If
            Dim rsetGetColUID As SAPbobsCOM.Recordset = Me.DoQuery(strGetColUID)
            Dim strColUID As String = ""

            Dim boolUniqID As Boolean = False
            For x As Integer = 0 To oDBDSMainSubGrid.Size - 1
                If Trim(oDBDSMainSubGrid.GetValue("U_UniqID", x)).Equals(UniqID.ToString) Then
                    boolUniqID = True
                    Exit For
                End If
            Next
            If boolUniqID = False Then
                ' Add the Rows in Sub Grid ...
                Me.SetNewLineSubGrid(UniqID, oMatSubGrid, oDBDSSubGrid, , , DefaulFields)

            ElseIf oDBDSMainSubGrid.Size >= 1 And boolUniqID = True Then
                For i As Integer = 0 To oDBDSMainSubGrid.Size - 1
                    oDBDSMainSubGrid.Offset = i
                    If Trim(oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset)).Equals("") = False And oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim.Equals("") = False Then
                        If CInt(oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset)) = UniqID Then
                            oDBDSSubGrid.InsertRecord(oDBDSSubGrid.Size)
                            oDBDSSubGrid.Offset = oDBDSSubGrid.Size - 1

                            oDBDSSubGrid.SetValue("LineID", oDBDSSubGrid.Offset, oDBDSSubGrid.Offset + 1)
                            oDBDSSubGrid.SetValue("U_UniqID", oDBDSSubGrid.Offset, UniqID)
                            If Not DefaulFields Is Nothing Then
                                For f As Int16 = 0 To DefaulFields.GetLength(0) - 1
                                    oDBDSSubGrid.SetValue(DefaulFields(f, 0), oDBDSSubGrid.Offset, DefaulFields(f, 1))
                                Next
                            End If


                            rsetGetColUID.MoveFirst()
                            For j As Integer = 0 To rsetGetColUID.RecordCount - 1
                                strColUID = "U_" & rsetGetColUID.Fields.Item(0).Value
                                oDBDSSubGrid.SetValue(strColUID, oDBDSSubGrid.Offset, oDBDSMainSubGrid.GetValue(strColUID, oDBDSMainSubGrid.Offset).Trim)
                                rsetGetColUID.MoveNext()
                            Next
                        End If
                    End If
                Next
                oMatSubGrid.LoadFromDataSource()
                oMatSubGrid.FlushToDataSource()
                ' Add the Rows in Sub Grid ...
                Me.SetNewLineSubGrid(UniqID, oMatSubGrid, oDBDSSubGrid, , , DefaulFields)
            End If
            Return True
        Catch ex As Exception
            Me.StatusBarErrorMsg("Global Fun. : Load Sub Grid " & ex.Message)
            Return False
        Finally
        End Try
    End Function

    Function SaveSubGrid(ByVal oMatMainGrid As SAPbouiCOM.Matrix, ByVal oDBDSMainSubGrid As SAPbouiCOM.DBDataSource, ByVal oMatSubGrid As SAPbouiCOM.Matrix, ByVal oDBDsSubGrid As SAPbouiCOM.DBDataSource, ByVal RowID As Integer) As Boolean
        Try
            Dim strGetColUID As String
            If blnIsHANA Then
                strGetColUID = "Select ""AliasID"" From CUFD Where  ""AliasID"" <> 'UniqID' AND ""TableID"" ='" & oDBDsSubGrid.TableName & "' "
            Else
                strGetColUID = "Select AliasID From CUFD Where  AliasID <> 'UniqID' AND TableID ='" & oDBDsSubGrid.TableName & "' "
            End If


            Dim rsetGetColUID As SAPbobsCOM.Recordset = Me.DoQuery(strGetColUID)
            Dim strColUID As String = ""

            Dim intInitSize As Integer = oDBDSMainSubGrid.Size
            Dim intCurrSize As Integer = oDBDSMainSubGrid.Size
            Dim oEmpty As Boolean = True


            For i As Integer = 0 To intInitSize - 1
                oDBDSMainSubGrid.Offset = i - (intInitSize - intCurrSize)
                Dim aa = oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim
                If oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim <> "" Then
                    If CInt(oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset)) = RowID Then
                        oDBDSMainSubGrid.RemoveRecord(oDBDSMainSubGrid.Offset)
                    End If
                ElseIf oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim.Equals("") Then
                    oDBDSMainSubGrid.RemoveRecord(oDBDSMainSubGrid.Offset)
                End If
                intCurrSize = oDBDSMainSubGrid.Size
            Next

            If oDBDSMainSubGrid.Size = 0 Then oDBDSMainSubGrid.InsertRecord(oDBDSMainSubGrid.Size)
            oDBDSMainSubGrid.Offset = 0

            If oDBDSMainSubGrid.Size = 1 And Trim(oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset)).Equals("") Then
                oEmpty = True
            Else
                oEmpty = False
            End If
            oMatSubGrid.FlushToDataSource()
            oMatSubGrid.LoadFromDataSource()

            For i As Integer = 0 To oMatSubGrid.VisualRowCount - 2
                oDBDsSubGrid.Offset = i

                If oEmpty = True Then
                    oDBDSMainSubGrid.Offset = oDBDSMainSubGrid.Size - 1
                    oDBDSMainSubGrid.SetValue("U_UniqID", oDBDSMainSubGrid.Offset, RowID)

                    rsetGetColUID.MoveFirst()
                    For j As Integer = 0 To rsetGetColUID.RecordCount - 1
                        strColUID = "U_" & rsetGetColUID.Fields.Item(0).Value
                        oDBDSMainSubGrid.SetValue(strColUID, oDBDSMainSubGrid.Offset, oDBDsSubGrid.GetValue(strColUID, i).Trim)
                        rsetGetColUID.MoveNext()
                    Next

                    If i <> (oMatSubGrid.VisualRowCount - 1) Then
                        oDBDSMainSubGrid.InsertRecord(oDBDSMainSubGrid.Size)
                    End If
                ElseIf oEmpty = False Then
                    oDBDSMainSubGrid.InsertRecord(oDBDSMainSubGrid.Size)
                    oDBDSMainSubGrid.Offset = oDBDSMainSubGrid.Size - 1
                    oDBDSMainSubGrid.SetValue("U_UniqID", oDBDSMainSubGrid.Offset, RowID)

                    rsetGetColUID.MoveFirst()
                    For j As Integer = 0 To rsetGetColUID.RecordCount - 1
                        strColUID = "U_" & rsetGetColUID.Fields.Item(0).Value
                        oDBDSMainSubGrid.SetValue(strColUID, oDBDSMainSubGrid.Offset, oDBDsSubGrid.GetValue(strColUID, i).Trim)
                        rsetGetColUID.MoveNext()
                    Next
                End If
            Next
            oMatMainGrid.LoadFromDataSource()

            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Save Sub Grid Method Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function

#End Region

#Region " ...  Set Date To Matrix ... "

    Sub setData(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSDetail As SAPbouiCOM.DBDataSource, _
    ByVal strRowNum As Integer, ByVal strFldName As String, ByVal strValue As String)
        Try
            If strRowNum <= 0 Then Return
            oDBDSDetail.Offset = strRowNum - 1
            'For i As Integer = 1 To oMatrix.Columns.Count
            '    oDBDSDetail.SetValue(oMatrix.Columns.Item(1).DataBind.Alias, oDBDSDetail.Offset, _
            '    oDBDSDetail.GetValue(oMatrix.Columns.Item(1).DataBind.Alias, oDBDSDetail.Offset))
            'Next              
            Try
                oDBDSDetail.SetValue(strFldName, oDBDSDetail.Offset, strValue)
                oMatrix.SetLineData(oDBDSDetail.Offset + 1)
            Catch ex As Exception
            End Try
            oMatrix.FlushToDataSource()
            ' oMatrix.Columns.Item(oDBDSDetail.Offset + 1).Cells.Item(strRowNum).Click()
        Catch ex As Exception
            oGFun.Msg("SetData Function Failed:" & ex.Message, "S", "W")
        Finally
        End Try
    End Sub
#End Region

#Region "...Genel İşlemler..."

    Private Function GetDocEntryFromXML(sXML As String) As Integer
        Dim xDoc As New XmlDocument()

        xDoc.LoadXml(sXML)
        Dim sDocEntry As String = xDoc.SelectSingleNode("DocumentParams/DocEntry").InnerText
        If sDocEntry <> "" Then
            Return Convert.ToInt32(sDocEntry)
        Else
            Return -1
        End If
    End Function
    'Sub SS()
    '    Dim oJurnal As SAPbobsCOM.JournalEntries = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
    '    oJurnal.Lines.AccountCode = "120010100"
    '    oJurnal.Lines.FCCurrency = "EUR"
    '    oJurnal.Lines.FCDebit = "500"

    '    oJurnal.Lines.Add()

    '    oJurnal.Lines.AccountCode = "118010200"
    '    oJurnal.Lines.FCCurrency = "EUR"
    '    oJurnal.Lines.FCCredit = "500"
    '    Dim oi As Integer = 0
    '    oi = oJurnal.Add
    '    If oi = 0 Then
    '    Else
    '        oApplication.MessageBox(oCompany.GetLastErrorDescription)
    '    End If
    'End Sub
    Function GunlukKurBilgileriniOku() As Array
        Dim arrKur(1, 1) As String
        Dim orsKurDegerleri As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If blnIsHANA Then
            orsKurDegerleri.DoQuery("SELECT DISTINCT ""Currency"", ""Rate"" FROM ORTT WHERE ""RateDate"" = CAST(CAST(NOW() AS varchar(10)) AS timestamp)")
        Else
            orsKurDegerleri.DoQuery("SELECT DISTINCT Currency, Rate FROM [dbo].[ORTT](NOLOCK) WHERE RateDate= CONVERT(DATETIME, CONVERT(VARCHAR(10),GETDATE(),112))")
        End If

        If orsKurDegerleri.RecordCount > 0 Then
            orsKurDegerleri.MoveFirst()
            Dim k As Integer = 0
            ReDim arrKur(1, orsKurDegerleri.RecordCount - 1)
            For k = 0 To orsKurDegerleri.RecordCount - 1
                'MessageBox.Show(orsKurDegerleri.Fields.Item(0).Value.ToString)
                arrKur(0, k) = orsKurDegerleri.Fields.Item(0).Value
                arrKur(1, k) = orsKurDegerleri.Fields.Item(1).Value
                orsKurDegerleri.MoveNext()
            Next

        Else
            arrKur = Nothing
        End If
        Return arrKur
    End Function
    Function MuhasebelestirKurFarkli(ByVal Account As String, ByVal AccountCode As String, ByVal ContraAccount As String, ByVal ContraAccountCode As String, ByVal dblAmount As Double, ByVal DueDate As String, ByVal dblMasrafTutar As Double, ByVal MasrafHesabi As String, ByVal ParaBirimi As String, ByVal Proje As String, ByVal Aciklama As String, ByVal KayitTarihi As String, ByVal BelgeTarihi As String, ByVal dblKur As Double) As String
        Dim strerr As String = ""
        Dim GunlukKur As Double = 1
        Dim ArrGunlukKur(,) As String
        Try

            Dim isatir As Integer = 1
            Dim vJE As SAPbobsCOM.JournalEntries

            ArrGunlukKur = oGFun.GunlukKurBilgileriniOku()
            If Not ArrGunlukKur Is Nothing Then
                For e As Integer = 0 To (ArrGunlukKur.Length / 2) - 1
                    If Not ArrGunlukKur(0, e) Is Nothing Then
                        If ArrGunlukKur(0, e) = ParaBirimi Then
                            GunlukKur = ArrGunlukKur(1, e)
                            Exit For
                        Else
                            GunlukKur = 1
                        End If
                    End If
                Next
            Else
                GunlukKur = 1
            End If



            'GunlukKur = 3.4

            Dim oQuery As String
            If blnIsHANA Then
                oQuery = "SELECT ""GLGainXdif"", ""GLLossXdif"" FROM OACP WHERE ""PeriodCat"" = YEAR(NOW())"
            Else
                oQuery = "SELECT GLGainXdif, GLLossXdif FROM dbo.OACP(nolock) WHERE PeriodCat = DATEPART(yy, getdate())"
            End If

            Dim rsetValue As SAPbobsCOM.Recordset = oGFun.DoQuery(oQuery)
            rsetValue.MoveFirst()
            Dim BorcKur As String = rsetValue.Fields.Item("GLGainXdif").Value.ToString()
            Dim AlacakKur As String = rsetValue.Fields.Item("GLLossXdif").Value.ToString()
            Dim KurFark As Double = (GunlukKur * dblAmount) - (dblKur * dblAmount)


            vJE = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            'vJE.AutoVAT = SAPbobsCOM.BoYesNoEnum.tYES 'test ..

            vJE.Memo = Aciklama
            vJE.DueDate = SetDateFormat2(DueDate)
            vJE.ReferenceDate = SetDateFormat2(KayitTarihi)
            vJE.TaxDate = SetDateFormat2(BelgeTarihi)

            vJE.Lines.AccountCode = RTrim(LTrim(Account))
            vJE.Lines.ShortName = RTrim(LTrim(AccountCode))
            vJE.Lines.ContraAccount = RTrim(LTrim(ContraAccountCode))

            If ParaBirimi = "TRY" Or ParaBirimi = "TL" Or ParaBirimi = "YTL" Then
                vJE.Lines.Credit = dblAmount
                vJE.Lines.Debit = 0
            Else
                vJE.Lines.FCCurrency = ParaBirimi
                vJE.Lines.FCCredit = dblAmount
                vJE.Lines.FCDebit = 0

                vJE.Lines.Credit = dblKur * dblAmount
                vJE.Lines.Debit = 0
            End If
            'vJE.Lines.Credit = dblAmount
            'vJE.Lines.Debit = 0

            vJE.Lines.DueDate = SetDateFormat2(DueDate)
            vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
            vJE.Lines.LineMemo = Aciklama
            vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)
            Call vJE.Lines.Add()



            If dblMasrafTutar <> 0 Then
                Call vJE.Lines.SetCurrentLine(isatir)
                vJE.Lines.AccountCode = RTrim(LTrim(MasrafHesabi))  '"118010200"
                vJE.Lines.ShortName = RTrim(LTrim(MasrafHesabi)) '"118010200"
                vJE.Lines.ContraAccount = RTrim(LTrim(AccountCode))  '"1201001"
                vJE.Lines.Credit = 0
                vJE.Lines.Debit = dblMasrafTutar
                vJE.Lines.DueDate = SetDateFormat2(DueDate)
                'vJE.Lines.Line_ID = 1
                vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
                vJE.Lines.LineMemo = Aciklama
                vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)

                isatir = isatir + 1
                Call vJE.Lines.Add()
            End If

            If KurFark > 0 Then
                Call vJE.Lines.SetCurrentLine(isatir)
                vJE.Lines.AccountCode = RTrim(LTrim(AlacakKur))
                'vJE.Lines.ShortName = RTrim(LTrim(MasrafHesabi)) '"118010200"
                'vJE.Lines.ContraAccount = RTrim(LTrim(AccountCode))
                vJE.Lines.Credit = KurFark
                vJE.Lines.Debit = 0
                vJE.Lines.DueDate = SetDateFormat2(DueDate)
                'vJE.Lines.Line_ID = 1
                vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
                vJE.Lines.LineMemo = Aciklama
                vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)

                isatir = isatir + 1
                Call vJE.Lines.Add()
            ElseIf KurFark < 0 Then

                Call vJE.Lines.SetCurrentLine(isatir)
                vJE.Lines.AccountCode = RTrim(LTrim(BorcKur))  '"118010200"
                vJE.Lines.ContraAccount = RTrim(LTrim(AccountCode))
                vJE.Lines.Credit = 0
                vJE.Lines.Debit = -1 * KurFark
                vJE.Lines.DueDate = SetDateFormat2(DueDate)

                'vJE.Lines.Line_ID = 1
                vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
                vJE.Lines.LineMemo = Aciklama
                vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)

                isatir = isatir + 1
                Call vJE.Lines.Add()
            End If

            Call vJE.Lines.SetCurrentLine(isatir)
            vJE.Lines.AccountCode = RTrim(LTrim(ContraAccount))  '"118010200"
            vJE.Lines.ShortName = RTrim(LTrim(ContraAccountCode)) '"118010200"
            vJE.Lines.ContraAccount = RTrim(LTrim(AccountCode))  '"1201001"

            If ParaBirimi = "TRY" Or ParaBirimi = "TL" Or ParaBirimi = "YTL" Then
                vJE.Lines.Credit = 0
                vJE.Lines.Debit = dblAmount - dblMasrafTutar
            Else
                vJE.Lines.FCCurrency = ParaBirimi
                vJE.Lines.FCCredit = 0
                vJE.Lines.FCDebit = dblAmount - (dblMasrafTutar / GunlukKur)

                vJE.Lines.Credit = 0
                vJE.Lines.Debit = GunlukKur * (dblAmount - (dblMasrafTutar / GunlukKur))
            End If

            vJE.Lines.ProjectCode = Proje
            vJE.Lines.DueDate = SetDateFormat2(DueDate)

            vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
            vJE.Lines.LineMemo = Aciklama
            vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)

            Dim vc_Message_Error_string As String
            Dim vc_Message_Result_Int32 As Integer = 0
            vc_Message_Result_Int32 = vJE.Add()

            Dim DocEntry As String = ""

            oCompany.GetNewObjectCode(DocEntry)
            'MessageBox.Show("n1: " + DocEntry)


            If vc_Message_Result_Int32 <> 0 Then
                Dim vm_GetLastErrorDescription_string As String = oCompany.GetLastErrorDescription()
                oCompany.GetLastError(vc_Message_Result_Int32, vc_Message_Error_string)
                'Return "Error Add-On (Procons): ERR-720 " + vc_Message_Error_string + " Error Code: " + vc_Message_Result_Int32.ToString()
                Return "Error :" + vc_Message_Error_string + " Hata Kodu: " + vc_Message_Result_Int32.ToString()
            Else
                'MsgBox("İslem basarılı")
                'vJE.SaveXML("c:\temp\JournalEntries" + DocEntry + ".xml")
                Return "BelgeNo" + DocEntry
            End If

        Catch ex As Exception
            oGFun.Msg("Error :" & ex.Message)

        End Try

    End Function

    Function TersMuhasebelestirKurFarkli(ByVal Account As String, ByVal AccountCode As String, ByVal ContraAccount As String, ByVal ContraAccountCode As String, ByVal dblAmount As Double, ByVal DueDate As Date, ByVal dblMasrafTutar As Double, ByVal MasrafHesabi As String, ByVal ParaBirimi As String, ByVal Proje As String, ByVal Aciklama As String, ByVal KayitTarihi As String, ByVal BelgeTarihi As String, ByVal dblKur As Double, ByVal dblKurOnceki As Double) As String
        Try
            'Account ve ContraAccount hesap kodları, AccountCode ise muhatapkodu dur..
            Dim isatir As Integer = 1
            Dim vJE As SAPbobsCOM.JournalEntries

            'GunlukKur = 3.4
            Dim oQuery As String
            If blnIsHANA Then
                oQuery = "SELECT ""GLGainXdif"", ""GLLossXdif"" FROM OACP WHERE ""PeriodCat"" = YEAR(NOW())"
            Else
                oQuery = "SELECT GLGainXdif, GLLossXdif FROM dbo.OACP(nolock) WHERE PeriodCat = DATEPART(yy, getdate())"
            End If

            Dim rsetValue As SAPbobsCOM.Recordset = oGFun.DoQuery(oQuery)
            rsetValue.MoveFirst()
            Dim BorcKur As String = rsetValue.Fields.Item("GLGainXdif").Value.ToString()
            Dim AlacakKur As String = rsetValue.Fields.Item("GLLossXdif").Value.ToString()
            Dim KurFark As Double = (dblKur * dblAmount) - (dblKurOnceki * dblAmount)

            vJE = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            vJE.Memo = Aciklama
            vJE.DueDate = SetDateFormat2(DateParse2(DueDate))
            vJE.ReferenceDate = SetDateFormat2(KayitTarihi)
            vJE.TaxDate = SetDateFormat2(BelgeTarihi)

            vJE.Lines.AccountCode = RTrim(LTrim(Account)) '"120010100" 
            vJE.Lines.ShortName = RTrim(LTrim(AccountCode)) '"1201001"
            vJE.Lines.ContraAccount = RTrim(LTrim(ContraAccountCode)) ' "118010200"
            vJE.Lines.Credit = 0
            vJE.Lines.Debit = dblAmount - (KurFark / dblKur)
            'vJE.Lines.DueDate = DueDate
            'vJE.Lines.ReferenceDate1 = Now
            vJE.Lines.LineMemo = Aciklama
            'vJE.Lines.TaxDate = Now
            vJE.Lines.DueDate = SetDateFormat2(DateParse2(DueDate))
            vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
            vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)


            Call vJE.Lines.Add()

            If dblMasrafTutar <> 0 Then
                Call vJE.Lines.SetCurrentLine(isatir)
                vJE.Lines.AccountCode = RTrim(LTrim(MasrafHesabi))  '"118010200"
                vJE.Lines.ShortName = RTrim(LTrim(MasrafHesabi)) '"118010200"
                vJE.Lines.ContraAccount = RTrim(LTrim(AccountCode))  '"1201001"
                vJE.Lines.Credit = dblMasrafTutar
                vJE.Lines.Debit = 0
                'vJE.Lines.DueDate = DueDate
                'vJE.Lines.Line_ID = 1
                'vJE.Lines.ReferenceDate1 = Now
                vJE.Lines.LineMemo = Aciklama
                'vJE.Lines.TaxDate = Now
                vJE.Lines.DueDate = SetDateFormat2(DateParse2(DueDate))
                vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
                vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)

                isatir = isatir + 1
                Call vJE.Lines.Add()
            End If

            If KurFark > 0 Then
                Call vJE.Lines.SetCurrentLine(isatir)
                vJE.Lines.AccountCode = RTrim(LTrim(AlacakKur))
                'vJE.Lines.ShortName = RTrim(LTrim(MasrafHesabi)) '"118010200"
                'vJE.Lines.ContraAccount = RTrim(LTrim(AccountCode))
                vJE.Lines.Credit = 0
                vJE.Lines.Debit = KurFark
                vJE.Lines.DueDate = SetDateFormat2(DueDate)
                'vJE.Lines.Line_ID = 1
                vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
                vJE.Lines.LineMemo = Aciklama
                vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)

                isatir = isatir + 1
                Call vJE.Lines.Add()
            ElseIf KurFark < 0 Then

                Call vJE.Lines.SetCurrentLine(isatir)
                vJE.Lines.AccountCode = RTrim(LTrim(BorcKur))  '"118010200"
                vJE.Lines.ContraAccount = RTrim(LTrim(AccountCode))
                vJE.Lines.Credit = -1 * KurFark
                vJE.Lines.Debit = 0
                vJE.Lines.DueDate = SetDateFormat2(DueDate)

                'vJE.Lines.Line_ID = 1
                vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
                vJE.Lines.LineMemo = Aciklama
                vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)

                isatir = isatir + 1
                Call vJE.Lines.Add()
            End If


            Call vJE.Lines.SetCurrentLine(isatir)
            vJE.Lines.AccountCode = RTrim(LTrim(ContraAccount))  '"118010200"
            vJE.Lines.ShortName = RTrim(LTrim(ContraAccountCode)) '"118010200"
            vJE.Lines.ContraAccount = RTrim(LTrim(AccountCode))  '"1201001"
            vJE.Lines.Credit = dblAmount - (dblMasrafTutar / dblKur) 'dblAmount - (KurFark / dblKur)
            vJE.Lines.ProjectCode = Proje
            vJE.Lines.Debit = 0
            'vJE.Lines.DueDate = DueDate
            ''vJE.Lines.Line_ID = 1
            'vJE.Lines.ReferenceDate1 = Now
            vJE.Lines.LineMemo = Aciklama

            vJE.Lines.DueDate = SetDateFormat2(DateParse2(DueDate))
            vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
            vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)

            Dim vc_Message_Error_string As String
            Dim vc_Message_Result_Int32 As Integer = 0
            vc_Message_Result_Int32 = vJE.Add()

            Dim DocEntry As String = ""

            oCompany.GetNewObjectCode(DocEntry)
            'MessageBox.Show("n1: " + DocEntry)


            If vc_Message_Result_Int32 <> 0 Then
                Dim vm_GetLastErrorDescription_string As String = oCompany.GetLastErrorDescription()
                oCompany.GetLastError(vc_Message_Result_Int32, vc_Message_Error_string)
                'Return ("Add-On Error (Procons): ERR-720 \n " + vc_Message_Error_string + " Error Code: " + vc_Message_Result_Int32)
                Return "Error :" + vc_Message_Error_string + " Hata Kodu: " + vc_Message_Result_Int32.ToString()
            Else
                'MsgBox("İslem basarılı")
                'vJE.SaveXML("c:\temp\JournalEntries" + DocEntry + ".xml")
                Return "BelgeNo" + DocEntry
            End If

        Catch ex As Exception
            'oGFun.Msg("Error :" & ex.Message)
            Return "Error :" & ex.Message
        End Try

    End Function


    Function Muhasebelestir(ByVal Account As String, ByVal AccountCode As String, ByVal ContraAccount As String, ByVal ContraAccountCode As String, ByVal dblAmount As Double, ByVal DueDate As String, ByVal dblMasrafTutar As Double, ByVal MasrafHesabi As String, ByVal ParaBirimi As String, ByVal Proje As String, ByVal Aciklama As String, ByVal KayitTarihi As String, ByVal BelgeTarihi As String, ByVal dblKur As Double) As String
        Dim strerr As String = ""
        Dim GunlukKur As Double = 1
        Dim ArrGunlukKur(,) As String
        Try

            Dim isatir As Integer = 1
            Dim vJE As SAPbobsCOM.JournalEntries

            ArrGunlukKur = oGFun.GunlukKurBilgileriniOku()
            If Not ArrGunlukKur Is Nothing Then
                For e As Integer = 0 To (ArrGunlukKur.Length / 2) - 1
                    If Not ArrGunlukKur(0, e) Is Nothing Then
                        If ArrGunlukKur(0, e) = ParaBirimi Then
                            GunlukKur = ArrGunlukKur(1, e)
                            Exit For
                        Else
                            GunlukKur = 1
                        End If
                    End If
                Next
            Else
                GunlukKur = 1
            End If


            vJE = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            'vJE.AutoVAT = SAPbobsCOM.BoYesNoEnum.tYES 'test ..

            vJE.Memo = Aciklama
            vJE.DueDate = SetDateFormat2(DueDate)
            vJE.ReferenceDate = SetDateFormat2(KayitTarihi)
            vJE.TaxDate = SetDateFormat2(BelgeTarihi)

            vJE.Lines.AccountCode = RTrim(LTrim(Account))
            vJE.Lines.ShortName = RTrim(LTrim(AccountCode))
            vJE.Lines.ContraAccount = RTrim(LTrim(ContraAccountCode))

            If ParaBirimi = "TRY" Or ParaBirimi = "TL" Or ParaBirimi = "YTL" Then
                vJE.Lines.Credit = dblAmount
                vJE.Lines.Debit = 0
            Else
                vJE.Lines.FCCurrency = ParaBirimi
                vJE.Lines.FCCredit = dblAmount
                vJE.Lines.FCDebit = 0

                vJE.Lines.Credit = dblKur * dblAmount
                vJE.Lines.Debit = 0
            End If
            'vJE.Lines.Credit = dblAmount
            'vJE.Lines.Debit = 0

            vJE.Lines.DueDate = SetDateFormat2(DueDate)
            vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
            vJE.Lines.LineMemo = Aciklama
            vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)
            Call vJE.Lines.Add()



            If dblMasrafTutar <> 0 Then
                Call vJE.Lines.SetCurrentLine(isatir)
                vJE.Lines.AccountCode = RTrim(LTrim(MasrafHesabi))  '"118010200"
                vJE.Lines.ShortName = RTrim(LTrim(MasrafHesabi)) '"118010200"
                vJE.Lines.ContraAccount = RTrim(LTrim(AccountCode))  '"1201001"
                vJE.Lines.Credit = 0
                vJE.Lines.Debit = dblMasrafTutar
                vJE.Lines.DueDate = SetDateFormat2(DueDate)
                'vJE.Lines.Line_ID = 1
                vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
                vJE.Lines.LineMemo = Aciklama
                vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)

                isatir = isatir + 1
                Call vJE.Lines.Add()
            End If


            Call vJE.Lines.SetCurrentLine(isatir)
            vJE.Lines.AccountCode = RTrim(LTrim(ContraAccount))  '"118010200"
            vJE.Lines.ShortName = RTrim(LTrim(ContraAccountCode)) '"118010200"
            vJE.Lines.ContraAccount = RTrim(LTrim(AccountCode))  '"1201001"

            If ParaBirimi = "TRY" Or ParaBirimi = "TL" Or ParaBirimi = "YTL" Then
                vJE.Lines.Credit = 0
                vJE.Lines.Debit = dblAmount - dblMasrafTutar
            Else
                vJE.Lines.FCCurrency = ParaBirimi
                vJE.Lines.FCCredit = 0
                vJE.Lines.FCDebit = dblAmount - (dblMasrafTutar / GunlukKur)

                vJE.Lines.Credit = 0
                vJE.Lines.Debit = GunlukKur * (dblAmount - (dblMasrafTutar / GunlukKur))
            End If

            vJE.Lines.ProjectCode = Proje
            vJE.Lines.DueDate = SetDateFormat2(DueDate)

            vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
            vJE.Lines.LineMemo = Aciklama
            vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)

            Dim vc_Message_Error_string As String
            Dim vc_Message_Result_Int32 As Integer = 0
            vc_Message_Result_Int32 = vJE.Add()

            Dim DocEntry As String = ""

            oCompany.GetNewObjectCode(DocEntry)
            'MessageBox.Show("n1: " + DocEntry)


            If vc_Message_Result_Int32 <> 0 Then
                Dim vm_GetLastErrorDescription_string As String = oCompany.GetLastErrorDescription()
                oCompany.GetLastError(vc_Message_Result_Int32, vc_Message_Error_string)
                'Return "Error Add-On (Procons): ERR-720 " + vc_Message_Error_string + " Error Code: " + vc_Message_Result_Int32.ToString()
                Return "Error :" + vc_Message_Error_string + " Hata Kodu: " + vc_Message_Result_Int32.ToString()
            Else
                'MsgBox("İslem basarılı")
                'vJE.SaveXML("c:\temp\JournalEntries" + DocEntry + ".xml")
                Return "BelgeNo" + DocEntry
            End If

        Catch ex As Exception
            oGFun.Msg("Error :" & ex.Message)

        End Try
    End Function

    Function TersMuhasebelestir(ByVal Account As String, ByVal AccountCode As String, ByVal ContraAccount As String, ByVal ContraAccountCode As String, ByVal dblAmount As Double, ByVal DueDate As Date, ByVal dblMasrafTutar As Double, ByVal MasrafHesabi As String, ByVal ParaBirimi As String, ByVal Proje As String, ByVal Aciklama As String, ByVal KayitTarihi As String, ByVal BelgeTarihi As String, ByVal dblKur As Double) As String
        Try
            'Account ve ContraAccount hesap kodları, AccountCode ise muhatapkodu dur..
            Dim isatir As Integer = 1
            Dim vJE As SAPbobsCOM.JournalEntries

            vJE = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            vJE.Memo = Aciklama
            vJE.DueDate = SetDateFormat2(DateParse2(DueDate))
            vJE.ReferenceDate = SetDateFormat2(KayitTarihi)
            vJE.TaxDate = SetDateFormat2(BelgeTarihi)

            vJE.Lines.AccountCode = RTrim(LTrim(Account)) '"120010100" 
            vJE.Lines.ShortName = RTrim(LTrim(AccountCode)) '"1201001"
            vJE.Lines.ContraAccount = RTrim(LTrim(ContraAccountCode)) ' "118010200"
            vJE.Lines.Credit = 0
            vJE.Lines.Debit = dblAmount
            'vJE.Lines.DueDate = DueDate
            'vJE.Lines.ReferenceDate1 = Now
            vJE.Lines.LineMemo = Aciklama
            'vJE.Lines.TaxDate = Now
            vJE.Lines.DueDate = SetDateFormat2(DateParse2(DueDate))
            vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
            vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)


            Call vJE.Lines.Add()

            If dblMasrafTutar <> 0 Then
                Call vJE.Lines.SetCurrentLine(isatir)
                vJE.Lines.AccountCode = RTrim(LTrim(MasrafHesabi))  '"118010200"
                vJE.Lines.ShortName = RTrim(LTrim(MasrafHesabi)) '"118010200"
                vJE.Lines.ContraAccount = RTrim(LTrim(AccountCode))  '"1201001"
                vJE.Lines.Credit = dblMasrafTutar
                vJE.Lines.Debit = 0
                'vJE.Lines.DueDate = DueDate
                'vJE.Lines.Line_ID = 1
                'vJE.Lines.ReferenceDate1 = Now
                vJE.Lines.LineMemo = Aciklama
                'vJE.Lines.TaxDate = Now
                vJE.Lines.DueDate = SetDateFormat2(DateParse2(DueDate))
                vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
                vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)

                isatir = isatir + 1
                Call vJE.Lines.Add()
            End If

            Call vJE.Lines.SetCurrentLine(isatir)
            vJE.Lines.AccountCode = RTrim(LTrim(ContraAccount))  '"118010200"
            vJE.Lines.ShortName = RTrim(LTrim(ContraAccountCode)) '"118010200"
            vJE.Lines.ContraAccount = RTrim(LTrim(AccountCode))  '"1201001"
            vJE.Lines.Credit = dblAmount - (dblMasrafTutar / dblKur)
            vJE.Lines.ProjectCode = Proje
            vJE.Lines.Debit = 0
            'vJE.Lines.DueDate = DueDate
            ''vJE.Lines.Line_ID = 1
            'vJE.Lines.ReferenceDate1 = Now
            vJE.Lines.LineMemo = Aciklama

            vJE.Lines.DueDate = SetDateFormat2(DateParse2(DueDate))
            vJE.Lines.ReferenceDate1 = SetDateFormat2(KayitTarihi)
            vJE.Lines.TaxDate = SetDateFormat2(BelgeTarihi)

            Dim vc_Message_Error_string As String
            Dim vc_Message_Result_Int32 As Integer = 0
            vc_Message_Result_Int32 = vJE.Add()

            Dim DocEntry As String = ""

            oCompany.GetNewObjectCode(DocEntry)
            'MessageBox.Show("n1: " + DocEntry)


            If vc_Message_Result_Int32 <> 0 Then
                Dim vm_GetLastErrorDescription_string As String = oCompany.GetLastErrorDescription()
                oCompany.GetLastError(vc_Message_Result_Int32, vc_Message_Error_string)
                'Return ("Add-On Error (Procons): ERR-720 \n " + vc_Message_Error_string + " Error Code: " + vc_Message_Result_Int32)
                Return "Error :" + vc_Message_Error_string + " Hata Kodu: " + vc_Message_Result_Int32.ToString()
            Else
                'MsgBox("İslem basarılı")
                'vJE.SaveXML("c:\temp\JournalEntries" + DocEntry + ".xml")
                Return "BelgeNo" + DocEntry
            End If

        Catch ex As Exception
            'oGFun.Msg("Error :" & ex.Message)
            Return "Error :" & ex.Message
        End Try

    End Function


    Private Sub TahsilatDokumanOusturmaOrnek()

        Dim vPay As SAPbobsCOM.Payments
        vPay = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
        vPay.Address = "622-7"
        vPay.ApplyVAT = 1
        vPay.CardCode = "D10006"
        vPay.CardName = "Card D10004"
        vPay.CashAccount = "288000"
        vPay.CashSum = 0 '5620.85
        vPay.CheckAccount = "280001"
        vPay.ContactPersonCode = 2
        vPay.DocCurrency = "Eur"
        vPay.DocDate = Now
        vPay.DocRate = 0
        vPay.DocTypte = 0
        vPay.HandWritten = 0
        vPay.JournalRemarks = "Incoming - D10004"
        vPay.LocalCurrency = 0
        'vPay.Printed = 0
        vPay.Reference1 = 14
        vPay.Series = 0
        vPay.SplitTransaction = 0
        vPay.TaxDate = Now
        vPay.TransferAccount = "10100"
        vPay.TransferDate = Now
        vPay.TransferSum = 0 '

        vPay.Invoices.AppliedFC = 0
        vPay.Invoices.AppliedSys = 5031.2
        vPay.Invoices.DocEntry = 8
        vPay.Invoices.DocLine = 0
        vPay.Invoices.DocRate = 0
        vPay.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
        'vPay.Invoices.LineNum = 0
        vPay.Invoices.SumApplied = 5031.2


        vPay.Checks.AccounttNum = "384838"
        vPay.Checks.BankCode = "CB"
        vPay.Checks.Branch = "Branch"
        vPay.Checks.CheckNumber = 3838383
        vPay.Checks.CheckSum = 4031.2
        vPay.Checks.Currency = "Eur"
        vPay.Checks.Details = "test1"
        vPay.Checks.DueDate = CDate("31/12/2002")
        'vPay.Checks.LineNum = 0
        vPay.Checks.Trnsfrable = 0
        Call vPay.Checks.Add()
        Call vPay.Checks.SetCurentLine(1)
        vPay.Checks.AccounttNum = "384838"
        vPay.Checks.BankCode = "CB"
        vPay.Checks.Branch = "Branch"
        vPay.Checks.CheckNumber = "3838383"
        vPay.Checks.CheckSum = 1000
        vPay.Checks.Currency = "EUR"
        vPay.Checks.Details = "test2"
        vPay.Checks.DueDate = Now
        'vPay.Checks.LineNum = 1
        vPay.Checks.Trnsfrable = 0
        If (vPay.Add() <> 0) Then
            MsgBox("Failed to add a payment")
        End If

        Dim vc_Message_Error_string As String
        Dim vc_Message_Result_Int32 As Integer = 0
        vc_Message_Result_Int32 = vPay.Add()

        Dim DocEntry As String = ""

        oCompany.GetNewObjectCode(DocEntry)
        'MessageBox.Show("n1: " + DocEntry)


        If vc_Message_Result_Int32 <> 0 Then
            Dim vm_GetLastErrorDescription_string As String = oCompany.GetLastErrorDescription()
            oCompany.GetLastError(vc_Message_Result_Int32, vc_Message_Error_string)
            'Return ("Add-On Error (Procons): ERR-720 \n " + vc_Message_Error_string + " Error Code: " + vc_Message_Result_Int32)
            MessageBox.Show("Add-On Error (Procons): ERR-720 \n " + vc_Message_Error_string + " Error Code: " + vc_Message_Result_Int32)
        Else
            'MsgBox("İslem basarılı")
            vPay.SaveXML("c:\temp\InvoiceEntries" + DocEntry + ".xml")
            MessageBox.Show("BelgeNo" + DocEntry)
        End If

        ''check for errors
        'Call vCompany.GetLastError(nErr, errMsg)
        'If (0 <> nErr) Then
        '    MsgBox("Found error:" + Str(nErr) + "," + errMsg)
        'Else
        '    MsgBox("Succeed in payment.add")
        'End If

        ''disconnect the company object, and release resource
        'Call vCompany.Disconnect()
        'vCompany = Nothing
        'Exit Sub
        'ErrorHandler:
        '        MsgBox("Exception:" + Err.Description)
    End Sub
#End Region
End Class
