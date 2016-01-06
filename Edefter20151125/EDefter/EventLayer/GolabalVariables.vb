Module GolabalVariables
#Region " ... Common For SAP ..."
    Public oCompany As SAPbobsCOM.Company
    Public oGFun As New GlobalFunctions
    Public oForm As SAPbouiCOM.Form
    Public strOndalikAyirac As String = String.Empty
    Public strTarihAyirac As String = String.Empty

#End Region

#Region "       ... Common For Forms ...        "
   
    Public Kur As Double = 0
    Public SirketAdi As String = ""

#End Region
#Region ""
    'E-Defter
    Public EDefEntegratorSecimXml = "EDefEntegratorSecim.xml", EDefEntegratorSecim_FormUID As String = "ENSEC"
    Public oEDefEntegratorSecim As New ClsEDefEntegratorSecim

    'E-Defter Liste Rapor
    Public EDefListeRaporlarXml = "EDefListeRaporlar.xml", EDefListeRaporlar_FormUID As String = "ELRAP"
    Public oEDefListeRaporlar As New ClsEDefListeRaporlar

    'E-Defter Liste Rapor Değer
    Public EDefListeRaporlarDegerXml = "EDefListeRaporlarDeger.xml", EDefListeRaporlarDeger_FormUID As String = "ELRAPD"
    Public oEDefListeRaporlarDeger As New ClsEDefListeRaporlarDeger

    'E-Defter Liste Rapor Değer
    Public EDefRaporXml = "EDefRapor.xml", EDefRapor_FormUID As String = "ELRAPVERI"
    Public oEDefRapor As New ClsEDefRapor

    'E-Defter Belge Ayarlari Giris
    Public EDefBelgeAyarlariXml = "EDefBelgeAyarlari.xml", EDefBelgeAyarlari_FormUID As String = "EBELA"
    Public oEDefBelgeAyarlari As New ClsEDefBelgeAyarlari



#End Region

#Region "       ... Gentral Purpose ...     "

    Public v_RetVal, v_ErrCode As Long
    Public v_ErrMsg As String = ""
    Public addonName As String = "EDefter"
    '   Attachment Option ...
    Public ShowFolderBrowserThread As Threading.Thread
    Public BankFileName As String
    Public boolModelForm As Boolean = False
    Public boolModelFormID As String = ""
    Public sQuery As String = ""

#End Region

End Module
