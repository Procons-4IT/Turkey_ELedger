Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Threading
Imports System.Runtime.InteropServices

Namespace SBOPlugins.Enumerations
    Public Enum eFileDialog
        en_OpenFile = 0
        en_SaveFile = 1
    End Enum
End Namespace

Namespace CoreFrieght_Intraspeed

    ''' <summary>
    ''' Wrapper for OpenFileDialog
    ''' </summary>
    Public Class GetFileNameClass
        Implements IDisposable
#Region "The class implements FileDialog for open in front of B1 window"

        <DllImport("user32.dll")> _
        Private Shared Function GetForegroundWindow() As IntPtr
        End Function

        Private _oFileDialog As System.Windows.Forms.FileDialog

        ' Properties
        Public Property FileName() As String
            Get
                Return _oFileDialog.FileName
            End Get
            Set(value As String)
                _oFileDialog.FileName = value
            End Set
        End Property

        Public ReadOnly Property FileNames() As String()
            Get
                Return _oFileDialog.FileNames
            End Get
        End Property

        Public Property Filter() As String
            Get
                Return _oFileDialog.Filter
            End Get
            Set(value As String)
                _oFileDialog.Filter = value
            End Set
        End Property

        Public Property InitialDirectory() As String
            Get
                Return _oFileDialog.InitialDirectory
            End Get
            Set(value As String)
                _oFileDialog.InitialDirectory = value
            End Set
        End Property

        '''/ Constructor
        'public GetFileNameClass()
        '{
        '    _oFileDialog = new OpenFileDialog();
        '}

        ' Constructor
        Public Sub New(dlg As SBOPlugins.Enumerations.eFileDialog)
            Select Case CInt(dlg)
                Case 0
                    _oFileDialog = New System.Windows.Forms.OpenFileDialog()
                    Exit Select
                Case 1
                    _oFileDialog = New System.Windows.Forms.SaveFileDialog()
                    Exit Select
                Case Else
                    Throw New ApplicationException("GetFileNameClass Incorrect Parameter")
            End Select
        End Sub

        Public Sub New()

            Me.New(SBOPlugins.Enumerations.eFileDialog.en_OpenFile)
        End Sub

        ' Dispose
        Public Sub Dispose() Implements IDisposable.Dispose
            _oFileDialog.Dispose()
        End Sub

        ' Methods

        Public Sub GetFileName()
            Dim ptr As IntPtr = GetForegroundWindow()

            Dim oWindow As New WindowWrapper(ptr)

            If _oFileDialog.ShowDialog(oWindow) <> System.Windows.Forms.DialogResult.OK Then
                _oFileDialog.FileName = String.Empty
            End If
            oWindow = Nothing
        End Sub
        ' End of GetFileName
#End Region

#Region "WindowWrapper : System.Windows.Forms.IWin32Window"

        Public Class WindowWrapper
            Implements System.Windows.Forms.IWin32Window
            Private _hwnd As IntPtr

            ' Property
            Public Overridable ReadOnly Property Handle() As IntPtr Implements IWin32Window.Handle
                Get
                    Return _hwnd
                End Get
            End Property

            ' Constructor
            Public Sub New(handle As IntPtr)
                _hwnd = handle
            End Sub
        End Class
#End Region
    End Class
End Namespace