Attribute VB_Name = "ModificaForms"
Option Explicit

'---------------------------------------------------------------------------------------------------'
'Declarações das Funções API
#If VBA7 Then
Public Declare PtrSafe Function FindWindow Lib "user32" Alias _
    "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare PtrSafe Function SendMessageA Lib "user32" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare PtrSafe Function ExtractIconA Lib "shell32.dll" _
    (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare PtrSafe Function GetActiveWindow Lib "user32.dll" () As Long
Public Declare PtrSafe Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
#Else
Public Declare Function FindWindow Lib "user32" Alias _
    "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessageA Lib "user32" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function ExtractIconA Lib "shell32.dll" (ByVal hInst As Long, ByVal lpszExeFileName As String, _
    ByVal nIconIndex As Long) As Long
Public Declare Function GetActiveWindow Lib "user32.dll" () As Long
Public Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal hWnd As Long, ByVal crey As Byte,ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
#End If
''---------------------------------------------------------------------------------------------------'
'Declarações das variáveis
Public hWnd As Long 'variável que identifica o Formulário ativo
'---------------------------------------------------------------------------------------------------'
'Declarações das constantes
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_STYLE = (-16)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H1
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40
Public Const SW_SHOW = 5
Public Const HWND_TOP = 0
Public Const LWA_ALPHA = &H2&
Public Const WM_SETICON = &H80
'---------------------------------------------------------------------------------------------------'


'---------------------- Rotina para personalizar userform ativo -------------------------------'
' verifica o userform ativo e adiciona botões de maximizar e minimizar, transparência

Public Sub PersonalizaForm()

    Dim wStyle As Long
    Dim xStyle As Long
    Dim bOpacity As Byte
    
    'pega o handle da janela ativa
    hWnd = GetActiveWindow
    bOpacity = 225 ' define opacidade da janela ativa
    
    'recupera os estilos da janela ativa
    wStyle = GetWindowLong(hWnd, GWL_STYLE)
    
    'modifica as configurações de estilo da janela
    wStyle = wStyle Or WS_MINIMIZEBOX 'adiciona o botão minimizar
    wStyle = wStyle Or WS_MAXIMIZEBOX 'adiciona o botão maximizar
    wStyle = wStyle Or WS_THICKFRAME 'adiciona uma borda de dimensionamento
    
    'aplica o estilo revisado
    Call SetWindowLong(hWnd, GWL_STYLE, wStyle)
    
    'recupera os estilos estendidos da janela ativa
    xStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    'modifica as configurações de estilo estendido da janela
    xStyle = xStyle Or WS_EX_LAYERED 'muda a opacidade
    xStyle = xStyle Or WS_EX_APPWINDOW 'adiciona janela na barra de tarefas
    
    'aplicar o estilo estendido revisado
    Call SetWindowLong(hWnd, GWL_EXSTYLE, xStyle)
    Call SetLayeredWindowAttributes(hWnd, 0, bOpacity, LWA_ALPHA)
    Call SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, _
    SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_HIDEWINDOW)
    Call SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, _
    SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW)

End Sub
'--------------------- Rotina para adicionar ícone na barra de título -------------------------------'
Public Sub AddIcon_BarraTitulo(hWnd, strIconBmpFile As String)

    Dim fLen As Long
    If Len(Dir(strIconBmpFile)) <> 0 Then
        fLen = ExtractIconA(0, strIconBmpFile, 0)
        SendMessageA FindWindow(vbNullString, hWnd.Caption), _
        WM_SETICON, False, fLen
    Else
        Exit Sub
    End If
    
End Sub

