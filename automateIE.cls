VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "automateIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public ie As InternetExplorer

'limpa o histórico do IE
Public Sub limparIE()
  Call Clear_Temp_Files
  'Call Clear_Cookies
  'Call Clear_History
  'Call Clear_Form_Data
  'Call Clear_Saved_Passwords
  'Call Clear_All
  'Call Clear_Clear_Add_ons_Settings
  Call Sleep(5000)
End Sub
Private Sub Clear_Temp_Files()
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8 "
End Sub
Private Sub Clear_Cookies()
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2"
End Sub
Private Sub Clear_History()
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1"
End Sub
Private Sub Clear_Form_Data()
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16"
End Sub
 Private Sub Clear_Saved_Passwords()
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32"
End Sub
Private Sub Clear_All()
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255"
End Sub
Private Sub Clear_Clear_Add_ons_Settings()
  Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351"
End Sub

'fecha qualquer instancia do IE
Public Sub fechaIE(ie As InternetExplorer)
  Dim objShell As Object, AllWindows As Object, Count As Long, window As Object
  '---------------------------------------------------
  Set objShell = CreateObject("Shell.Application")
  '================================
  Do
    Set AllWindows = objShell.Windows
    Count = 0
    '----------------------------------
    For Each window In AllWindows
      If window.Name = "Internet Explorer" Then
        Sleep (500)
        On Error Resume Next
        window.Quit
        Count = Count + 1 'To close a shell window 'window.close
      End If
    Next
    If Count = 0 Then Exit Do
  Loop
  '================================
  If Not ie Is Nothing Then
    Set ie = Nothing
  End If
  '--------------------------------
  Call Sleep(3 * 1000)
End Sub

'cria uma nova instância do IE
Public Sub criaIE()
  Call Sleep(5 * 1000)
  Set ie = New InternetExplorer
  Call carregaIE(ie) 'propriedades do IE
End Sub

'carrega o IE
Private Sub carregaIE(ByRef ie As InternetExplorer)
  ie.Visible = True
  ie.Silent = False
  ie.Top = 0
  ie.Left = 0
  ie.Width = 400
  ie.Height = 780
  ie.Navigate "about:blank"
  Call AguardaIE(ie)
End Sub

'vai para a pagina
Public Sub gotoPage(webUrl As String, isPageHeavy As Boolean)
  ie.Visible = False
  ie.Visible = True
  ie.Navigate webUrl
  'If waitTenSeconds Then Call Sleep(10000)  '- se a página for pesada
  Call AguardaIE(ie)
  Call AguardaIE_longLogin(isPageHeavy)  'Call AguardaIE(ie)
End Sub

'aguarda o IE carregar
Private Sub AguardaIE(ByRef ie As Object)
On Error GoTo tryAgain
Again:
  Do
    DoEvents
  Loop Until ie.ReadyState = READYSTATE_COMPLETE And ie.Busy = False
  'While ie.Busy Or ie.ReadyState <> READYSTATE_COMPLETE: DoEvents: Wend
Exit Sub
tryAgain:
    Set ie = GetIE 'get new/existing IE object
    GoTo Again
End Sub

'aguarda o IE Bradesco login
Private Sub AguardaIE_longLogin(isPageHeavy As Boolean)
On Error GoTo tryAgain
  Dim ie As InternetExplorer, nrTentativas As Long
Again:
  '--------------------------------------
  Do
    '___________________
    If isPageHeavy Then
      Sleep 5000 'se a página for pesada, demora mais tempo
    Else
      Sleep 2500
    End If
    '-------------------------
    Set ie = GetIE
    DoEvents
    ie.Width = 800
    If ie.Width = 800 Then 'descobre
      Exit Sub
    Else
      ie.Width = 400
    End If
    nrTentativas = nrTentativas + 1
  Loop Until nrTentativas = 10 'loop infinito
  '--------------------------------------
  Call ieTamanhoFull   'monta a janela com tamanho grande
Exit Sub
tryAgain:
    Set ie = GetIE 'get new/existing IE object
    GoTo Again
End Sub

'clica na classe
Public Sub classeClique(classe As String, item As Long)
  Call ie.Document.getElementsByClassName(classe).item(item).Click 'clica no botão entrar
  Call AguardaIE(ie)
  Call AguardaIE_longLogin(False)
End Sub

'preenche o campo de 'ID'
Public Sub preencheID(id As String, valor As String)
  ie.Document.getElementById(id).Focus
  ie.Document.getElementById(id).Value = valor
  ie.Document.getElementById(id).Focus
  Call Sleep(1500)
End Sub

'clica no id
Public Sub idClique(id As String, Optional isPageHeavy As Boolean)
  ie.Document.getElementById(id).Focus
  Call Sleep(1000)
  ie.Document.getElementById(id).Click
  Call AguardaIE(ie)
  Call AguardaIE_longLogin(isPageHeavy)
End Sub

'monta a janela com tamanho grande
Private Sub ieTamanhoFull()
  ie.Width = 1600
  ie.Height = 1200
End Sub

'pega o IE novamente
Private Function GetIE() As Object
On Error GoTo sai
'return an object for the open Internet Explorer window, or create new one
  For Each GetIE In CreateObject("Shell.Application").Windows() 'Loop to find
    If (Not GetIE Is Nothing) And GetIE.Name = "Internet Explorer" Then Exit For 'Found!
  Next GetIE
  'If criaNovoIE Then Set GetIE = CreateObject("InternetExplorer.Application") 'Create
  Set GetIE = GetIE
  GetIE.Visible = True 'Make IE window visible
sai:
End Function

Private Sub Class_Initialize()
  Call fechaIE(ie) 'fecha todas as instâncias do IE
  Call limparIE 'limpa o histórico do IE
  Call criaIE   'cria uma nova instância do IE
End Sub
