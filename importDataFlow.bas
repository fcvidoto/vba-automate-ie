Attribute VB_Name = "importDataFlow"
Option Explicit

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'faz o login no site do power bi
Sub automacaoBI()
  Dim ie As InternetExplorer, z As Long, menu(1 To 3, 1 To 2) As Variant, menuF(1 To 1, 1 To 2) As Variant
  Dim usuario As String, senha As String, webUrl As String
  Dim a As automateIE: Set a = New automateIE
  '-------------------------------------------------------
  usuario = "*********"
  senha = "***********"
  webUrl = "https://app.powerbi.com/"
  '-------------------------------------------------------
  Call gotoPowerBIPage(ie, webUrl, usuario, senha, a) 'loga no power BI
  Call a.gotoPage("**************", True)  'vai para a pagina dos dataflows
  '-------------------------------------------------------
  'dataflow
      menu(1, 1) = 1: menu(1, 2) = "vba - demais tabelas - base teste"
      menu(2, 1) = 2: menu(2, 2) = "vba- series"
      menu(3, 1) = 3: menu(3, 2) = "principais tabelas - base de teste"
      '--//
      For z = 1 To UBound(menu) '<---- (3) Quantidade de dataflows(cliques)
        Call a.gotoPage("**************", False) 'vai para a pagina dos dataflows
        Set ie = a.ie
        Call clickDataFlow(ie, menu(z, 2))
      Next z
  '-------------------------------------------------------
  'datasets
  Call a.classeClique("artifactTab mat-tab-link ng-star-inserted", 0)
  Call a.classeClique("mat-sort-header-button", 2) 'clica na classe
  Call a.classeClique("mat-sort-header-button", 2) 'clica na classe
  '---------------
      menuF(1, 1) = 1: menuF(1, 2) = "pnl_produto 3.8"
      '--//
      For z = 1 To UBound(menuF) '<---- (3) Quantidade de dataflows(cliques)
        Set ie = a.ie
        Call clickDataSet(ie, menuF(z, 2))
      Next z
  '****************'
  '-------------------------------------------------------
End Sub

'clica no botão do dataflow
Public Sub clickDataFlow(ByRef ie As InternetExplorer, tituloDataFlow As Variant)
On Error GoTo trata
  Dim div As Object, i As Long
  For Each div In ie.Document.getElementsByClassName("row ng-star-inserted")
    i = i + 1
    Debug.Print LCase(div.innerText)
    If LCase(div.innerText) Like "*" & tituloDataFlow & "*" Then
      'ie.Document.parentWindow.execScript "window.alert('1')"
      ie.Document.parentWindow.execScript "document.getElementsByClassName('pbi-glyph pbi-glyph-refresh').item(" & i & ");", "jscript"
      div.getElementsByTagName("button").item(0).Click 'clica no botão, refreshdata
      Exit Sub
    End If
    'Debug.Print div.innerText
  Next div
trata:
End Sub

'clica no botão do dataflow
Public Sub clickDataSet(ByRef ie As InternetExplorer, tituloDataFlow As Variant)
On Error GoTo trata
  Dim div As Object, i As Long
  For Each div In ie.Document.getElementsByClassName("row ng-star-inserted")
    i = i + 1
    If LCase(div.innerText) Like "*" & tituloDataFlow & "*" Then
      'ie.Document.parentWindow.execScript "window.alert('1')"
      ie.Document.parentWindow.execScript "document.getElementsByClassName('pbi-glyph pbi-glyph-refresh').item(" & i & ");", "jscript"
      div.getElementsByTagName("button").item(1).Click 'clica no botão, refreshdata
      Exit Sub
    End If
    'Debug.Print div.innerText
  Next div
trata:
End Sub

'loga no power BI
Public Sub gotoPowerBIPage(ByRef ie As InternetExplorer, webUrl As String, usuario As String, senha As String, a As automateIE)
  '-------------------------------------------------------
  Call a.gotoPage(webUrl, True)  'vai para a página
  '-------------------------------------------------------
  Call a.classeClique("button", 0) 'clica na classe
  Call a.preencheID("i0116", usuario) 'preenche o campo de valor ID
  Call a.idClique("idSIButton9", True) 'clica no id
  '-------------------------------------------------------
  'segunda tela de login
  Call a.preencheID("userNameInput", usuario) 'preenche o campo de valor ID
  Call a.preencheID("passwordInput", senha) 'preenche o campo de valor ID
  Call a.idClique("submitButton") 'clica no id
  '-------------------------------------------------------
  'msg de não aparecer novamente
  Call a.idClique("KmsiCheckboxField") 'clica no id // check do controle
  Call a.idClique("idSIButton9") 'clica no id // check do controle
End Sub

'vai para a pagina do dataset
Sub gotoDataSet(ie As InternetExplorer)
  'clica em dataflow
  Call ie.Document.getElementsByClassName("artifactTab mat-tab-link ng-star-inserted").item(0).Click
  'Call AguardaIE(ie)
  'Call AguardaIE_longLogin  'Call AguardaIE(ie)
End Sub
