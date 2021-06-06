Attribute VB_Name = "SendMailsCode"
Sub send_email()
    Dim emailEnvio As String
    emailEnvio = QuienEnvia()
    Dim perfil As String
    perfil = QuePerfil()
    
    Dim olApp As Object
    Set olApp = CreateObject("Outlook.Application")
'   Set olApp = New Outlook.Application
    
    Dim oAccount As Object
    Set oAccount = BuscarCuenta(emailEnvio, perfil, olApp)
    If oAccount Is Nothing Then
        MsgBox ("No encontró la cuenta " & emailEnvio)
        Exit Sub
    End If
    
    Dim sh As Worksheet
    Set sh = Worksheets("Alertas")
    
    Dim each_row As Integer
    Dim last_row As Integer
    last_row = Application.WorksheetFunction.CountA(sh.Range("A:A"))

    For each_row = 2 To last_row
        Dim msg As Object
          Set msg = olApp.createitem(0)
          msg.To = sh.Range("A" & each_row).Value
          nombre_contacto = sh.Range("B" & each_row).Value
          cuit = sh.Range("C" & each_row).Value
          razon_social = sh.Range("D" & each_row).Value
          msg.cc = sh.Range("E" & each_row).Value
          
          Subject = Personalizar(sh.Range("F" & each_row).Value, nombre_contacto, cuit, razon_social)
          msg.Subject = Subject
          Content = Personalizar(sh.Range("G" & each_row).Value, nombre_contacto, cuit, razon_social)
          msg.body = Content
          
          Status = sh.Range("I" & each_row).Value
          If Status <> "Sent" Then
            If sh.Range("H" & each_row).Value <> "" Then
                msg.attachments.Add sh.Range("H" & each_row).Value
            End If
            msg.send
            Cells(each_row, 9).Value = "Sent"
          End If
    Next each_row
End Sub

Private Function Personalizar(ByVal text As String, ByVal nc As String, ByVal cuit As String, ByVal rs As String) As String
    Personalizar = Replace(Replace(Replace(text, _
             "<NOMBRE CONTACTO>", nc), _
             "<CUIT>", cuit), _
             "<RAZON SOCIAL>", rs)
End Function

Private Function BuscarCuenta(ByVal email As String, ByVal perfil As String, ByVal olApp As Object) As Object

    ' Get a session object.
    Dim olNs As Object
    Set olNs = olApp.getNamespace("MAPI")
    olNs.Logon perfil, , True
    
    ' Create an instance of the Inbox folder.
    ' If Outlook is not already running, this has the side
    ' effect of initializing MAPI.
'    Dim mailFolder As Outlook.Folder
'    Set mailFolder = olNs.GetDefaultFolder(olFolderInbox)
    
    Dim oAccount As Object
    For Each oAccount In olApp.Session.Accounts
        Debug.Print "oAccount: " & oAccount
        If oAccount = email Then
            Set BuscarCuenta = oAccount
            Exit Function
        End If
    Next
End Function

Private Function QuienEnvia() As String
    Dim sh As Worksheet
    Set sh = Worksheets("Configuración")
    QuienEnvia = sh.Range("B2").Value
End Function
Private Function QuePerfil() As String
    Dim sh As Worksheet
    Set sh = Worksheets("Configuración")
    QuePerfil = sh.Range("B3").Value
End Function



