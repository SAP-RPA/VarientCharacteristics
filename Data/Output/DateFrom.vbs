
'-Begin-----------------------------------------------------------------

'-Sub Main--------------------------------------------------------------
Sub Main()

  Dim SapGuiAuto, app, connection, session

  Set SapGuiAuto = GetObject("SAPGUI")
  If Not IsObject(SapGuiAuto) Then
    Exit Sub
  End If

  Set app = SapGuiAuto.GetScriptingEngine
  If Not IsObject(app) Then
    Exit Sub
  End If

  Set connection = app.Children(0)
  If Not IsObject(connection) Then
    Exit Sub
  End If

  If connection.DisabledByServer = True Then
    Exit Sub
  End If

  Set session = connection.Children(0)
  If Not IsObject(session) Then
    Exit Sub
  End If

  If session.Info.IsLowSpeedConnection = True Then
    Exit Sub
  End If

session.findById("wnd[0]").maximize
session.findById("wnd[3]/usr/cntlGRID1/shellcont/shell").text = "CHROME  PINK"


End Sub

'-Main------------------------------------------------------------------
Main

'-End-------------------------------------------------------------------
