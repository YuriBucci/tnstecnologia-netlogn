'::#########################################################################
'::|_   _| \ | / ___|  |_   _|__  ___ _ __   ___ | | ___   __ _(_)
'::  | | |  \| \___ \    | |/ _ \/ __| _ \ / _  \| |/ _ \ / _  | |/ _  |
'::  | | | |\  |___) |   | |  __/ (__| | | | (_) | | (_) | (_| | | (_| |
'::  |_| |_| \_|____/    |_|\___|\___|_| |_|\___/|_|\___/ \__, |_|\__,_|
'::                                                       |___/
'::#########################################################################
'::Title                  : Script Logon
'::Description            : Script Logon Padrão TNS Tecnologia
'::Author                 : Yuri Bucci
'::Facebook               : https://www.facebook.com/YuriBucci
'::Site                   : www.tnsinformatica.com.br
'::Date                   : 26/02/2019
'::Version                : 5.0
'::'#########################################################################


On error Resume Next

Err.clear 0


Set oShell = CreateObject("Shell.Application")

Set WshShell = WScript.CreateObject("WScript.Shell")

oShell.ShellExecute "\\servidor-dc\netlogon\logon.bat", , , , 0

'==================ATALHOS NO DESKTOP==================

Set objShell=Wscript.CreateObject("Wscript.shell")
strDesktopFolder=objShell.SpecialFolders("Desktop") & _
"\"
Set objShortcut=objShell.CreateShortcut(strDesktopFolder & _
"SERVIDOR.lnk")
objShortCut.TargetPath = "\\servidor-dc\"
objShortCut.Description = "Servidor-DC"
objShortCut.Save


'==================BOAS VINDAS AO USUÁRIO==================

'Boas Vindas Ao Usuario

Set objUser = WScript.CreateObject("WScript.Network")
wuser=objUser.UserName
If Time <= "12:00:00" Then
MsgBox ("Bom Dia "+Wuser+", você acaba de ingressar na rede corporativa da Setter Embalagens por favor respeite as políticas de segurança e bom trabalho!")
ElseIf Time >= "12:00:01" And Time <= "18:00:00" Then
MsgBox ("Boa Tarde "+Wuser+", você acaba de ingressar na rede corporativa da Setter Embalagens, por favor respeite as políticas de segurança e bom trabalho!")
Else
MsgBox ("Boa Noite "+wuser+", você acaba de ingressar na rede corporativa da Setter Embalagens, por favor respeite as políticas de segurança e bom trabalho!")
End If
