@ECHO OFF

::#########################################################################
::|_   _| \ | / ___|  |_   _|__  ___ _ __   ___ | | ___   __ _(_)
::  | | |  \| \___ \    | |/ _ \/ __| _ \ / _  \| |/ _ \ / _  | |/ _  |
::  | | | |\  |___) |   | |  __/ (__| | | | (_) | | (_) | (_| | | (_| |
::  |_| |_| \_|____/    |_|\___|\___|_| |_|\___/|_|\___/ \__, |_|\__,_|
::                                                       |___/
::#########################################################################
::Title                  : Script Logon Único
::Description            : Script Logon Único Padrão TNS Tecnologia
::Author                 : Yuri Bucci
::Facebook               : https://www.facebook.com/YuriBucci
::Site                   : www.tnsinformatica.com.br
::Date                   : 26/02/2019
::Version                : 5.0
::'#########################################################################

::---------------------------------------------------------------------------------------
::---------------------------------------------------------------------------------------

:: BGINFO ::


\\servidor-dc\netlogon\Bginfo.exe \\servidor-dc\netlogon\padrao.bgi /timer:0 /nolicprompt /silent

:: DELETA TODAS AS CONEXÃOS MAPEADAS E MAPEIA PASTA DO USUÁRIO ::

net use * /del /yes

net use H: \\servidor-dc\Usuarios\%username%


::---------------------------------------------------------------------------------------
::---------------------------------------------------------------------------------------

:: MAPEAMENTO DE REDE BASEADO EM GRUPO (CASO NECESSARIO) ::


set GRUPO=G_GERAL

net user %username% /domain | find /i "%GRUPO%" 1>nul

if %errorlevel%==0 (

net use G: \\servidor-dc\Geral

)

set GRUPO=G_FINANCEIRO

net user %username% /domain | find /i "%GRUPO%" 1>nul

if %errorlevel%==0 (

net use F: \\servidor-dc\Financeiro

)


::---------------------------------------------------------------------------------------
::---------------------------------------------------------------------------------------

:: MAPEAMENTO DE IMPRESSORA BASEADO EM GRUPO ::

set GRUPO=IMP_PLANEJAMENTO

net user %username% /domain | find /i "%GRUPO%" 1>nul

if %errorlevel%==0 (


:: ADICIONA/ATUALIZA PORTA DA IMPRESSORA ::
cscript \\servidor-dc\netlogon\prnport.vbs -a -r IP_192.168.0.176 -h 192.168.0.176 -o raw

:: INSTALA DRIVER DA IMPRESSORA ::
cscript \\servidor-dc\netlogon\prndrvr.vbs -a -m "Kyocera ECOSYS M2035dn KX" -h "\\servidor-dc\Suporte TNS\98 - IMPRESSORAS\KYOCERA\KX7.4_v7.4.0830\64bit\" -i "\\servidor-dc\Suporte TNS\98 - IMPRESSORAS\KYOCERA\KX7.4_v7.4.0830\64bit\OEMSETUP.INF"

:: ADICIONA A IMPRESSORA ::
cscript \\servidor-dc\netlogon\prnmngr.vbs -a -p "IMPRESSORA_PLANEJAMENTO" -m "Kyocera ECOSYS M2035dn KX" -r "IP_192.168.0.176"

:: SETA A IMPRESSORA COMO PADRAO ::
cscript \\servidor-dc\netlogon\prnmngr.vbs -a -p "IMPRESSORA_PLANEJAMENTO" -m "Kyocera ECOSYS M2035dn KX" -r "IP_192.168.0.176" -t

)

set GRUPO=IMP_ADMINISTRATIVO

net user %username% /domain | find /i "%GRUPO%" 1>nul

if %errorlevel%==0 (

:: ADICIONA/ATUALIZA PORTA DA IMPRESSORA ::
cscript \\servidor-dc\netlogon\prnport.vbs -a -r IP_192.168.0.59 -h 192.168.0.59 -o raw

:: INSTALA DRIVER DA IMPRESSORA ::
cscript \\servidor-dc\netlogon\prndrvr.vbs -a -m "Kyocera ECOSYS M2035dn KX" -h "\\servidor-dc\Suporte TNS\98 - IMPRESSORAS\KYOCERA\KX7.4_v7.4.0830\64bit\" -i "\\servidor-dc\Suporte TNS\98 - IMPRESSORAS\KYOCERA\KX7.4_v7.4.0830\64bit\OEMSETUP.INF"

:: ADICIONA A IMPRESSORA ::
cscript \\servidor-dc\netlogon\prnmngr.vbs -a -p "IMPRESSORA_ADMINISTRATIVO" -m "Kyocera ECOSYS M2035dn KX" -r "IP_192.168.0.59"

)