'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation. All rights reserved.
'
' Abstract:
' prnmngr.vbs - printer script for WMI on Windows 
'     used to add, delete, and list printers and connections
'     also for getting and setting the default printer
'
' Usage:
' prnmngr [-adxgtl?][co] [-s server][-p printer][-m driver model][-r port]
'                       [-u user name][-w password]
'
' Examples:
' prnmngr -a -p "printer" -m "driver" -r "lpt1:"
' prnmngr -d -p "printer" -s server
' prnmngr -ac -p "\\server\printer"
' prnmngr -d -p "\\server\printer"
' prnmngr -x -s server
' prnmngr -l -s server
' prnmngr -g
' prnmngr -t -p "printer"
'
'----------------------------------------------------------------------

option explicit

'
' Debugging trace flags, to enable debug output trace message
' change gDebugFlag to true.
'
const kDebugTrace = 1
const kDebugError = 2
dim   gDebugFlag

gDebugFlag = false

'
' Operation action values.
'
const kActionUnknown           = 0
const kActionAdd               = 1
const kActionAddConn           = 2
const kActionDel               = 3
const kActionDelAll            = 4
const kActionDelAllCon         = 5
const kActionDelAllLocal       = 6
const kActionList              = 7
const kActionGetDefaultPrinter = 8
const kActionSetDefaultPrinter = 9

const kErrorSuccess            = 0
const KErrorFailure            = 1

const kFlagCreateOnly          = 2

const kNameSpace               = "root\cimv2"

'
' Generic strings
'
const L_Empty_Text                 = ""
const L_Space_Text                 = " "
const L_Error_Text                 = "Erro"
const L_Success_Text               = "Êxito"
const L_Failed_Text                = "Falha"
const L_Hex_Text                   = "0x"
const L_Printer_Text               = "Impressora"
const L_Operation_Text             = "Operação"
const L_Provider_Text              = "Provedor"
const L_Description_Text           = "Descrição"
const L_Debug_Text                 = "Depurar:"
const L_Connection_Text            = "Conexão"

'
' General usage messages
'
const L_Help_Help_General01_Text   = "Uso: prnmngr [-adxgtl?][c] [-s servidor][-p impressora][-m modelo de driver]"
const L_Help_Help_General02_Text   = "               [-r port][-u user name][-w password]"
const L_Help_Help_General03_Text   = "Argumentos:"
const L_Help_Help_General04_Text   = "-a     - adicionar impressora local"
const L_Help_Help_General05_Text   = "-ac    - adicionar conexão de impressora"
const L_Help_Help_General06_Text   = "-d     - excluir impressora"
const L_Help_Help_General07_Text   = "-g     - obter a impressora padrão"
const L_Help_Help_General08_Text   = "-l     - listar impressoras"
const L_Help_Help_General09_Text   = "-m     - modelo de driver"
const L_Help_Help_General10_Text   = "-p     - nome da impressora"
const L_Help_Help_General11_Text   = "-r     - nome da porta"
const L_Help_Help_General12_Text   = "-s     - nome do servidor"
const L_Help_Help_General13_Text   = "-t     - definir a impressora padrão"
const L_Help_Help_General14_Text   = "-u     - nome do usuário"
const L_Help_Help_General15_Text   = "-w     - senha"
const L_Help_Help_General16_Text   = "-x     - excluir todas as impressoras"
const L_Help_Help_General17_Text   = "-xc    - excluir todas as conexões de impressora"
const L_Help_Help_General18_Text   = "-xo    - excluir todas as impressoras locais"
const L_Help_Help_General19_Text   = "-?     - exibir uso do comando"
const L_Help_Help_General20_Text   = "Exemplos:"
const L_Help_Help_General21_Text   = "prnmngr -a -p ""printer"" -m ""driver"" -r ""lpt1:"""
const L_Help_Help_General22_Text   = "prnmngr -d -p ""printer"" -s server"
const L_Help_Help_General23_Text   = "prnmngr -ac -p ""\\server\printer"""
const L_Help_Help_General24_Text   = "prnmngr -d -p ""\\server\printer"""
const L_Help_Help_General25_Text   = "prnmngr -x -s server"
const L_Help_Help_General26_Text   = "prnmngr -xo"
const L_Help_Help_General27_Text   = "prnmngr -l -s server"
const L_Help_Help_General28_Text   = "prnmngr -g"
const L_Help_Help_General29_Text   = "prnmngr -t -p ""\\server\printer"""

'
' Messages to be displayed if the scripting host is not cscript
'
const L_Help_Help_Host01_Text      = "Este script deve ser executado a partir do prompt de comando usando CScript.exe."
const L_Help_Help_Host02_Text      = "Por exemplo: CScript script.vbs argumentos"
const L_Help_Help_Host03_Text      = ""
const L_Help_Help_Host04_Text      = "Para definir CScript como o aplicativo padrão para a execução de arquivos .VBS, execute:"
const L_Help_Help_Host05_Text      = "     CScript //H:CScript //S"
const L_Help_Help_Host06_Text      = "Você poderá em seguida executar ""script.vbs argumentos"" sem incluir CScript antes to script."

'
' General error messages
'
const L_Text_Error_General01_Text  = "O host de script não pôde ser determinado."
const L_Text_Error_General02_Text  = "Não é possível analisar a linha de comando."
const L_Text_Error_General03_Text  = "Código de erro do Win32"

'
' Miscellaneous messages
'
const L_Text_Msg_General01_Text    = "Impressora adicionada"
const L_Text_Msg_General02_Text    = "Não é possível adicionar impressora"
const L_Text_Msg_General03_Text    = "Conexão de impressora adicionada"
const L_Text_Msg_General04_Text    = "Não foi possível adicionar conexão com a impressora"
const L_Text_Msg_General05_Text    = "Impressora excluída"
const L_Text_Msg_General06_Text    = "Não foi possível excluir a impressora"
const L_Text_Msg_General07_Text    = "Tentando excluir impressora"
const L_Text_Msg_General08_Text    = "Não é possível excluir impressoras"
const L_Text_Msg_General09_Text    = "Número de impressoras locais e conexões enumeradas"
const L_Text_Msg_General10_Text    = "Número de impressoras locais e conexões excluídas"
const L_Text_Msg_General11_Text    = "Não é possível enumerar impressoras"
const L_Text_Msg_General12_Text    = "A impressora padrão é"
const L_Text_Msg_General13_Text    = "Não foi possível obter a impressora padrão"
const L_Text_Msg_General14_Text    = "Não foi possível definir a impressora padrão"
const L_Text_Msg_General15_Text    = "Agora a impressora padrão é"
const L_Text_Msg_General16_Text    = "Número de conexões de impressora enumeradas"
const L_Text_Msg_General17_Text    = "Número de conexões de impressora excluídas"
const L_Text_Msg_General18_Text    = "Número de impressoras locais enumeradas"
const L_Text_Msg_General19_Text    = "Número de impressoras locais excluídas"

'
' Printer properties
'
const L_Text_Msg_Printer01_Text    = "Nome do servidor"
const L_Text_Msg_Printer02_Text    = "Nome da impressora"
const L_Text_Msg_Printer03_Text    = "Nome de compartilhamento"
const L_Text_Msg_Printer04_Text    = "Nome do driver"
const L_Text_Msg_Printer05_Text    = "Nome da porta"
const L_Text_Msg_Printer06_Text    = "Comentário"
const L_Text_Msg_Printer07_Text    = "Local"
const L_Text_Msg_Printer08_Text    = "Arquivo separador"
const L_Text_Msg_Printer09_Text    = "Processador de impressão"
const L_Text_Msg_Printer10_Text    = "Tipo de dados"
const L_Text_Msg_Printer11_Text    = "Parâmetros"
const L_Text_Msg_Printer12_Text    = "Atributos"
const L_Text_Msg_Printer13_Text    = "Prioridade"
const L_Text_Msg_Printer14_Text    = "Prioridade padrão"
const L_Text_Msg_Printer15_Text    = "Hora inicial"
const L_Text_Msg_Printer16_Text    = "Até hora"
const L_Text_Msg_Printer17_Text    = "Contagem de trabalhos"
const L_Text_Msg_Printer18_Text    = "Média de páginas por minuto"
const L_Text_Msg_Printer19_Text    = "Status da impressora"
const L_Text_Msg_Printer20_Text    = "Status de impressora estendido"
const L_Text_Msg_Printer21_Text    = "Estado de erro detectado"
const L_Text_Msg_Printer22_Text    = "Estado de erro estendido detectado"


'
' Printer status
'
const L_Text_Msg_Status01_Text     = "Outros"
const L_Text_Msg_Status02_Text     = "Desconhecido"
const L_Text_Msg_Status03_Text     = "Ocioso"
const L_Text_Msg_Status04_Text     = "Imprimindo"
const L_Text_Msg_Status05_Text     = "Em aquecimento"
const L_Text_Msg_Status06_Text     = "Impressão parada"
const L_Text_Msg_Status07_Text     = "Offline"
const L_Text_Msg_Status08_Text     = "Em pausa"
const L_Text_Msg_Status09_Text     = "Erro"
const L_Text_Msg_Status10_Text     = "Ocupada"
const L_Text_Msg_Status11_Text     = "Não disponível"
const L_Text_Msg_Status12_Text     = "Aguardando"
const L_Text_Msg_Status13_Text     = "Processando"
const L_Text_Msg_Status14_Text     = "Inicializando"
const L_Text_Msg_Status15_Text     = "Economia de energia"
const L_Text_Msg_Status16_Text     = "Exclusão pendente"
const L_Text_Msg_Status17_Text     = "E/S ativa"
const L_Text_Msg_Status18_Text     = "Alimentação manual"
const L_Text_Msg_Status19_Text     = "Sem erros"
const L_Text_Msg_Status20_Text     = "Pouco papel"
const L_Text_Msg_Status21_Text     = "Sem papel"
const L_Text_Msg_Status22_Text     = "Toner baixo"
const L_Text_Msg_Status23_Text     = "Sem toner"
const L_Text_Msg_Status24_Text     = "Porta aberta"
const L_Text_Msg_Status25_Text     = "Obstruído"
const L_Text_Msg_Status26_Text     = "Serviço solicitado"
const L_Text_Msg_Status27_Text     = "Bandeja de saída cheia"
const L_Text_Msg_Status28_Text     = "Problema no papel"
const L_Text_Msg_Status29_Text     = "Não é possível imprimir a página"
const L_Text_Msg_Status30_Text     = "Intervenção do usuário necessária"
const L_Text_Msg_Status31_Text     = "Memória insuficiente"
const L_Text_Msg_Status32_Text     = "Servidor desconhecido"

'
' Debug messages
'
const L_Text_Dbg_Msg01_Text        = "Na função AddPrinter"
const L_Text_Dbg_Msg02_Text        = "Na função AddPrinterConnection"
const L_Text_Dbg_Msg03_Text        = "Na função DelPrinter"
const L_Text_Dbg_Msg04_Text        = "Na função DelAllPrinters"
const L_Text_Dbg_Msg05_Text        = "Na função ListPrinters"
const L_Text_Dbg_Msg06_Text        = "Na função GetDefaultPrinter"
const L_Text_Dbg_Msg07_Text        = "Na função SetDefaultPrinter"
const L_Text_Dbg_Msg08_Text        = "Na função ParseCommandLine"

main

'
' Main execution starts here
'
sub main

    dim iAction
    dim iRetval
    dim strServer
    dim strPrinter
    dim strDriver
    dim strPort
    dim strUser
    dim strPassword

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(L_Help_Help_Host01_Text & vbCRLF & L_Help_Help_Host02_Text & vbCRLF & _
                          L_Help_Help_Host03_Text & vbCRLF & L_Help_Help_Host04_Text & vbCRLF & _
                          L_Help_Help_Host05_Text & vbCRLF & L_Help_Help_Host06_Text & vbCRLF)

        wscript.quit

    end if

    '
    ' Get command line parameters
    '
    iRetval = ParseCommandLine(iAction, strServer, strPrinter, strDriver, strPort, strUser, strPassword)

    if iRetval = kErrorSuccess then

        select case iAction

            case kActionAdd
                 iRetval = AddPrinter(strServer, strPrinter, strDriver, strPort, strUser, strPassword)

            case kActionAddConn
                 iRetval = AddPrinterConnection(strPrinter, strUser, strPassword)

            case kActionDel
                 iRetval = DelPrinter(strServer, strPrinter, strUser, strPassword)

            case kActionDelAll
                 iRetval = DelAllPrinters(kActionDelAll, strServer, strUser, strPassword)

            case kActionDelAllCon
                 iRetval = DelAllPrinters(kActionDelAllCon, strServer, strUser, strPassword)

            case kActionDelAllLocal
                 iRetval = DelAllPrinters(kActionDelAllLocal, strServer, strUser, strPassword)

            case kActionList
                 iRetval = ListPrinters(strServer, strUser, strPassword)

            case kActionGetDefaultPrinter
                 iRetval = GetDefaultPrinter(strUser, strPassword)

            case kActionSetDefaultPrinter
                 iRetval = SetDefaultPrinter(strPrinter, strUser, strPassword)

            case kActionUnknown
                 Usage(true)
                 exit sub

            case else
                 Usage(true)
                 exit sub

        end select

    end if

end sub

'
' Add a printer with minimum settings. Use prncnfg.vbs to
' set the complete configuration of a printer
'
function AddPrinter(strServer, strPrinter, strDriver, strPort, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg01_Text
    DebugPrint kDebugTrace, L_Text_Msg_Printer01_Text & L_Space_Text & strServer
    DebugPrint kDebugTrace, L_Text_Msg_Printer02_Text & L_Space_Text & strPrinter
    DebugPrint kDebugTrace, L_Text_Msg_Printer04_Text & L_Space_Text & strDriver
    DebugPrint kDebugTrace, L_Text_Msg_Printer05_Text & L_Space_Text & strPort

    dim oPrinter
    dim oService
    dim iRetval

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oPrinter = oService.Get("Win32_Printer").SpawnInstance_

    else

        AddPrinter = kErrorFailure

        exit function

    end if

    oPrinter.DriverName = strDriver
    oPrinter.PortName   = strPort
    oPrinter.DeviceID   = strPrinter

    oPrinter.Put_(kFlagCreateOnly)

    if Err.Number = kErrorSuccess then

        wscript.echo L_Text_Msg_General01_Text & L_Space_Text & strPrinter

        iRetval = kErrorSuccess

    else

        wscript.echo L_Text_Msg_General02_Text & L_Space_Text & strPrinter & L_Space_Text & L_Error_Text _
                     & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

        iRetval = kErrorFailure

    end if

    AddPrinter = iRetval

end function

'
' Add a printer connection
'
function AddPrinterConnection(strPrinter, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg02_Text

    dim oPrinter
    dim oService
    dim iRetval
    dim uResult

    '
    ' Initialize return value
    '
    iRetval = kErrorFailure

    '
    ' We connect to the local server
    '
    if WmiConnect("", kNameSpace, strUser, strPassword, oService) then

        set oPrinter = oService.Get("Win32_Printer")

    else

        AddPrinterConnection = kErrorFailure

        exit function

    end if

    '
    ' Check if Get was successful
    '
    if Err.Number = kErrorSuccess then

        '
        ' The Err object indicates whether the WMI provider reached the execution
        ' of the function that adds a printer connection. The uResult is the Win32
        ' error code returned by the static method that adds a printer connection
        '
        uResult = oPrinter.AddPrinterConnection(strPrinter)

        if Err.Number = kErrorSuccess then

            if uResult = kErrorSuccess then

                wscript.echo L_Text_Msg_General03_Text & L_Space_Text & strPrinter

                iRetval = kErrorSuccess

            else

                wscript.echo L_Text_Msg_General04_Text & L_Space_Text & L_Text_Error_General03_Text _
                             & L_Space_text & uResult

            end if

        else

            wscript.echo L_Text_Msg_General04_Text & L_Space_Text & strPrinter & L_Space_Text _
                         & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                         & Err.Description

        end if

    else

        wscript.echo L_Text_Msg_General04_Text & L_Space_Text & strPrinter & L_Space_Text _
                     & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                     & Err.Description

    end if

    AddPrinterConnection = iRetval

end function

'
' Delete a printer or a printer connection
'
function DelPrinter(strServer, strPrinter, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg03_Text
    DebugPrint kDebugTrace, L_Text_Msg_Printer01_Text & L_Space_Text & strServer
    DebugPrint kDebugTrace, L_Text_Msg_Printer02_Text & L_Space_Text & strPrinter

    dim oService
    dim oPrinter
    dim iRetval

    iRetval = kErrorFailure

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set oPrinter = oService.Get("Win32_Printer.DeviceID='" & strPrinter & "'")

    else

        DelPrinter = kErrorFailure

        exit function

    end if

    '
    ' Check if Get was successful
    '
    if Err.Number = kErrorSuccess then

        oPrinter.Delete_

        if Err.Number = kErrorSuccess then

            wscript.echo L_Text_Msg_General05_Text & L_Space_Text & strPrinter

            iRetval = kErrorSuccess

        else

            wscript.echo L_Text_Msg_General06_Text & L_Space_Text & strPrinter & L_Space_Text _
                         & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) _
                         & L_Space_Text & Err.Description

            '
            ' Try getting extended error information
            '
            call LastError()

        end if

    else

        wscript.echo L_Text_Msg_General06_Text & L_Space_Text & strPrinter & L_Space_Text _
                     & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) _
                     & L_Space_Text & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

    end if

    DelPrinter = iRetval

end function

'
' Delete all local printers and connections on a machine
'
function DelAllPrinters(kAction, strServer, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg04_Text

    dim Printers
    dim oPrinter
    dim oService
    dim iResult
    dim iTotal
    dim iTotalDeleted
    dim strPrinterName
    dim bDelete
    dim bConnection
    dim strTemp

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set Printers = oService.InstancesOf("Win32_Printer")

    else

        DelAllPrinters = kErrorFailure

        exit function

    end if

    if Err.Number <> kErrorSuccess then

        wscript.echo L_Text_Msg_General11_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        DelAllPrinters = kErrorFailure

        exit function

    end if

    iTotal = 0
    iTotalDeleted = 0

    for each oPrinter in Printers

        strPrinterName = oPrinter.DeviceID

        bConnection = oPrinter.Network

        if kAction = kActionDelAll then

            bDelete = 1

            iTotal = iTotal + 1

        elseif kAction = kActionDelAllCon and bConnection then

            bDelete = 1

            iTotal = iTotal + 1

        elseif kAction = kActionDelAllLocal and not bConnection then

            bDelete = 1

            iTotal = iTotal + 1

        else

            bDelete = 0

        end if

        if bDelete = 1 then

            if bConnection then

                strTemp = L_Space_Text & L_Connection_Text & L_Space_Text

            else

                strTemp = L_Space_Text

            end if

            '
            ' Delete printer instance
            '
            oPrinter.Delete_

            if Err.Number = kErrorSuccess then

                wscript.echo L_Text_Msg_General05_Text & strTemp & oPrinter.DeviceID

                iTotalDeleted = iTotalDeleted + 1

            else

                wscript.echo L_Text_Msg_General06_Text & strTemp & strPrinterName _
                             & L_Space_Text & L_Error_Text & L_Space_Text & L_Hex_Text _
                             & hex(Err.Number) & L_Space_Text & Err.Description

                '
                ' Try getting extended error information
                '
                call LastError()

                '
                ' Continue deleting the rest of the printers despite this error
                '
                Err.Clear

            end if

        end if

    next

    wscript.echo L_Empty_Text

    if kAction = kActionDelAll then

        wscript.echo L_Text_Msg_General09_Text & L_Space_Text & iTotal
        wscript.echo L_Text_Msg_General10_Text & L_Space_Text & iTotalDeleted

    elseif kAction = kActionDelAllCon then

        wscript.echo L_Text_Msg_General16_Text & L_Space_Text & iTotal
        wscript.echo L_Text_Msg_General17_Text & L_Space_Text & iTotalDeleted

    elseif kAction = kActionDelAllLocal then

        wscript.echo L_Text_Msg_General18_Text & L_Space_Text & iTotal
        wscript.echo L_Text_Msg_General19_Text & L_Space_Text & iTotalDeleted

    else

    end if

    DelAllPrinters = kErrorSuccess

end function

'
' List the printers
'
function ListPrinters(strServer, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg05_Text

    dim Printers
    dim oService
    dim oPrinter
    dim iTotal

    if WmiConnect(strServer, kNameSpace, strUser, strPassword, oService) then

        set Printers = oService.InstancesOf("Win32_Printer")

    else

        ListPrinters = kErrorFailure

        exit function

    end if

    if Err.Number <> kErrorSuccess then

        wscript.echo L_Text_Msg_General11_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        ListPrinters = kErrorFailure

        exit function

    end if

    iTotal = 0

    for each oPrinter in Printers

        iTotal = iTotal + 1

        wscript.echo L_Empty_Text
        wscript.echo L_Text_Msg_Printer01_Text & L_Space_Text & strServer
        wscript.echo L_Text_Msg_Printer02_Text & L_Space_Text & oPrinter.DeviceID
        wscript.echo L_Text_Msg_Printer03_Text & L_Space_Text & oPrinter.ShareName
        wscript.echo L_Text_Msg_Printer04_Text & L_Space_Text & oPrinter.DriverName
        wscript.echo L_Text_Msg_Printer05_Text & L_Space_Text & oPrinter.PortName
        wscript.echo L_Text_Msg_Printer06_Text & L_Space_Text & oPrinter.Comment
        wscript.echo L_Text_Msg_Printer07_Text & L_Space_Text & oPrinter.Location
        wscript.echo L_Text_Msg_Printer08_Text & L_Space_Text & oPrinter.SepFile
        wscript.echo L_Text_Msg_Printer09_Text & L_Space_Text & oPrinter.PrintProcessor
        wscript.echo L_Text_Msg_Printer10_Text & L_Space_Text & oPrinter.PrintJobDataType
        wscript.echo L_Text_Msg_Printer11_Text & L_Space_Text & oPrinter.Parameters
        wscript.echo L_Text_Msg_Printer12_Text & L_Space_Text & CSTR(oPrinter.Attributes)
        wscript.echo L_Text_Msg_Printer13_Text & L_Space_Text & CSTR(oPrinter.Priority)
        wscript.echo L_Text_Msg_Printer14_Text & L_Space_Text & CStr(oPrinter.DefaultPriority)

        if CStr(oPrinter.StartTime) <> "" and CStr(oPrinter.UntilTime) <> "" then

            wscript.echo L_Text_Msg_Printer15_Text & L_Space_Text & Mid(Mid(CStr(oPrinter.StartTime), 9, 4), 1, 2) & "h" & Mid(Mid(CStr(oPrinter.StartTime), 9, 4), 3, 2)
            wscript.echo L_Text_Msg_Printer16_Text & L_Space_Text & Mid(Mid(CStr(oPrinter.UntilTime), 9, 4), 1, 2) & "h" & Mid(Mid(CStr(oPrinter.UntilTime), 9, 4), 3, 2)

        end if

        wscript.echo L_Text_Msg_Printer17_Text & L_Space_Text & CStr(oPrinter.Jobs)
        wscript.echo L_Text_Msg_Printer18_Text & L_Space_Text & CStr(oPrinter.AveragePagesPerMinute)
        wscript.echo L_Text_Msg_Printer19_Text & L_Space_Text & PrnStatusToString(oPrinter.PrinterStatus)
        wscript.echo L_Text_Msg_Printer20_Text & L_Space_Text & ExtPrnStatusToString(oPrinter.ExtendedPrinterStatus)
        wscript.echo L_Text_Msg_Printer21_Text & L_Space_Text & DetectedErrorStateToString(oPrinter.DetectedErrorState)
        wscript.echo L_Text_Msg_Printer22_Text & L_Space_Text & ExtDetectedErrorStateToString(oPrinter.ExtendedDetectedErrorState)

        Err.Clear

    next

    wscript.echo L_Empty_Text
    wscript.echo L_Text_Msg_General09_Text & L_Space_Text & iTotal

    ListPrinters = kErrorSuccess

end function

'
' Get the default printer
'
function GetDefaultPrinter(strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg06_Text

    dim oService
    dim oPrinter
    dim iRetval
    dim oEnum

    iRetval = kErrorFailure

    '
    ' We connect to the local server
    '
    if WmiConnect("", kNameSpace, strUser, strPassword, oService) then

        set oEnum    = oService.ExecQuery("select DeviceID from Win32_Printer where default=true")

    else

        SetDefaultPrinter = kErrorFailure

        exit function

    end if

    if Err.Number = kErrorSuccess then

         for each oPrinter in oEnum

            wscript.echo L_Text_Msg_General12_Text & L_Space_Text & oPrinter.DeviceID

         next

         iRetval = kErrorSuccess

    else

        wscript.echo L_Text_Msg_General13_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

    end if

    GetDefaultPrinter = iRetval

end function

'
' Set the default printer
'
function SetDefaultPrinter(strPrinter, strUser, strPassword)

    'on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg07_Text

    dim oService
    dim oPrinter
    dim iRetval
    dim uResult

    iRetval = kErrorFailure

    '
    ' We connect to the local server
    '
    if WmiConnect("", kNameSpace, strUser, strPassword, oService) then

        set oPrinter = oService.Get("Win32_Printer.DeviceID='" & strPrinter & "'")

    else

        SetDefaultPrinter = kErrorFailure

        exit function

    end if

    '
    ' Check if Get was successful
    '
    if Err.Number = kErrorSuccess then

        '
        ' The Err object indicates whether the WMI provider reached the execution
        ' of the function that sets the default printer. The uResult is the Win32
        ' error code of the spooler function that sets the default printer
        '
        uResult = oPrinter.SetDefaultPrinter

        if Err.Number = kErrorSuccess then

            if uResult = kErrorSuccess then

                wscript.echo L_Text_Msg_General15_Text & L_Space_Text & strPrinter

                iRetval = kErrorSuccess

            else

                wscript.echo L_Text_Msg_General14_Text & L_Space_Text _
                             & L_Text_Error_General03_Text& L_Space_Text & uResult

            end if

        else

            wscript.echo L_Text_Msg_General14_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                         & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        end if

    else

        wscript.echo L_Text_Msg_General14_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

        '
        ' Try getting extended error information
        '
        call LastError()

    end if

    SetDefaultPrinter = iRetval

end function

'
' Converts the printer status to a string
'
function PrnStatusToString(Status)

    dim str

    str = L_Empty_Text

    select case Status

        case 1
            str = str + L_Text_Msg_Status01_Text + L_Space_Text

        case 2
            str = str + L_Text_Msg_Status02_Text + L_Space_Text

        case 3
            str = str + L_Text_Msg_Status03_Text + L_Space_Text

        case 4
            str = str + L_Text_Msg_Status04_Text + L_Space_Text

        case 5
            str = str + L_Text_Msg_Status05_Text + L_Space_Text

        case 6
            str = str + L_Text_Msg_Status06_Text + L_Space_Text

        case 7
            str = str + L_Text_Msg_Status07_Text + L_Space_Text

    end select

    PrnStatusToString = str

end function

'
' Converts the extended printer status to a string
'
function ExtPrnStatusToString(Status)

    dim str

    str = L_Empty_Text

    select case Status

        case 1
            str = str + L_Text_Msg_Status01_Text + L_Space_Text

        case 2
            str = str + L_Text_Msg_Status02_Text + L_Space_Text

        case 3
            str = str + L_Text_Msg_Status03_Text + L_Space_Text

        case 4
            str = str + L_Text_Msg_Status04_Text + L_Space_Text

        case 5
            str = str + L_Text_Msg_Status05_Text + L_Space_Text

        case 6
            str = str + L_Text_Msg_Status06_Text + L_Space_Text

        case 7
            str = str + L_Text_Msg_Status07_Text + L_Space_Text

        case 8
            str = str + L_Text_Msg_Status08_Text + L_Space_Text

        case 9
            str = str + L_Text_Msg_Status09_Text + L_Space_Text

        case 10
            str = str + L_Text_Msg_Status10_Text + L_Space_Text

        case 11
            str = str + L_Text_Msg_Status11_Text + L_Space_Text

        case 12
            str = str + L_Text_Msg_Status12_Text + L_Space_Text

        case 13
            str = str + L_Text_Msg_Status13_Text + L_Space_Text

        case 14
            str = str + L_Text_Msg_Status14_Text + L_Space_Text

        case 15
            str = str + L_Text_Msg_Status15_Text + L_Space_Text

        case 16
            str = str + L_Text_Msg_Status16_Text + L_Space_Text

        case 17
            str = str + L_Text_Msg_Status17_Text + L_Space_Text

        case 18
            str = str + L_Text_Msg_Status18_Text + L_Space_Text

    end select

    ExtPrnStatusToString = str

end function

'
' Converts the detected error state to a string
'
function DetectedErrorStateToString(Status)

    dim str

    str = L_Empty_Text

    select case Status

        case 0
            str = str + L_Text_Msg_Status02_Text + L_Space_Text

        case 1
            str = str + L_Text_Msg_Status01_Text + L_Space_Text

        case 2
            str = str + L_Text_Msg_Status01_Text + L_Space_Text

        case 3
            str = str + L_Text_Msg_Status20_Text + L_Space_Text

        case 4
            str = str + L_Text_Msg_Status21_Text + L_Space_Text

        case 5
            str = str + L_Text_Msg_Status22_Text + L_Space_Text

        case 6
            str = str + L_Text_Msg_Status23_Text + L_Space_Text

        case 7
            str = str + L_Text_Msg_Status24_Text + L_Space_Text

        case 8
            str = str + L_Text_Msg_Status25_Text + L_Space_Text

        case 9
            str = str + L_Text_Msg_Status07_Text + L_Space_Text

        case 10
            str = str + L_Text_Msg_Status26_Text + L_Space_Text

        case 11
            str = str + L_Text_Msg_Status27_Text + L_Space_Text

    end select

    DetectedErrorStateToString = str

end function

'
' Converts the extended detected error state to a string
'
function ExtDetectedErrorStateToString(Status)

    dim str

    str = L_Empty_Text

    select case Status

        case 0
            str = str + L_Text_Msg_Status02_Text + L_Space_Text

        case 1
            str = str + L_Text_Msg_Status01_Text + L_Space_Text

        case 2
            str = str + L_Text_Msg_Status01_Text + L_Space_Text

        case 3
            str = str + L_Text_Msg_Status20_Text + L_Space_Text

        case 4
            str = str + L_Text_Msg_Status21_Text + L_Space_Text

        case 5
            str = str + L_Text_Msg_Status22_Text + L_Space_Text

        case 6
            str = str + L_Text_Msg_Status23_Text + L_Space_Text

        case 7
            str = str + L_Text_Msg_Status24_Text + L_Space_Text

        case 8
            str = str + L_Text_Msg_Status25_Text + L_Space_Text

        case 9
            str = str + L_Text_Msg_Status07_Text + L_Space_Text

        case 10
            str = str + L_Text_Msg_Status26_Text + L_Space_Text

        case 11
            str = str + L_Text_Msg_Status27_Text + L_Space_Text

        case 12
            str = str + L_Text_Msg_Status28_Text + L_Space_Text

        case 13
            str = str + L_Text_Msg_Status29_Text + L_Space_Text

        case 14
            str = str + L_Text_Msg_Status30_Text + L_Space_Text

        case 15
            str = str + L_Text_Msg_Status31_Text + L_Space_Text

        case 16
            str = str + L_Text_Msg_Status32_Text + L_Space_Text

    end select

    ExtDetectedErrorStateToString = str

end function

'
' Debug display helper function
'
sub DebugPrint(uFlags, strString)

    if gDebugFlag = true then

        if uFlags = kDebugTrace then

            wscript.echo L_Debug_Text & L_Space_Text & strString

        end if

        if uFlags = kDebugError then

            if Err <> 0 then

                wscript.echo L_Debug_Text & L_Space_Text & strString & L_Space_Text _
                             & L_Error_Text & L_Space_Text & L_Hex_Text & hex(Err.Number) _
                             & L_Space_Text & Err.Description

            end if

        end if

    end if

end sub

'
' Parse the command line into its components
'
function ParseCommandLine(iAction, strServer, strPrinter, strDriver, strPort, strUser, strPassword)

    on error resume next

    DebugPrint kDebugTrace, L_Text_Dbg_Msg08_Text

    dim oArgs
    dim iIndex

    iAction = kActionUnknown
    iIndex  = 0

    set oArgs = wscript.Arguments

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-a"
                iAction = kActionAdd

            case "-ac"
                iAction = kActionAddConn

            case "-d"
                iAction = kActionDel

            case "-x"
                iAction = kActionDelAll

            case "-xc"
                iAction = kActionDelAllCon

            case "-xo"
                iAction = kActionDelAllLocal

            case "-l"
                iAction = kActionList

            case "-g"
                iAction = kActionGetDefaultPrinter

            case "-t"
                iAction = kActionSetDefaultPrinter

            case "-s"
                iIndex = iIndex + 1
                strServer = RemoveBackslashes(oArgs(iIndex))

            case "-p"
                iIndex = iIndex + 1
                strPrinter = oArgs(iIndex)

            case "-m"
                iIndex = iIndex + 1
                strDriver = oArgs(iIndex)

            case "-u"
                iIndex = iIndex + 1
                strUser = oArgs(iIndex)

            case "-w"
                iIndex = iIndex + 1
                strPassword = oArgs(iIndex)

            case "-r"
                iIndex = iIndex + 1
                strPort = oArgs(iIndex)

            case "-?"
                Usage(true)
                exit function

            case else
                Usage(true)
                exit function

        end select

        iIndex = iIndex + 1

    wend

    if Err = kErrorSuccess then

        ParseCommandLine = kErrorSuccess

    else

        wscript.echo L_Text_Error_General02_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_text & Err.Description

        ParseCommandLine = kErrorFailure

    end if

end  function

'
' Display command usage.
'
sub Usage(bExit)

    wscript.echo L_Help_Help_General01_Text
    wscript.echo L_Help_Help_General02_Text
    wscript.echo L_Help_Help_General03_Text
    wscript.echo L_Help_Help_General04_Text
    wscript.echo L_Help_Help_General05_Text
    wscript.echo L_Help_Help_General06_Text
    wscript.echo L_Help_Help_General07_Text
    wscript.echo L_Help_Help_General08_Text
    wscript.echo L_Help_Help_General09_Text
    wscript.echo L_Help_Help_General10_Text
    wscript.echo L_Help_Help_General11_Text
    wscript.echo L_Help_Help_General12_Text
    wscript.echo L_Help_Help_General13_Text
    wscript.echo L_Help_Help_General14_Text
    wscript.echo L_Help_Help_General15_Text
    wscript.echo L_Help_Help_General16_Text
    wscript.echo L_Help_Help_General17_Text
    wscript.echo L_Help_Help_General18_Text
    wscript.echo L_Help_Help_General19_Text
    wscript.echo L_Empty_Text
    wscript.echo L_Help_Help_General20_Text
    wscript.echo L_Help_Help_General21_Text
    wscript.echo L_Help_Help_General22_Text
    wscript.echo L_Help_Help_General23_Text
    wscript.echo L_Help_Help_General24_Text
    wscript.echo L_Help_Help_General25_Text
    wscript.echo L_Help_Help_General26_Text
    wscript.echo L_Help_Help_General27_Text
    wscript.echo L_Help_Help_General28_Text
    wscript.echo L_Help_Help_General29_Text

    if bExit then

        wscript.quit(1)

    end if

end sub

'
' Determines which program is being used to run this script.
' Returns true if the script host is cscript.exe
'
function IsHostCscript()

    on error resume next

    dim strFullName
    dim strCommand
    dim i, j
    dim bReturn

    bReturn = false

    strFullName = WScript.FullName

    i = InStr(1, strFullName, ".exe", 1)

    if i <> 0 then

        j = InStrRev(strFullName, "\", i, 1)

        if j <> 0 then

            strCommand = Mid(strFullName, j+1, i-j-1)

            if LCase(strCommand) = "cscript" then

                bReturn = true

            end if

        end if

    end if

    if Err <> 0 then

        wscript.echo L_Text_Error_General01_Text & L_Space_Text & L_Error_Text & L_Space_Text _
                     & L_Hex_Text & hex(Err.Number) & L_Space_Text & Err.Description

    end if

    IsHostCscript = bReturn

end function

'
' Retrieves extended information about the last error that occurred
' during a WBEM operation. The methods that set an SWbemLastError
' object are GetObject, PutInstance, DeleteInstance
'
sub LastError()

    on error resume next

    dim oError

    set oError = CreateObject("WbemScripting.SWbemLastError")

    if Err = kErrorSuccess then

        wscript.echo L_Operation_Text            & L_Space_Text & oError.Operation
        wscript.echo L_Provider_Text             & L_Space_Text & oError.ProviderName
        wscript.echo L_Description_Text          & L_Space_Text & oError.Description
        wscript.echo L_Text_Error_General03_Text & L_Space_Text & oError.StatusCode

    end if

end sub

'
' Connects to the WMI service on a server. oService is returned as a service
' object (SWbemServices)
'
function WmiConnect(strServer, strNameSpace, strUser, strPassword, oService)

    on error resume next

    dim oLocator
    dim bResult

    oService = null

    bResult  = false

    set oLocator = CreateObject("WbemScripting.SWbemLocator")

    if Err = kErrorSuccess then

        set oService = oLocator.ConnectServer(strServer, strNameSpace, strUser, strPassword)

        if Err = kErrorSuccess then

            bResult = true

            oService.Security_.impersonationlevel = 3

            '
            ' Required to perform administrative tasks on the spooler service
            '
            oService.Security_.Privileges.AddAsString "SeLoadDriverPrivilege"

            Err.Clear

        else

            wscript.echo L_Text_Msg_General11_Text & L_Space_Text & L_Error_Text _
                         & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                         & Err.Description

        end if

    else

        wscript.echo L_Text_Msg_General10_Text & L_Space_Text & L_Error_Text _
                     & L_Space_Text & L_Hex_Text & hex(Err.Number) & L_Space_Text _
                     & Err.Description

    end if

    WmiConnect = bResult

end function

'
' Remove leading "\\" from server name
'
function RemoveBackslashes(strServer)

    dim strRet

    strRet = strServer

    if Left(strServer, 2) = "\\" and Len(strServer) > 2 then

        strRet = Mid(strServer, 3)

    end if

    RemoveBackslashes = strRet

end function
