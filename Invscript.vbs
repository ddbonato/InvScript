On Error Resume Next

'----------------------------Mensagem inicio
CreateObject("WScript.Shell").Popup "Capturando Informações...Aguarde"& vbCrLf & vbCrLf, 3, "InvScript"         
'----------------------------- Pegar nome PC
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery( _
"SELECT * FROM Win32_ComputerSystem",,48)
For Each objItem in colItems
nomepc = objItem.Caption
Next

'-----------------------------------------deletar arquivo caso exista
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile ".\"+nomepc+".txt"
obj.DeleteFile ".\Model\chavetemp.txt"
Set obj=Nothing



'----------------------------- Criar o arquivo

Dim fso, txtfile
Set fso = CreateObject("Scripting.FileSystemObject")
Set txtfile = fso.CreateTextFile(".\"+nomepc+".txt", True)



'----------------------Pegar SN  ------------------------------------------

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * From Win32_BIOS")

For Each objItem in colItems

txtfile.write ("|NUMERO DE SERIE|")
txtfile.WriteBlankLines(1)
nserie = objItem.SerialNumber
If InStr(nserie,"Default") > 0 Then
txtfile.Write ("O número de Série não está armazenado na placa mãe.")
txtfile.WriteBlankLines(1)
ElseIf InStr(nserie,"00000000") > 0 Then
txtfile.Write ("O número de Série não está armazenado na placa mãe.")
txtfile.WriteBlankLines(1)
Else
txtfile.Write (nserie)
txtfile.WriteBlankLines(1)
End If
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(1)
Next





'----------------------------- Nome PC
txtfile.write ("|HOSTNAME|")
txtfile.WriteBlankLines(1)
txtfile.Write (nomepc)
txtfile.WriteBlankLines(1)
txtfile.Write ("==================================================")
'--------------- Software
'strComputer = "."
'strProperties = "*"'"CSName, Caption, OSType, Version, OSProductSuite, BuildNumber, 'ProductType, OSLanguage, CSDVersion, InstallDate, RegisteredUser, Organization, 'SerialNumber, WindowsDirectory, SystemDirectory"
'objClass = "Win32_OperatingSystem"
'strQuery = "SELECT " & strProperties & " FROM " & objClass
'Set colOS = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + 'wbemFlagForwardOnly)
'For Each objItem in colOS
'txtfile.write ("Sistema Operacional: " +objItem.Caption)
'txtfile.WriteBlankLines(1)
'If SO_Type = 16 Then
'SO_Name = "Microsoft Windows 95"
'ElseIf SO_Type = 17 Then
'SO_Name = "Microsoft Windows 98"
'End If
'If SO_ProdType = 1 Then
'SO_ProdType = "Esta??o de Trabalho"
'ElseIf SO_ProdType = 2 Then
'SO_ProdType = "Controlador de Dom?nio"
'ElseIf SO_ProdType = 3 Then
'SO_ProdType = "Servidor"
'End If
'If SO_Language = 1033 Then
'SO_Language = "Ingl?s - Estados Unidos"
'ElseIf SO_Language = 1046 Then
'SO_Language = "Portugu?s - Brasil"
'Else
'SO_Language = "Outro idioma"
'End If
'If SO_Suite = 1 Then
'SO_Suite = "Small Business"
'ElseIf SO_Suite = 2 Then
'SO_Suite = "Enterprise"
'ElseIf SO_Suite = 4 Then
'SO_Suite = "Backoffice"
'ElseIf SO_Suite = 8 Then
'SO_Suite = "Communication Server"
'ElseIf SO_Suite = 16 Then
'SO_Suite = "Terminal Server"
'ElseIf SO_Suite = 18 Then
'SO_Suite = "Enterprise e Terminal Server"
'ElseIf SO_Suite = 32 Then
'SO_Suite = "Small Business (Restrito)"
'ElseIf SO_Suite = 64 Then
'SO_Suite = "Embedded NT"
'ElseIf SO_Suite = 128 Then
'SO_Suite = "Data Center"
'ElseIf SO_Suite = 256 Then
'SO_Suite = "Single User"
'ElseIf SO_Suite = 512 Then
'SO_Suite = "Personal"
'ElseIf SO_Suite = 1024 Then
'SO_Suite = "Blade"
'End If
'Next

'------------------------------------------------------Funcao Chave windows
'Set WshShell = CreateObject("WScript.Shell")
'Function ConvertToKey(Key)

'Const KeyOffset = 52
'i = 28
'Chars = "BCDFGHJKMPQRTVWXY2346789"
'Do
'Cur = 0
'x = 14
'Do
'Cur = Cur * 256
'Cur = Key(x + KeyOffset) + Cur
'Key(x + KeyOffset) = (Cur \ 24) And 255
'Cur = Cur Mod 24
'x = x -1
'Loop While x >= 0
'i = i -1
'KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
'If (((29 - i) Mod 6) = 0) And (i <> -1) Then
'i = i -1
'KeyOutput = "-" & KeyOutput
'End If
'Loop While i >= 0
'ConvertToKey = KeyOutput
'End Function

'txtfile.write("Chave: " +ConvertToKey(WshShell.RegRead("HKLM\SOFTWARE\Microsoft'\Windows NT\CurrentVersion\DigitalProductId")))
'txtfile.WriteBlankLines(1)
'--------------------------------------------------Descobrir Office
'strComputer = "."
'Set objWMIService = GetObject("winmgmts:" _
'     & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
' Set colSoft = objWMIService.ExecQuery("SELECT * FROM Win32_Product WHERE Name Like 'Microsoft Office%'")

 'If colSoft.Count = 0 Then
 '   txtfile.write ("Não possui office: Libre Office" )
 'else
 '   For Each objItem In colSoft
 '       txtfile.write ("Office: " & objitem.caption) 
 '      exit for
 '   Next
 'End If
txtfile.WriteBlankLines(1)
txtfile.Write("|OFFICE & WINDOWS|")
txtfile.WriteBlankLines(1)



'----------------------------------------------------MAC
'Descobrir sistema
strComputer = "."
strProperties = "*"'"CSName, Caption, OSType, Version, OSProductSuite, BuildNumber, ProductType, OSLanguage, CSDVersion, InstallDate, RegisteredUser, Organization, SerialNumber, WindowsDirectory, SystemDirectory"
objClass = "Win32_OperatingSystem"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colOS = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colOS
sistema = objItem.Caption
next 


'If Windows XP
if sistema = "Microsoft Windows XP Professional" then 
strQuery = "SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionID > ''"
Set objWMIService = GetObject( "winmgmts://./root/CIMV2" )
Set colItems      = objWMIService.ExecQuery( strQuery, "WQL", 48 )
txtfile.write ("|MAC|")
contatodormac = 0
For Each objItem In colItems
contadormac = contadormac + 1
if not isnull(objItem.MACAddress) then txtfile.write (vbCrLf & "MAC " & contadormac & ": " & objItem.MACAddress)
Next
txtfile.WriteBlankLines(1)
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(1)
Else 
    txtfile.write ("|MAC|")
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter where physicaladapter=true")
    for Each objItem in colItems
        if not isnull(objItem.MACAddress) then txtfile.write (vbCrLf & objItem.description & ": " & objItem.MACAddress)
        next 
txtfile.WriteBlankLines(1)
    txtfile.Write ("==================================================")
txtfile.WriteBlankLines(1)
    End If


'--------------------------------------------------------------------Placa mae
strComputer = "." 
Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
 
Set colItems = objWMIService.ExecQuery("Select * from Win32_BaseBoard") 
txtfile.write("|PLACA MÃE|")
txtfile.WriteBlankLines(1)
For Each objItem in colItems 
    placamae = objItem.Manufacturer
    modelo = objItem.Product
    txtfile.write(placamae &"-"& modelo)
Next
txtfile.WriteBlankLines(1)
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(1)

'-------------------------------Processador
txtfile.write ("|PROCESSADOR|")
txtfile.WriteBlankLines(1)
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery( _
"SELECT * FROM Win32_Processor",,48)
For Each objItem in colItems


'------------------------------------------------- Nome do processador
txtfile.write(objItem.name)
txtfile.WriteBlankLines(2)
Next
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(1)
'----------------------------------Memoria
txtfile.write ("|MEMORIA|")
txtfile.WriteBlankLines(1)
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery( _
"SELECT * FROM Win32_physicalmemory",,48)
'Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
cont = 0
memoriatotal = 0
For Each objItem in colItems

cont = (cont + 1)
txtfile.write ("Modulo " & cont & ": " & objItem.capacity/1048576 & " MB")
memoriatotal = (objItem.capacity/1048576 + memoriatotal) 
txtfile.WriteBlankLines(1)
Next
txtfile.write("Memoria total: " & (memoriatotal/1024) &" GB")
txtfile.WriteBlankLines(1)
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(1)
'---------------------------------- hd
txtfile.write ("|HD/SSD| ")
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery( _
"SELECT * FROM Win32_diskdrive",,48)
contadorhd = 0
For Each objItem in colItems
'------------------------------------------------- modelo do disco
'txtfile.write ("Disco:")
'txtfile.WriteBlankLines(1)
'txtfile.write (objItem.caption)
'txtfile.WriteBlankLines(1)
'----------------------------------------------------- Interface
'txtfile.write ("Interface:")
'txtfile.WriteBlankLines(1)
'txtfile.write (objItem.interfacetype)
txtfile.WriteBlankLines(1)
contadorhd = (contadorhd + 1)
txtfile.write ("Disco "& contadorhd)

'----------------------------------------------------- Capacidade
capacidade = int(objItem.size/1073741824)
If capacidade > 900 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 1 TB")
ElseIf capacidade > 695 And capacidade < 750 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 750 GB")
ElseIf capacidade > 400 And capacidade < 500 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 500 GB")
ElseIf capacidade > 231 And capacidade < 250 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 250 GB")
ElseIf capacidade > 225 And capacidade < 240 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 240 GB")
ElseIf capacidade > 140 And capacidade < 160 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 160 GB")
ElseIf capacidade > 110 And capacidade < 120 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 120 GB")
ElseIf capacidade > 70 And capacidade < 80 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 80 GB")
End If
txtfile.WriteBlankLines(1)
txtfile.write ("Tamanho Real: ")
txtfile.write (Int(objItem.size/1073741824) & " GB")
txtfile.WriteBlankLines(1)
txtfile.Write ("--------------------------------------------------")
Next
txtfile.WriteBlankLines(1)
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(1)

'--------------------------------------------------------------Pegar informação Placa de vídeo
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoController")

For Each objItem in colItems

txtfile.write ("Adptador de vídeo: " & objItem.Description)
Next
txtfile.WriteBlankLines(1)
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(1)

'------------------------------------------------- Nome do adaptador

txtfile.write ("|IP|")
txtfile.WriteBlankLines(1)
strComputer = "."
strProperties = "Description, MACAddress, IPAddress, IPSubnet, DefaultIPGateway, DNSServerSearchOrder, DNSDomain, DNSDomainSuffixSearchOrder, DHCPEnabled, DHCPServer, WINSPrimaryServer, WINSSecondaryServer, ServiceName"
objClass = "Win32_NetworkAdapterConfiguration"
strQuery = "SELECT " & strProperties & " FROM " & objClass & " WHERE IPEnabled = True AND ServiceName <> 'AsyncMac' AND ServiceName <> 'VMnetx' AND ServiceName <> 'VMnetadapter' AND ServiceName <> 'Rasl2tp' AND ServiceName <> 'PptpMiniport' AND ServiceName <> 'Raspti' AND ServiceName <> 'NDISWan' AND ServiceName <> 'RasPppoe' AND ServiceName <> 'NdisIP' AND ServiceName <> ''"
Set colAdapters = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
'--------------------------------------------------------rede
For Each objItem in colAdapters
'For Each objItem in colItems
'txtfile.write ("Adaptador:")
'txtfile.WriteBlankLines(1)
'txtfile.write (objItem.Description)
'txtfile.WriteBlankLines(1)
'------------------------------------------------- IP
'txtfile.write ("IP: ")
'txtfile.WriteBlankLines(1)
IP_Address = objItem.IPAddress
txtfile.write (IP_Address(i))
txtfile.WriteBlankLines(1)
Next
txtfile.WriteBlankLines(1)
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(1)
txtfile.write("*****Instituto das Cidades Inteligentes******")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@          @@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@               @@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@           @@@@@@@      @@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@         @@@@@@@@@@@@@@     @@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@         @@@@@@@@@@       @@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@             @@   @@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@          @@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@ @@@@@     @@@@  @@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@ @@@@  @@@@@@@@  @@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@ @@@@@    @@@@@  @@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
Wscript.Echo "Informações adicionadas com êxito!" & vbCrLf & vbCrLf & vbCrLf & "Script desenvolvido por Daniel Bonato | 2020"
wscript.quit