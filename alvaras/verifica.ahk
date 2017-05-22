#include Gdip.ahk

TITLE := "Alvarás Automatizados"
MsgBox, 0, %TITLE%, Sistema de Verificação dos Alvarás Automatizados
shortSleep := 100
row := 0
col_val := "C"
col_arr := "E"
col_alv := "F"
col_ret := "G"
col_cod := "H"
col_pro := "M"
col_com := "O"
dir := A_Desktop . "\" . "AlvarasAutomatizados_" . A_Now
FileCreateDir, %dir%
pToken := Gdip_Startup()
MsgBox, 0, %TITLE%, Em seguida`, selecione o arquivo com a planilha dos Alvarás Automatizados.
Sleep, %shortSleep%
FileSelectFile, pathxl, , , Selecione Arquivo com Alvarás Automatizados, *.xlsx
Sleep, %shortSleep%
If !IsObject(xl)
	xl := ComObjCreate("Excel.Application")
xl.Workbooks.Open(pathxl)
xl.Visible := True
Sleep, %shortSleep%
ControlSend, , {Enter}, ahk_pid %pwpid%
MsgBox, 0, Alvarás Automatizados, 
(LTrim
Para garantir o correto funcionamento do script,
Verifique e Confirme as seguintes informações.
)
processo := xl.Range("G2").Text
MsgBox, 4, %TITLE%, O número do Processo Administrativo é %processo%
IfMsgBox, No
{
    Return
}
firstrow := 4
lastrow := xl.Range("A" xl.Rows.Count).End(xlUp := -4162).Row

row := firstrow
num_alv := lastrow - firstrow
/*
MsgBox, 0, %TITLE%, 
(LTrim
Primeira Linha com Alvará = %row%
Última Linha com Alvará = %lastrow%
)
MsgBox, 0, %TITLE%, 
(LTrim
Num alvar. =  %num_alv%

)
*/
arr := xl.Range(col_arr . row).Text
val := xl.Range(col_val . row).Text
cod := xl.Range(col_cod . row).Text
com := xl.Range(col_com . row).Text
MsgBox, 4, %TITLE%, 
(LTrim
Dados do Primeiro Alvará:

Arrecadação Nro: %arr%
Valor: %val%
Código: %cod%
Comentário:
%com%

)
IfMsgBox, No
{
    Return
}
InputBox, matr, Alvarás Automatizados, Digite sua matrícula para login no SOE:
InputBox, pwr, Alvarás Automatizados, Digite sua senha para login no SOE:, hide
Sleep, %shortSleep%
Run, C:\Program Files (x86)\pw3270\pw3270.exe, , , pwpid
WinWait, ahk_pid %pwpid%
Sleep, 333
Sleep, 1000
ControlSendRaw, , ims, ahk_pid %pwpid%
Sleep, %shortSleep%
ControlSend, , {Enter}, ahk_pid %pwpid%
Sleep, 1000
ControlSend, , {Enter}, ahk_pid %pwpid%
Sleep, 2000
ControlSendRaw, , drpe, ahk_pid %pwpid%
Sleep, %shortSleep%
ControlSend, , {Tab}, ahk_pid %pwpid%
Sleep, %shortSleep%
ControlSendRaw, , %matr%, ahk_pid %pwpid%
Sleep, %shortSleep%
ControlSend, , {Tab}, ahk_pid %pwpid%
Sleep, %shortSleep%
ControlSendRaw, , %pwr%, ahk_pid %pwpid%
Sleep, %shortSleep%
ControlSend, , {Enter}, ahk_pid %pwpid%
Sleep, %shortSleep%
MsgBox, 52, %TITLE%, 
(LTrim
- Verifique se o terminal está aberto e logado no SOE;
- Aproveite AGORA para reposicionar as janelas do terminal e planilha se desejado;
- Enquanto o script estiver executando, NÃO se deve reposicionar nem interagir com o janela do terminal. Use no máximo os botões minimizar e restaurar.

Tudo Pronto?
)
IfMsgBox, No
{
    Return
}
Send, {LShift Down}
Loop, 4
{
    ControlSend, , {Tab 4}, ahk_pid %pwpid%
}
Send, {LShift Up}
Sleep, %shortSleep%
ControlSendRaw, , des, ahk_pid %pwpid%
Sleep, %shortSleep%
ControlSendRaw, , arr-con-nro, ahk_pid %pwpid%
Sleep, %shortSleep%
ControlSend, , {Tab}, ahk_pid %pwpid%
Sleep, %shortSleep%
ControlSend, , {Tab}, ahk_pid %pwpid%
Sleep, %shortSleep%
ControlSendRaw, , sar, ahk_pid %pwpid%
Sleep, %shortSleep%
ControlSend, , {Enter}, ahk_pid %pwpid%
Sleep, 2000
freq := ceil(num_alv / ceil(num_alv / 20))
While row <= lastrow
{
    alv := row - firstrow
    If (!mod(alv, freq) || row == lastrow)
    {
        arr := xl.Range(col_arr . row).Text
        val := xl.Range(col_val . row).Text
        cod := xl.Range(col_cod . row).Text
        alv := xl.Range(col_alv . row).Text
        pro := xl.Range(col_pro . row).Text
        com := xl.Range(col_com . row).Text
        Sleep, %shortSleep%
        If (cod == 304 || cod == 386 || cod == 681 || cod == 760 || cod == 1064 || cod == 1065 || cod == 1066 || cod == 1067 || cod == 761 || cod == 1083 || cod == 478)
        {
            ControlSendRaw, , %arr%, ahk_pid %pwpid%
            Sleep, %shortSleep%
            ControlSend, , {Enter}, ahk_pid %pwpid%
            Sleep, 5000
            fname := dir . "\" . A_Now . "_" . arr . "_original.png"
            WinActivate, ahk_pid %pwpid%
            Sleep, 333
            Send, {Alt Down}{PrintScreen}{Alt Up}
            Sleep, 333
            pBitmap := Gdip_CreateBitmapFromClipboard()
            sBitmap := Gdip_SaveBitmapToFile(pBitmap, fname, 100)
            dBitmap := Gdip_DisposeImage(pBitmap)
            Sleep, %shortSleep%
            ControlSend, , {Enter}, ahk_pid %pwpid%
            Sleep, %shortSleep%
            ControlSend, , {Tab}, ahk_pid %pwpid%
            Sleep, %shortSleep%
            ControlSend, , {Tab}, ahk_pid %pwpid%
            Sleep, %shortSleep%
            ControlSend, , {n}, ahk_pid %pwpid%
            Sleep, %shortSleep%
            ControlSend, , {Enter}, ahk_pid %pwpid%
            Sleep, 1000
            fname := dir . "\" . A_Now . "_" . arr . "_observação.png"
            WinActivate, ahk_pid %pwpid%
            Sleep, 333
            Send, {Alt Down}{PrintScreen}{Alt Up}
            Sleep, 333
            pBitmap := Gdip_CreateBitmapFromClipboard()
            sBitmap := Gdip_SaveBitmapToFile(pBitmap, fname, 100)
            dBitmap := Gdip_DisposeImage(pBitmap)
            Sleep, %shortSleep%
            ControlSend, , {Enter}, ahk_pid %pwpid%
            Sleep, 1000
            fname := dir . "\" . A_Now . "_" . arr . "_comentário.png"
            WinActivate, ahk_pid %pwpid%
            Sleep, 333
            Send, {Alt Down}{PrintScreen}{Alt Up}
            Sleep, 333
            pBitmap := Gdip_CreateBitmapFromClipboard()
            sBitmap := Gdip_SaveBitmapToFile(pBitmap, fname, 100)
            dBitmap := Gdip_DisposeImage(pBitmap)
            Sleep, %shortSleep%
            ControlSend, , {Enter}, ahk_pid %pwpid%
            Sleep, 1000
            fname := dir . "\" . A_Now . "_" . arr . "_substituta.png"
            WinActivate, ahk_pid %pwpid%
            Sleep, 333
            Send, {Alt Down}{PrintScreen}{Alt Up}
            Sleep, 333
            pBitmap := Gdip_CreateBitmapFromClipboard()
            sBitmap := Gdip_SaveBitmapToFile(pBitmap, fname, 100)
            dBitmap := Gdip_DisposeImage(pBitmap)
            Sleep, %shortSleep%
            ControlSend, , {n}, ahk_pid %pwpid%
            Sleep, %shortSleep%
            ControlSend, , {Enter}, ahk_pid %pwpid%
            Sleep, %shortSleep%
            xl.Range(row . ":" . row).Interior.ColorIndex := 4
            Sleep, 2000
        }
        Else
        {
            xl.Range("A" . row).Interior.ColorIndex := 3
            Sleep, %shortSleep%
        }
    }
    row += 1
    /*
    MsgBox, 0, , Próxima linha a executar: %row%
    */
    Sleep, %shortSleep%
}
MsgBox, 0, %TITLE%, Encerrando sessão SOE.
Sleep, %shortSleep%
ControlSend, , {F12}, ahk_pid %pwpid%
Sleep, %shortSleep%
ControlSend, , {Enter}, ahk_pid %pwpid%
Sleep, 2000
sToken := Gdip_Shutdown(pToken)
WinClose, ahk_pid %pwpid%
Sleep, 333
MsgBox, 0, %TITLE%, Fim da Execução
