/*
 *
 *
 */

/*
 *	includes - Libs
 */
#include <Gdip>


/*
 *	includes - others
 */
#include alvara_defines.ahk


/*
 *	globals
 */
TITLE := "Alvarás Automatizados"
shortSleep := 200
row := 0


/*
 *	autoexecute
 */
MsgBox, 0, %TITLE%, Sistema de Verificação dos Alvarás Automatizados
Sleep, %shortSleep%

pToken := Gdip_Startup()

dir := A_Desktop . "\" . "AlvarasAutomatizados_" . A_Now
FileCreateDir, %dir%
Sleep, %shortSleep%

MsgBox, 0, %TITLE%, Em seguida`, selecione o arquivo com a planilha dos Alvarás Automatizados.
Sleep, %shortSleep%

FileSelectFile, pathxl, , , Selecione Arquivo com Alvarás Automatizados, *.xlsx
Sleep, %shortSleep%

If !IsObject(xl)
	xl := ComObjCreate("Excel.Application")
xl.Workbooks.Open(pathxl)
xl.Visible := True
Sleep, %shortSleep%

MsgBox, 0, %TITLE%, 
(LTrim
Para garantir o correto funcionamento do script,
Verifique e Confirme as seguintes informações.
)
processo := xl.Range("G2").Text

MsgBox, 4, %TITLE%, O número do Processo Administrativo é %processo%
IfMsgBox, No
    Return

lastrow := xl.Range("A" xl.Rows.Count).End(xlUp := -4162).Row
firstrow := xl.Range("A" lastrow).End(xlUp := -4162).Row + 1

MsgBox, 4, %TITLE%, 
(LTrim
Primeira Linha com Alvará = %firstrow%
Última Linha com Alvará = %lastrow%
)
IfMsgBox, No
    Return

row := firstrow
num_alv := lastrow - firstrow

arr := xl.Range(_COL_ARR . row).Text
val := xl.Range(_COL_VAL . row).Text
cod := xl.Range(_COL_COD . row).Text
alv := SubStr(xl.Range(_COL_ALV . row).Text, -10)
pro := SubStr("000" . xl.Range(_COL_PROC . row).Text, -13)

MsgBox, 4, %TITLE%, 
(LTrim
Dados do Primeiro Alvará:

Arrecadação Nro: %arr%
Valor: %val%
Código: %cod%

Comentário:
Apropriacao do Alvara n. %alv%, expedido nos autos do processo  n. %pro%.

)
IfMsgBox, No
    Return

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
    Return

Send, {LShift Down}
Loop, 4 {
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
    If (!mod(alv, freq) || row == lastrow) {

        arr := xl.Range(_COL_ARR . row).Text
        val := xl.Range(_COL_VAL . row).Text
        cod := xl.Range(_COL_COD . row).Text
        alv := xl.Range(_COL_ALV . row).Text
        pro := xl.Range(_COL_PROC . row).Text
        
		Sleep, %shortSleep%
        
		If (cod == 304 || cod == 386 || cod == 640 || cod == 681 || cod == 760 || cod == 978
				|| cod == 1008 || cod == 1064 || cod == 1065 || cod == 1066 || cod == 1067
				|| cod == 1083 || cod == 1161 || cod == 1162
				|| cod == 478 || cod == 490
				|| cod == 761) {
            
			; Printing GA Original
			ControlSendRaw, , %arr%, ahk_pid %pwpid%
            Sleep, %shortSleep%
            
			ControlSend, , {Enter}, ahk_pid %pwpid%
            Sleep, 5000
			
			fname := dir . "\" . A_Now . "_" . arr . "_original.png"
            
			WinActivate, ahk_pid %pwpid%
            Sleep, 333
            Send, {Alt Down}{PrintScreen}{Alt Up}
            Sleep, 333
           
			Sleep, %shortSleepi%
			pBitmap := Gdip_CreateBitmapFromClipboard()
            sBitmap := Gdip_SaveBitmapToFile(pBitmap, fname, 100)
            dBitmap := Gdip_DisposeImage(pBitmap)
            Sleep, 500
            
			; Printing Observação PGE
			ControlSend, , {Enter}, ahk_pid %pwpid%
            Sleep, %shortSleep%
            
			ControlSend, , {Tab}, ahk_pid %pwpid%
            Sleep, %shortSleep%
            
			ControlSend, , {Tab}, ahk_pid %pwpid%
            Sleep, %shortSleep%
            
			ControlSend, , {n}, ahk_pid %pwpid%
            Sleep, %shortSleep%
            
			ControlSend, , {Enter}, ahk_pid %pwpid%
            Sleep, 2000
            
			fname := dir . "\" . A_Now . "_" . arr . "_observação.png"
            
			WinActivate, ahk_pid %pwpid%
            Sleep, 333
            Send, {Alt Down}{PrintScreen}{Alt Up}
            Sleep, 333
            
			Sleep, %shortSleep%
			pBitmap := Gdip_CreateBitmapFromClipboard()
            sBitmap := Gdip_SaveBitmapToFile(pBitmap, fname, 100)
            dBitmap := Gdip_DisposeImage(pBitmap)
            Sleep, 500
            
			; Printing Comentário Correção
			ControlSend, , {Enter}, ahk_pid %pwpid%
            Sleep, 2000
            
			fname := dir . "\" . A_Now . "_" . arr . "_comentário.png"
            
			WinActivate, ahk_pid %pwpid%
            Sleep, 333
            Send, {Alt Down}{PrintScreen}{Alt Up}
            Sleep, 333
            
			Sleep, %shortSleepi%
			pBitmap := Gdip_CreateBitmapFromClipboard()
            sBitmap := Gdip_SaveBitmapToFile(pBitmap, fname, 100)
            dBitmap := Gdip_DisposeImage(pBitmap)
            Sleep, 500
            
			; Printing GA Substituta
			ControlSend, , {Enter}, ahk_pid %pwpid%
            Sleep, 2000
            
			fname := dir . "\" . A_Now . "_" . arr . "_substituta.png"
            
			WinActivate, ahk_pid %pwpid%
            Sleep, 333
            Send, {Alt Down}{PrintScreen}{Alt Up}
            Sleep, 333
            
			Sleep, %shortSleep%
			pBitmap := Gdip_CreateBitmapFromClipboard()
            sBitmap := Gdip_SaveBitmapToFile(pBitmap, fname, 100)
            dBitmap := Gdip_DisposeImage(pBitmap)
            Sleep, 500

            ControlSend, , {n}, ahk_pid %pwpid%
            Sleep, %shortSleep%
            ControlSend, , {Enter}, ahk_pid %pwpid%
            Sleep, %shortSleep%
            
			xl.Range(row . ":" . row).Interior.ColorIndex := 4
            
			Sleep, 2000
        } Else {
            xl.Range("A" . row).Interior.ColorIndex := 3
            Sleep, %shortSleep%
        }
    }

    row += 1
    ;MsgBox, 0, , Próxima linha a executar: %row%
    Sleep, %shortSleep%
}

MsgBox, 0, %TITLE%, Encerrando sessão SOE.
Sleep, %shortSleep%

ControlSend, , {F12}, ahk_pid %pwpid%
Sleep, %shortSleep%

ControlSend, , {Enter}, ahk_pid %pwpid%
Sleep, 2000

WinClose, ahk_pid %pwpid%
Sleep, 333

sToken := Gdip_Shutdown(pToken)

MsgBox, 0, %TITLE%, Fim da Execução

Exitapp
