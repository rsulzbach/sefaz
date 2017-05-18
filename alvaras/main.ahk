/********************************************
 *
 *
 ********************************************/

/*
 *	configure
 */
_CONFIG_CONFIRM_CHANGE = 0


/*
 *	globals
 */
VERS = 1.002
TITLE := "Alvarás Automatizados - " . VERS
shortSleep := 200
row := 0
col_val := "J"
col_arr := "E"
col_alv := "G"
col_ret := "K"
col_cod := "L"
col_pro := "C"
col_add := "M"

/*
 *	autoexecute
 */


if (_CONFIG_CONFIRM_CHANGE) {
	confirm_cmd := "S"
} else {
	confirm_cmd := "N"
}

MsgBox, 0, %TITLE%, Sistema de Apropriação dos Alvarás Automatizados
Sleep, %shortSleep%

MsgBox, 0, %TITLE%, Em seguida`, selecione o arquivo com a planilha dos Alvarás Automatizados.
Sleep, %shortSleep%

FileSelectFile, pathxl, , , Selecione Arquivo com Alvarás Automatizados, *.xlsx
Sleep, %shortSleep%

If !IsObject(xl)
	xl := ComObjCreate("Excel.Application")

xl.Workbooks.Open(pathxl)
xl.Visible := True

ControlSend, , {Enter}, ahk_pid %pwpid%
Sleep, %shortSleep%

MsgBox, 0, Alvarás Automatizados, 
(LTrim
Para garantir o correto funcionamento do script,
Verifique e Confirme as seguintes informações.
)
processo := xl.Range("G2").Text

MsgBox, 4, %TITLE%, O número do Processo Administrativo é %processo%
IfMsgBox, No
    Return

lastrow := xl.Range("A" xl.Rows.Count).End(xlUp := -4162).Row
row := 4

MsgBox, 0, %TITLE%, 
(LTrim
Primeira Linha com Alvará = %row%
Última Linha com Alvará = %lastrow%
)

arr := xl.Range(col_arr . row).Text
val := xl.Range(col_val . row).Text
cod := xl.Range(col_cod . row).Text
alv := SubStr(xl.Range(col_alv . row).Text, -10)
pro := SubStr("000" . xl.Range(col_pro . row).Text, -13)

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

ControlSendRaw, , arr-alt-gui, ahk_pid %pwpid%
Sleep, %shortSleep%

ControlSend, , {Tab}, ahk_pid %pwpid%
Sleep, %shortSleep%

ControlSend, , {Tab}, ahk_pid %pwpid%
Sleep, %shortSleep%

ControlSendRaw, , sar, ahk_pid %pwpid%
Sleep, %shortSleep%

ControlSend, , {Enter}, ahk_pid %pwpid%
Sleep, 2000

While row <= lastrow {
	arr := xl.Range(col_arr . row).Text
	val := xl.Range(col_val . row).Text
	cod := xl.Range(col_cod . row).Text
	alv := SubStr(xl.Range(col_alv . row).Text, -10)
	pro := SubStr("000" . xl.Range(col_pro . row).Text, -13)
	add := xl.Range(col_add . row).Text

    Sleep, %shortSleep%
    
	If (cod == 304 || cod == 386 || cod == 640 || cod == 681 || cod == 760 || cod == 1064 
			|| cod == 1065 || cod == 1066 || cod == 1067 || cod == 1083 || cod == 1161) {

		ControlSendRaw, , %arr%, ahk_pid %pwpid%
        Sleep, %shortSleep%
        
		ControlSend, , {Enter}, ahk_pid %pwpid%
        Sleep, 5000
        
		ControlSend, , {Enter}, ahk_pid %pwpid%
        Sleep, %shortSleep%
        
		
        ; MsgBox, 0, , Vai para posição do código
        Loop, 11 {
            ControlSend, , {Tab}, ahk_pid %pwpid%
            Sleep, %shortSleep%
        }
        
		ControlSend, , {End}, ahk_pid %pwpid%
        Sleep, %shortSleep%
        
		ControlSendRaw, , %cod%, ahk_pid %pwpid%
        Sleep, %shortSleep%
        
		ControlSend, , {F5}, ahk_pid %pwpid%
        Sleep, 5000
       
		gosub ConfirmationScreen
           
		; Now we update excel with date
		xl.Range(col_ret . row).Value := A_DD . "/" . A_MM . "/" . A_YYYY
		Sleep, 2000
    
	} Else If (cod == 478 || cod == 490) {
		; add vazio
		If (!add) {
			xl.Range(col_ret . row).Value := "err: CPF/CNPJ"
            Sleep, %shortSleep%
        
			xl.Range(row . ":" . row).Interior.ColorIndex := 6
            Sleep, %shortSleep%

			gosub NextRow
		}
        
		ControlSendRaw, , %arr%, ahk_pid %pwpid%
        Sleep, %shortSleep%
        
		ControlSend, , {Enter}, ahk_pid %pwpid%
        Sleep, 5000
        
		ControlSend, , {Enter}, ahk_pid %pwpid%
        Sleep, %shortSleep%
        
		ControlSend, , {End}, ahk_pid %pwpid%
        Sleep, %shortSleep%
        
		ControlSendRaw, , %add%, ahk_pid %pwpid%
        Sleep, %shortSleep%
        
        ; MsgBox, 0, , Vai para posição do código
        Sleep, %shortSleep%
        Loop, 11 {
            ControlSend, , {Tab}, ahk_pid %pwpid%
            Sleep, %shortSleep%
        }
        
		ControlSend, , {End}, ahk_pid %pwpid%
        Sleep, %shortSleep%
        
		ControlSendRaw, , %cod%, ahk_pid %pwpid%
        Sleep, %shortSleep%
        
		ControlSend, , {F5}, ahk_pid %pwpid%
        Sleep, 5000

		gosub ConfirmationScreen
        
		; Now we update excel with date
		xl.Range(col_ret . row).Value := A_DD . "/" . A_MM . "/" . A_YYYY
            Sleep, 2000
	
	} Else If (cod == 761) {

        xl.Range(col_ret . row).Value := "err: TODO(" . cod . ")"     
        Sleep, %shortSleep%
        
		xl.Range(row . ":" . row).Interior.ColorIndex := 6
        Sleep, %shortSleep%

    } Else {

        xl.Range(col_ret . row).Value := "err: INVÁLIDO(" . cod . ")"
        Sleep, %shortSleep%

		xl.Range(row . ":" . row).Interior.ColorIndex := 3
        Sleep, %shortSleep%

    }

NextRow:	 
    row += 1
    ; MsgBox, 0, , Próxima linha a executar: %row%
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

MsgBox, 0, %TITLE%, Fim da Execução

Exitapp

/*
 * Subrotines
 */

ConfirmationScreen:
{
	;msgbox, 0, , ConfirmationScreen Subrotine
	
	; First we clean every line
	Loop, 4 {
        ControlSend, , {End}, ahk_pid %pwpid%
        Sleep, %shortSleep%
        ControlSend, , {Tab}, ahk_pid %pwpid%
        Sleep, %shortSleep%
    }
    
	; Fill in observation
	ControlSendRaw, , apropriacao do alvara n. %alv%`, expedido nos autos do processo  n. %pro%., ahk_pid %pwpid%
    Sleep, %shortSleep%
    
	; Go to Processo field
	ControlSend, , {Tab}, ahk_pid %pwpid%
    Sleep, %shortSleep%
    
	; Fill in with Processão number
	ControlSendRaw, , %processo%, ahk_pid %pwpid%
    Sleep, %shortSleep%
    
	; Go to confirmation field
	ControlSend, , {Tab}, ahk_pid %pwpid%
    Sleep, %shortSleep%
    
	; Fill in confirmation
	ControlSendRaw, , %confirm_cmd%, ahk_pid %pwpid%
    Sleep, 1000
    
	ControlSend, , {Enter}, ahk_pid %pwpid%
    Sleep, %shortSleep%

} return

