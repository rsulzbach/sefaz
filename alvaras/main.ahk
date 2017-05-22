/*
 *
 *
 */

/*
 *	include
 */
;MsgBox, 0, , dir = %A_ScriptDir%
#include alvara_defines.ahk


/*
 *	configure
 */
_CONFIG_CONFIRM_CHANGE = 0


/*
 *	globals
 */
VERS := 1.010
TITLE := "Alvarás Automatizados - " . VERS
shortSleep := 200
row := 0


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

	mun := SubStr(xl.Range(_COL_CGCTE . row).Text, 1, 3)
	arr := xl.Range(_COL_ARR . row).Text
	val := xl.Range(_COL_VAL . row).Text
	cod := xl.Range(_COL_COD . row).Text
	alv := SubStr(xl.Range(_COL_ALV . row).Text, -10)
	pro := SubStr("000" . xl.Range(_COL_PROC . row).Text, -13)
	add := xl.Range(_COL_ADD . row).Text

    Sleep, %shortSleep%

	if !(mun == 096 || mun == 900) {
		;MsgBox, 0, , Mun(%mun%) diferente de 096 ou 900.

		; Flags invalid mun
		xl.Range(_COL_RETURN . row).Value := "IE: " . mun . "/xxxxxxx"
		; Paints row in yellow
		xl.Range(row . ":" . row).Interior.ColorIndex := 6
        Sleep, %shortSleep%

		goto NextRow
	}

	If (cod == 304 || cod == 386 || cod == 640 || cod == 681 || cod == 760
			|| cod == 1064 || cod == 1065 || cod == 1066 || cod == 1067
			|| cod == 1083 || cod == 1161) {

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
		xl.Range(_COL_RETURN . row).Value := A_DD . "/" . A_MM . "/" . A_YYYY
		Sleep, 2000
    
	} Else If (cod == 478 || cod == 490) {
		; add vazio
		If (!add) {
			xl.Range(_COL_RETURN . row).Value := "err: CPF/CNPJ"
            Sleep, %shortSleep%
        
			xl.Range(row . ":" . row).Interior.ColorIndex := 6
            Sleep, %shortSleep%

			goto NextRow
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
		xl.Range(_COL_RETURN . row).Value := A_DD . "/" . A_MM . "/" . A_YYYY
            Sleep, 2000
	
	} Else If (cod == 761) {

        xl.Range(_COL_RETURN . row).Value := "err: TODO(" . cod . ")"     
        Sleep, %shortSleep%
        
		xl.Range(row . ":" . row).Interior.ColorIndex := 6
        Sleep, %shortSleep%

    } Else {

        xl.Range(_COL_RETURN . row).Value := "err: INVÁLIDO(" . cod . ")"
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
 *	Subrotines
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

