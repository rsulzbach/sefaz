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

 
/*
 *	globals
 */
VERS := 1.123
TITLE := "Auto Alvarás - v" . VERS
shortSleep := 200
row := 0


/*
 *	autoexecute
 */
 
Gui, Add, Checkbox, vcfg_onlyPoa, Corrigir apenas Poa (096 e 900)

Gui, Add, Checkbox, Checked vcfg_confirm, Efetivar correções

Gui, Add, Button, gExecute vbtnExecute, Executar

Gui, Add, Button, gGuiClose vbtnClose, Fechar

Gui, Show, w250 h100 Center, %TITLE% . Config

Return


Execute:

;StartTicks := A_TickCount 

; Disable Controls
GuiControl, disable, cfg_onlyPoa
GuiControl, disable, cfg_confirm
GuiControl, disable, btnExecute
Gui, Submit, NoHide ;this command submits the guis' datas' state

If (cfg_confirm == 1) {
    confirm_cmd := "S"
} Else {
    confirm_cmd := "N"
}
	
MsgBox, 0, %TITLE%, Em seguida`, selecione o arquivo com a planilha dos Alvarás Automatizados.
Sleep, %shortSleep%

FileSelectFile, pathxl, , , Selecione Arquivo com Alvarás Automatizados, *.xlsx
Sleep, %shortSleep%

if (pathxl == "")
	Return

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
processao := xl.Range(_CELL_PROCESSAO).Text

MsgBox, 4, %TITLE%, O número do Processo Administrativo é %processao%
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
pro := "xxx" . SubStr(xl.Range(_COL_PROC . row).Text, -10)

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

InputBox, matr, %TITLE%, Digite sua matrícula para login no SOE:
InputBox, pwr, %TITLE%, Digite sua senha para login no SOE:, hide
Sleep, %shortSleep%

Run, C:\Program Files\pw3270\pw3270.exe, , , pwpid
WinWait, ahk_pid %pwpid%
Sleep, 333
Sleep, 4000

ControlSendRaw, , ims, ahk_pid %pwpid%
Sleep, %shortSleep%

ControlSend, , {Enter}, ahk_pid %pwpid%
Sleep, 2000

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

desviar("arr-alt-gui")

While row <= lastrow {

	; We start colecting all info from new row
	mun := SubStr(xl.Range(_COL_CGCTE . row).Text, 1, 3)
	arr := xl.Range(_COL_ARR . row).Text
	val := xl.Range(_COL_VAL . row).Text
	cod := xl.Range(_COL_COD . row).Text
	alv := SubStr(xl.Range(_COL_ALV . row).Text, -10)
	pro := "xxx" . SubStr(xl.Range(_COL_PROC . row).Text, -10)
	add := xl.Range(_COL_ADD . row).Text
	Sleep, %shortSleep%

	if cfg_onlyPoa && !(mun == 096 || mun == 900) {
		;MsgBox, 0, , Mun(%mun%) diferente de 096 ou 900.

		; Flags invalid mun
		xl.Range(_COL_RETURN . row).Value := "IE: " . mun . "/xxxxxxx"
	    Sleep, %shortSleep%
		; Paints row in yellow
		xl.Range(_COL_FIRST . row . ":" . _COL_LAST . row).Interior.ColorIndex := _FILLING_COLOR_WARNING
        Sleep, %shortSleep%

		goto NextRow
	}

	/*
	 * Caso Geral
	 */
	if HasVal(a_COD_GERAL, cod) {

		gosub ChangeGA		

		; Now we update excel with date
		xl.Range(_COL_RETURN . row).Value := A_DD . "/" . A_MM . "/" . A_YYYY
	    Sleep, %shortSleep%
		; Changes row cells Filling to NoFill
		xl.Range(_COL_FIRST . row . ":" . _COL_LAST . row).Interior.ColorIndex := _FILLING_COLOR_SUCCESS
    	Sleep, %shortSleep%

	/*
	 *	Exige Identificação do Contribuinte
	 */
	} Else If HasVal(a_COD_IDENT, cod) {
		; empty add field
		; can't change this GA without ID
		If (!add) {
			; update excel with warning
			xl.Range(_COL_RETURN . row).Value := "err: CPF/CNPJ"
            Sleep, %shortSleep%
			xl.Range(_COL_FIRST . row . ":" . _COL_LAST . row).Interior.ColorIndex := _FILLING_COLOR_WARNING
            Sleep, %shortSleep%

			goto NextRow
		}
        
		gosub ChangeGA
        
		; Now we update excel with date
		xl.Range(_COL_RETURN . row).Value := A_DD . "/" . A_MM . "/" . A_YYYY
	    Sleep, %shortSleep%
		; Changes row cells Filling to NoFill
		xl.Range(_COL_FIRST . row . ":" . _COL_LAST . row).Interior.ColorIndex := _FILLING_COLOR_SUCCESS
    	Sleep, %shortSleep%

	/*
	 *	COD 761
	 *	Related do DATs - uses another transaction
	 */
	} Else If (cod == 761) {
		; empty add field
		; can't change this GA without DAT
		If (!add) {
			; update excel with warning
			xl.Range(_COL_RETURN . row).Value := "err: DAT"
            Sleep, %shortSleep%
			xl.Range(_COL_FIRST . row . ":" . _COL_LAST . row).Interior.ColorIndex := _FILLING_COLOR_WARNING
            Sleep, %shortSleep%

			goto NextRow
		}

		gosub ChangeGA761

		; Now we update excel with date
		xl.Range(_COL_RETURN . row).Value := A_DD . "/" . A_MM . "/" . A_YYYY
	    Sleep, %shortSleep%
		; Changes row cells Filling to NoFill
		xl.Range(_COL_FIRST . row . ":" . _COL_LAST . row).Interior.ColorIndex := _FILLING_COLOR_SUCCESS
    	Sleep, %shortSleep%
 
    /*
	 *	An Unknown Cod
	 */
	} Else {
		; don't know what to do
		; update excel with error
		xl.Range(_COL_RETURN . row).Value := "err: INVÁLIDO(" . cod . ")"
		Sleep, %shortSleep%
		xl.Range(_COL_FIRST . row . ":" . _COL_LAST . row).Interior.ColorIndex := _FILLING_COLOR_ERROR
		Sleep, %shortSleep%

    }

NextRow:	 
	Sleep, 1000

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

/*
ElapsedTicks = A_TickCount - StartTicks
ElapsedTime = 20000101000000
ElapsedTime += ElapsedTicks/1000, Seconds

FormatTime fElapsedTime, %ElapsedTime%, HH:mm:ss

MsgBox, 0, %TITLE%, %fElapsedTime%
*/
MsgBox, 0, %TITLE%, Fim da Execução

Exitapp
return

/*
 *
 *	Subroutines
 *	
 */

/*
 */
ChangeGA:
{
	;msgbox, 0, , ChangeGA Subroutine

	; entering GA
	ControlSendRaw, , %arr%, ahk_pid %pwpid%
	Sleep, %shortSleep%

	; sometimes it takes a long time to enter the GA
	ControlSend, , {Enter}, ahk_pid %pwpid%
	Sleep, 5000

	ControlSend, , {Enter}, ahk_pid %pwpid%
	Sleep, %shortSleep%
    
	; We start at ID field
	; Changing ID Field if we have new info
	If (add) {
		ControlSend, , {End}, ahk_pid %pwpid%
		Sleep, %shortSleep%
		ControlSendRaw, , %add%, ahk_pid %pwpid%
		Sleep, %shortSleep%
	}
    
    ; MsgBox, 0, , Vai para posição do código
	Loop, 11 {
		ControlSend, , {Tab}, ahk_pid %pwpid%
		Sleep, %shortSleep%
    }
	
	; Changing Cod Field
	ControlSend, , {End}, ahk_pid %pwpid%
	Sleep, %shortSleep%
	ControlSendRaw, , %cod%, ahk_pid %pwpid%
	Sleep, %shortSleep%
	Sleep, 1000
	
	; Done changing things. Update
	ControlSend, , {F5}, ahk_pid %pwpid%
	Sleep, 4000
	
	gosub ConfirmationScreen

} return

/*
 */
ChangeGA761:
{
	;msgbox, 0, , ChangeGA761 Subroutine

	desviar("arr-alt-alv")

	ControlSendRaw, , %arr%, ahk_pid %pwpid%
	Sleep, %shortSleep%
	    
	ControlSend, , {Enter}, ahk_pid %pwpid%
	Sleep, 2000
	
	ControlSendRaw, , x, ahk_pid %pwpid%
	Sleep, %shortSleep%
	
	ControlSend, , {Enter}, ahk_pid %pwpid%
	Sleep, 5000
	
	ControlSend, , {Enter}, ahk_pid %pwpid%
	Sleep, %shortSleep%
	    
	; MsgBox, 0, , Vai para posição da Referencia
	ControlSend, , {Tab}, ahk_pid %pwpid%
	Sleep, %shortSleep%
	
	ControlSend, , {End}, ahk_pid %pwpid%
	Sleep, %shortSleep%
	
	ControlSendRaw, , %add%, ahk_pid %pwpid%
	Sleep, %shortSleep%
	
	; MsgBox, 0, , Vai para posição do código
	Sleep, %shortSleep%
	Loop, 10 {
		ControlSend, , {Tab}, ahk_pid %pwpid%
		Sleep, %shortSleep%
    }
	    
	ControlSend, , {End}, ahk_pid %pwpid%
	Sleep, %shortSleep%
	
	ControlSendRaw, , %cod%, ahk_pid %pwpid%
	Sleep, %shortSleep%
	Sleep, 1000
	
	ControlSend, , {F5}, ahk_pid %pwpid%
	Sleep, 4000
	
	gosub ConfirmationScreen
	
	; Go back to default transaction
	desviar("arr-alt-gui")

} return

/*
 */
ConfirmationScreen:
{
	;msgbox, 0, , ConfirmationScreen Subroutine
	
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
	ControlSendRaw, , %processao%, ahk_pid %pwpid%
    Sleep, %shortSleep%
    
	; Now PROA number fills the field, moving the cursor to the next field
	; Go to confirmation field
	;ControlSend, , {Tab}, ahk_pid %pwpid%
    ;Sleep, %shortSleep%
    
	; Fill in confirmation
	ControlSendRaw, , %confirm_cmd%, ahk_pid %pwpid%
    Sleep, 1000
    
	ControlSend, , {Enter}, ahk_pid %pwpid%
    Sleep, 2000

} return

/*
 */
GuiClose: 
ExitApp


/*
 *
 *	Functions
 *
 */

HasVal(haystack, needle)
{
	for index, value in haystack {
        if (value = needle)
            return index
	}

    if !(IsObject(haystack)) {
        throw Exception("Bad haystack!", -1, haystack)
	}
	    
	return 0
}

/*
 *	void	desviar(string transaction)
 *	DESCRIPTION :
 *		
 *	INPUTS :
 *		string transaction: transaction name 
 *
 *	OUTPUTS :
 *		void
 */
desviar(transaction)
{
	global
	
	Send, {LShift Down}
	Loop, 4 {
	    ControlSend, , {Tab 4}, ahk_pid %pwpid%
	}
	Send, {LShift Up}
	Sleep, %shortSleep%
	
	ControlSendRaw, , des, ahk_pid %pwpid%
	Sleep, %shortSleep%
	
	ControlSendRaw, , %transaction%, ahk_pid %pwpid%
	Sleep, %shortSleep%
	
	ControlSend, , {Tab}, ahk_pid %pwpid%
	Sleep, %shortSleep%
	
	ControlSend, , {Tab}, ahk_pid %pwpid%
	Sleep, %shortSleep%
	
	ControlSendRaw, , sar, ahk_pid %pwpid%
	Sleep, %shortSleep%
	
	ControlSend, , {Enter}, ahk_pid %pwpid%
	Sleep, %shortSleep%

	Sleep, 2000

	return
}

