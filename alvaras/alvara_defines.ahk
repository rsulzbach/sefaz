/*
 *
 *
 */

;MsgBox, 0, , alvara_defines.ahk included

/*
 * Excel
 */
_EXCEL_COLOR_NOFILL := 0
_EXCEL_COLOR_WHITE := 2
_EXCEL_COLOR_RED := 3
_EXCEL_COLOR_YELLOW := 6

/*
 * Meu Arquivo Excel: Diversos
 */
_FILLING_COLOR_SUCCESS := _EXCEL_COLOR_NOFILL
_FILLING_COLOR_WARNING := _EXCEL_COLOR_YELLOW 
_FILLING_COLOR_ERROR := _EXCEL_COLOR_RED

/*
 * Meu Arquivo Excel: Colunas
 */
_COL_CGCTE := "A"
_COL_VAL := "J"
_COL_ARR := "E"
_COL_ALV := "G"
_COL_RETURN := "K"
_COL_COD := "L"
_COL_PROC := "C"
_COL_ADD := "M"

_COL_FIRST := _COL_CGCTE
_COL_LAST := _COL_ADD


/*
 * Meu Arquivo Excel: Células
 */
_CELL_PROCESSAO := "G2"


/*
 * Lista de Códigos do Caso Geral
 */
a_COD_GERAL := []
a_COD_GERAL.Push(304)
a_COD_GERAL.Push(330)
a_COD_GERAL.Push(386)
a_COD_GERAL.Push(451)
a_COD_GERAL.Push(479)
a_COD_GERAL.Push(547)
;a_COD_GERAL.Push(640)	; PROCON  
a_COD_GERAL.Push(643)
a_COD_GERAL.Push(681)
a_COD_GERAL.Push(760)
a_COD_GERAL.Push(762)
a_COD_GERAL.Push(942)
a_COD_GERAL.Push(978)
a_COD_GERAL.Push(1008)
a_COD_GERAL.Push(1065)
a_COD_GERAL.Push(1066)
a_COD_GERAL.Push(1067)
a_COD_GERAL.Push(1199)


/*
 * Lista de Códigos que Exige Identificação do Contribuinte
 */
a_COD_IDENT := []
a_COD_IDENT.Push(305)
a_COD_IDENT.Push(319)
a_COD_IDENT.Push(378)
a_COD_IDENT.Push(478)
;a_COD_IDENT.Push(490)	; manual - necessário também cod município
a_COD_IDENT.Push(1047)
a_COD_IDENT.Push(1064)
a_COD_IDENT.Push(1074)
a_COD_IDENT.Push(1083)	; Extinta CEERGS - BADESUL 
a_COD_IDENT.Push(1084)	; Extinta CEERGS - BADESUL 
a_COD_IDENT.Push(1100)
a_COD_IDENT.Push(1161)	; Extinta CEERGS - BADESUL 
a_COD_IDENT.Push(1162)	; Extinta CEERGS - BADESUL
a_COD_IDENT.Push(1587)

