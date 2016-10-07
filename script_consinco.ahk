/*
 * This script collects product data from an ERP system and inserts it into
 * another ERP via its graphic interface. It's an ugly script, with poor
 * coding/identation, because it was developed under pressure, but helped a lot
 * when changing the ERP system in a brazilian supermarket. Published for
 * learning/historical purposes.
 *
 * Copyright (C) 2015 João Dalben
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */


/*
 * Script para inserção automática de dados na interface de cadastro de famílias
 * e produtos do ERP Consinco.
 *
 * Esta ferramenta foi desenvolvida em Dezembro/2015 durante a época de implan-
 * tação do ERP Consinco nos Supermercados Dalben Ltda. Basicamente, um conjunto
 * de códigos de produto de um determinado ERP são passados como argumento, os
 * dados associados a esses produtos são coletados de uma  base de dados através
 * de um web server, e a inserção de dados na interface de cadastro do ERP
 * Consinco é feita de maneira automática.
 *
 * Como esse script foi escrito sob pressão, dado o curto prazo para conclusão
 * do cadastro de produtos, o código está pouco legível. Além disso, o percurso
 * pelos campos de inserção de dados é feito de maneira pouco softisticada
 * (gambiarra), e se baseia na distância de TABS entre os campos. A maneira
 * mais adequada para se percorrer os campos da interface de cadastro seria
 * através da função ControlFocus (pesquise sobre na referência do AHK).
 *
 * Entretanto, por escolha própria, decidi não mexer no código para fins de
 * preservação do programa original utilzado na época.
 *
 * Comentários por João Dalben, 7 de Outubro de 2016.
 *
 */

#NoEnv
#SingleInstance force
SendMode Input
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode 2
CoordMode, Mouse, Relative
SetTimer, Check, 20

/* Variáveis que identificam os títulos das janelas a serem ativadas */
Familia = familia - Bloco de notas
GR = GS-8420
Consinco = acrux mercari

/* Janela inicial. É solicitado que seja digitado o conjunto de códigos do
sistema antigo que irá compor a nova família. */
InitialTip = Digite os codigos dos produtos que irao compor a familia, separados por hífen (-):`n`nExemplo: 6587-7451-858784`n

Gui Add, Text,, %InitialTip%
Gui Add, Edit, hwndhProductList vProductList gOnChangeMyProductList Limit200 w400
Gui Add, Button, gValidateProductList Default, OK
Gui Show, W500
return

ValidateProductList:
	Gui Submit
	Gui destroy
	gosub ContinueAfterProductListInput

OnChangeMyProductList:
    Gui, Submit, NoHide
    NewText := RegExReplace(ProductList, "[^0-9-]", "")
    If NewText != %ProductList%
    {
        ControlGet, cursorPos, CurrentCol,, %ProductList%, A
        GuiControl, Text, ProductList, %NewText%
        cursorPos := cursorPos - 2
        SendMessage, 0xB1, cursorPos, cursorPos,, ahk_id %hProductList%
    }
return

ContinueAfterProductListInput:

/* Realiza uma consulta a um web server com os produtos passados.
 * É retornado um CSV contendo os dados do produto separados por ponto e
 * vírgula, e cada produto separado por arroba.
 * Exemplo:
 *
 * Descricao;ICMS;QTDCAIXA@Descricao;ICMS;QTDCAIXA;...
 */
product_list_tmp := StrSplit(ProductList, "-")
num_products_desired := product_list_tmp.Length()

MyUrl := "http://172.16.0.60/relatorios/transicao-consinco.php?produtos=" . ProductList
UrlDownloadToFile, %MyUrl%, C:\Users\usr121\Desktop\trans.txt
FileRead, filetext, C:\Users\usr121\Desktop\trans.txt

word_array := StrSplit(filetext, "@", "@")
word_array.Pop()

family := {}


/* Os dados do arquivos CSV são alocados em um vetor de produtos, e cada produto
 * é um vetor de dados contendo todas as informações a serem transferidas
 * para o novo sistema.
 */

Produtos_f = Produtos que irão compor a familia:`n
for i, e in word_array
{
	tmp := StrSplit(e, ";", ";")
	tmp.pop()

	p := {}
	p.code := tmp.1
	p.description := tmp.2
	p.ncm := tmp.3
	p.ipi := tmp.4
	p.percipi := tmp.5
	p.pis_in := tmp.6
	p.pis_out := tmp.7
	p.cofins_in := tmp.8
	p.cofins_out := tmp.9
	p.tabela := tmp.10
	p.receita := tmp.11
	p.qtd_cx := tmp.12
	p.comprador := tmp.13
	p.margem := tmp.14
	p.dtcad := tmp.15
	p.codembcompra := tmp.16
	p.tipoembcompra := tmp.17
	p.razaoforn := tmp.18
	p.cpf1 := tmp.19
	p.cpf2 := tmp.20
	p.emb := {}
	p.ref := {}
        
        /* Os códigos de barra são alocados em um vetor, que é apontado como
         * um elemento do vetor produto.
         */
	
	MyUrlBarras := "http://172.16.0.60/relatorios/transicao-consinco-barra-produto.php?produto=" . p.code
	UrlDownloadToFile, %MyUrlBarras%, C:\Users\usr121\Desktop\trans-barra-produto.txt
	FileRead, barstext, C:\Users\usr121\Desktop\trans-barra-produto.txt
	bar_array := StrSplit(barstext, "@", "@")
	bar_array.Pop()
	
	for j, f in bar_array {
		tmpbar := StrSplit(f, ";", ";")
		tmpbar.pop()		
		
		b := {}
		b.Push(tmpbar.1)
		b.Push(tmpbar.2)
		p.emb.Push(b)
	}

        /* Operação idêntica ao código de barras, mas para as referências 
         * (código do produto no sistema do fornecedor).
         */
	
	MyUrlRef := "http://172.16.0.60/relatorios/transicao-consinco-referencia-produto.php?produto=" . p.code
	UrlDownloadToFile, %MyUrlRef%, C:\Users\usr121\Desktop\trans-ref-produto.txt
	FileRead, refstext, C:\Users\usr121\Desktop\trans-ref-produto.txt
	ref_array := StrSplit(refstext, "@", "@")
	ref_array.Pop()
	
	for j, f in ref_array {
		tmpref := StrSplit(f, ";", ";")
		tmpref.pop()
		
		r := {}
		r.Push(tmpref.1)
		r.Push(tmpref.2)
		p.ref.Push(r)
	}

	Produtos_f := Produtos_f . "`n->  " . p.code . " - " . p.description
	family.Push(p)
}


num_products_family := family.Length()
if (num_products_desired != num_products_family) {
	MsgBox Algum dos codigos que voce digitou nao existe no GR. Por favor, revise os codigos informados.
	ExitApp
	Return
}

/*
 * O usuário escolhe o nome da nova família
 */

Produtos_f = %Produtos_f%`n`nEscolha o nome da família (máximo 40 caracteres dentre 0-9 a-z A-Z   ' / - ,):

desc35char := p.description 
StringMid, desc35char, desc35char, 1, 35

Gui Add, Text,, %Produtos_f%
Gui Add, Edit, hwndhFamilyName vFamilyName gOnChangeMyText Limit35 w400, %desc35char%
Gui Add, Button, gValidate Default, OK
Gui Show, W500
return

Validate:
	Gui Submit
	Gui destroy
	gosub StartFillingData


OnChangeMyText:
    Gui, Submit, NoHide ; Get the info entered in the GUI
    NewText := RegExReplace(FamilyName, "[^0-9a-zA-Z' ,&%/.""]", "") ;Allow digits letters underscore_ apostrophe' space dash-
    If NewText != %FamilyName% ; Check if any invalid characters were removed
    {
        ControlGet, cursorPos, CurrentCol,, %FamilyName%, A ; Get current cursor position
        GuiControl, Text, FamilyName, %NewText% ; Write text back to Edit control
        cursorPos := cursorPos - 2
        SendMessage, 0xB1, cursorPos, cursorPos,, ahk_id %hFamilyName% ; EM_SETSEL ; Add hwndh to Edit control for this to work
    }
return

/*
 * A partir deste rótulo (StartFillingData), é realizada a inserção dos dados no
 * sistema Consinco. Basicamente um dado é escrito no campo correspondente
 * e N tabs são dados para que o script pule para o próximo campo desejado,
 * onde a informação apropriada é inserida, e assim por diante.
 *
 * Este é um método pouco sofisticado de alternância de campos numa janela.
 * Se um novo campo for inserido numa versão futura, o número de tabs entrados
 * para alternar entre campos será diferente, e o script irá falhar.
 *
 * O método mais adequado seria utilizar o Id de cada campo em conjunto com
 * a função ControFocus (pesquise sobre na referência do AHK).
 */
StartFillingData:
MudaJanela(Consinco)
SendInput %FamilyName%
SendInput {Tab}
SendInput {Enter}
SendInput +{Tab}

KeyWait Control, D
Sleep 200
SendInput +{Tab}
SendInput {Right}

/*
****************** IMPOSTOS ******************
*/

/*
>>> NCM
*/
NKeys("{Tab}", 4)
SendInput % p.ncm


/*
>>> IPI
*/
If (p.ipi = 50)
{
	If (p.percipi != 0) {
		NKeys("{Tab}", 4)
		SendInput % p.percipi
		NKeys("{Tab}", 12)
	}
} else {
	NKeys("{Tab}", 16)UNs			
}
SendInput % p.ipi
NKeys("{Tab}", 1)
SendInput % p.ipi
NKeys("{Tab}", 3)

/*
>>> PIS E COFINS
*/
SendInput % p.pis_in
NKeys("{Tab}", 1)
SendInput % p.cofins_in
qtd_tabs := 1
If (p.pis_in >= 50 and p.pis_in <= 56)
	qtd_tabs++
	
If (p.cofins_in >= 50 and p.cofins_in <= 56)
	qtd_tabs++
	
NKeys("{Tab}", qtd_tabs)
SendInput % p.pis_out
NKeys("{Tab}", 1)
SendInput % p.cofins_out

If ((p.pis_out >= 02 and p.pis_out <= 09) or (p.cofins_out >= 02 and p.cofins_out <= 09))
{
NKeys("{Tab}", 1)
SendInput {Enter}
MsgBox % "Selecione:`n`nTabela: " . p.tabela . "`nCodigo: " . p.receita
KeyWait Control, D
}

NKeys("{Tab}", 14)
SendInput {F4}
KeyWait Control, D

/*
****************** EMBALAGENS ******************
*/

SendInput {Right}
If (p.qtd_cx = 1) {
	NKeys("{Insert}", 1)
	SendInput % p.tipoembcompra
} else {
NKeys("{Insert}", 2)
SendInput % p.qtd_cx
NKeys("{Tab}", 1)
SendInput % p.tipoembcompra
NKeys("{Tab}", 6)
SendInput % p.qtd_cx
}
SendInput {F4}
Sleep 700
NKeys("{Tab}", 1)
Sleep 100
NKeys("{Tab}", 1)
Sleep 100
NKeys("{Tab}", 1)
Sleep 100
NKeys("{Tab}", 1)
KeyWait Control, D

/*
****************** FORNECEDORES ******************
*/

SendInput {Right}
SendInput {Tab}
SendInput {Insert}
NKeys("{Tab}", 2)
SendInput % p.cpf1
SendInput {Tab}
SendInput % p.cpf2
SendInput {F8}
KeyWait Control, D
Sleep 200
SendInput {F4}
Sleep 200
SendInput {Tab}
SendInput {Enter}

/*
****************** DIVISAO ******************
*/

SendInput % p.comprador
NKeys("{Tab}", 2)
SendInput {Down}
SendInput {Tab}
SendInput {Down}
NKeys("{Tab}", 3)
SendInput {Down}
NKeys("{Tab}", 2)
SendInput % p.margem
NKeys("{Tab}", 4)
SendInput {Right}
NKeys("{Tab}", 2)
SendInput {Space}
SendInput {F2}
Tributacao := GetTributacaoMaisComum(family)
SendInput % Tributacao
SendInput {F8}
SendInput {Enter}
NKeys("+{Tab}", 9)


/*
****************** NACIONAL/IMPORTADO ******************
*/

din_height := 130
barras_nac =
for i, e in family {
	for j, f in e.emb {
			barras_nac := barras_nac . "->  " . f.1 . " (" . f.2 . ")`n"
			din_height := din_height + 12
		}
}

Gui, Add, Text, x10 y15, Escolha a nacionalidade do produto baseado nos codigos `nde barras dos produtos desta família.
for i, e in family
	for j, f in e.emb
		for k, g in f
			Gui Add, Text, x20 y60, %barras_nac%

opts := din_height -45
okbutton := din_height - 30
Gui, Add, Radio, x30 y%opts%  vNacionalOpt Checked, Nacional
Gui, Add, Radio, x190 y%opts%  vImportadoOpt , Estrangeiro
Gui, Add, Button, x100 y%okbutton% w70 h22 gNacionalidadeOK, Ok
Gui, Show, x386 y160 w320 h%din_height% , Page Setup
Return

NacionalidadeOK:
	Gui, submit
	Gui, destroy
	gosub PosNac
	
PosNac:
If (NacionalOpt = 1)
	XScrollDownNac := 1
else
	XScrollDownNac := 3

NKeys("{Down}", XScrollDownNac)
NKeys("{Tab}", 7)
SendInput {Left}


KeyWait Control, D
SendInput {F3}
SendInput {F10}


/*
****************** LOCAL VENDA/LISTA ******************
*/

NKeys("{Tab}", 4)
SendInput {Right}
SendInput {Tab}
NKeys("{Down}", 2)
SendInput {Insert}
SendInput {Tab}
SendInput {Space}
SendInput {Down}
SendInput {F4}
NKeys("{Tab}", 4)
SendInput {Right}
NKeys("{Tab}", 4)
SendInput {Insert}
SendInput {Space}
NKeys("{Tab}", 4)
SendInput {F4}
Sleep 500
NKeys("{Tab}", 2)
Sleep 500
SendInput {Space}


/*
****************** LOOP PRODUTOS ******************
*/

for z, prod in family {
SendInput {Space}
Complement := ""
GenericDescription := ""
ReducedDescription := ""


/*
****************** DESCRICOES ******************
*/

Complement := GetComplement(FamilyName, prod.description)
StringUpper, Complement, Complement
GenericDescription := GetGenericDescription(FamilyName, Complement)
StringUpper, GenericDescription, GenericDescription
ReducedDescription := GetReducedDescription(FamilyName, Complement, GenericDescription)
StringUpper, ReducedDescription, ReducedDescription

ch_ref := ""
refs_opts := {}
tmp_text := ""
for i, f in prod.ref {
			tmp_text := f.1 . " - " . f.2
			refs_opts.Push(tmp_text)
}
refs_text := "Fornecedor atual no GR:`n" . prod.razaoforn 
ch_ref := NRadios(refs_opts, "Escolha uma referencia", refs_text)

ch_util_venda := 1
util_venda_text := {}
tmp_text := ""
for i, f in prod.emb {
			tmp_text := f.2 . " - " . f.1
			util_venda_text.Push(tmp_text)
}

ch_util_venda := NRadios(util_venda_text, "Escolha um codigo de barras para marcar como Util.Venda", "Escolha a barra para ser utilizada para venda.`nA ultima barra vendida foi a " GetBarraUltimaVenda(prod.code) ".")

MudaJanela(Consinco)
SendInput {F2}
SendInput {Tab}
SendInput % Complement
NKeys("{Tab}", 1)
SendInput ^+{End}
SendInput {Del}
SendInput % GenericDescription
SendInput {Tab}
SendInput % ReducedDescription

SendInput {F3}
Sleep 500
SendInput ^{Tab}

/*
****************** CODIGOS ******************
*/

/*
*** REFERENCIA
*/

KeyWait Control, D
SendInput {Tab}
SendInput {Insert}
SendInput {T}
SendInput {Tab}
SendInput {U}
SendInput {Tab}
SendInput % prod.code
SendInput {Tab}
SendInput % prod.dtcad

/*
*** REFERENCIA
*/

SendInput {Insert}
NKeys("{F}", 2)
SendInput {Tab}
if (prod.tipoembcompra = "CX")
	NKeys("{C}", 2)
else if (prod.tipoembcompra = "FD")
	NKeys("{F}", 2)
else if (prod.tipoembcompra = "UN")
	NKeys("{U}", 2)
else if (prod.tipoembcompra = "PC")
	NKeys("{P}", 2)
else {
		MsgBox Tipo de unidade utilizada para venda nao é CX, UN ou FD. Encerrando o script. Avise o Joao sobre esse erro.
		ExitApp
		Return
}
SendInput {Tab}
SendInput % prod.ref[ch_ref].1

/*
*** UTIL VENDA
*/

util_venda_len := Strlen(prod.emb[ch_util_venda].1)

SendInput {Insert}
if (util_venda_len >= 8 and ult_venda_len <= 13)
	NKeys("{E}", 2)
else if (util_venda_len = 14)
	NKeys("{D}", 2)
else {
	MsgBox Tipo de codigo da embalagem de compra nao é EAN nem DUN. Encerrando o script. Avise o Joao sobre esse erro.
	ExitApp
	Return
}
SendInput {Tab}

if (prod.emb[ch_util_venda].2 = "CX")
	NKeys("{C}", 2)
else if (prod.emb[ch_util_venda].2 = "FD")
	NKeys("{F}", 2)
else if (prod.emb[ch_util_venda].2 = "UN")
	NKeys("{U}", 2)
else if (prod.emb[ch_util_venda].2 = "PC")
	NKeys("{P}", 2)
else {
		MsgBox Tipo de unidade utilizada para vendo nao é CX, UN ou FD. Encerrando o script. Avise o Joao sobre esse erro.
		ExitApp
		Return
}
	
SendInput {Tab}
SendInput % prod.emb[ch_util_venda].1
SendInput {F4}
Sleep 500
SendInput {Down}
SendInput {Up}
SendInput {Insert}
SendInput {Space}
SendInput {F4}

/*
*** EMBALAGEM DE COMPRA
*/

emb_compra_len := Strlen(prod.codembcompra)

if (emb_compra_len = 14) {
	SendInput {Insert}
	SendInput {Insert}
	NKeys("{D}", 2)
	SendInput {Tab}
	
	if (prod.tipoembcompra = "CX")
		NKeys("{C}", 2)
	else if (prod.tipoembcompra = "FD")
		NKeys("{F}", 2)
	else {
			MsgBox Tipo de unidade utilizada para vendo nao é CX nem FD. Encerrando o script. Avise o Joao sobre esse erro.
			ExitApp
			Return
	}
	SendInput {Tab}
	SendInput % prod.codembcompra
}


KeyWait Control, D
SendInput {F4}
Sleep 500
SendInput {F10}
Sleep 500
NKeys("{Tab}", 2)
}

SendInput {F10}
MsgBox Fim do cadastro da familia

ExitApp
Return

NCheckBoxes(opts, wintitle)
{
global
Loop, % opts.Length() {   ; set some random values for the vars
    Random, n, 0, 1
    Var%A_Index% := n
}
   
Loop, % opts.Length()
{
    Gui, Add, Checkbox, % "x10 y+10 checked" Var%A_Index% " vVar" A_Index, Checkbox %A_Index%
}

Gui, Add, Button, gButton1, Submit
Gui, Show, w300, % wintitle
WinWaitClose, % wintitle
return Output

Button1:
Gui, Submit
Gui Destroy
Output := ""
Loop, % opts.Length()
    Output .= Var%A_Index%
return

GuiClose:
ExitApp
}

NRadios(opts, wintitle, wintext)
{
global
if(wintext = "")
	wintext := "Escolha uma das opcoes:"
Gui, Add, Text,, %wintext%
tmp := opts.RemoveAt(1)
Gui, Add, Radio, altsubmit vChosenRadio Checked, %tmp%
for i, e in opts
	Gui, Add, Radio, altsubmit, %e%
Gui Add, Button, gGetChosenRadio Default, OK
Gui, Show, w500, % wintitle
WinWaitClose, % wintitle
Return ChosenRadio

GetChosenRadio:
Gui Submit
Gui Destroy
Return
}


NKeys(k, n)
{
	Loop, %n%
	{
		SendInput %k%
		Sleep, 80
	}

}

MudaJanela(j)
{
	IfWinExist, %j% 
	{
		WinActivate ; use the window found above
		WinWait, %j%, , 1
		If ErrorLevel
		{
			MsgBox, % "Nao foi possivel mudar para a janela" . %j%
		}
	}
	else
	{
		MsgBox % "A janela " . j . " nao esta aberta. Por favor, abra-a antes de prosseguir."
		ExitApp
		return
	}
}

GetComplement(f_name, desc) {
global
tituloComp := "Escolha o complemento"
Gui Add, Text,x10 y20, % "Digite o complemento do produto, baseado na descrição atual do GR:`n(máximo 15 caracteres dentre 0-9 a-z A-Z   ' / - , %wildcard%)`n`n-> " . desc
Gui Add, Edit, x10 y77 w250 hwndhComplement vComplement gOnChangeMyComplement Limit15
Gui, Add, Text, x10 y100, `nDescrição completa (nome da família + complemento):
Gui, Add, Text, x10 y130 w400 vTypedComplement, -> %f_name%
Gui Add, Button, x190 y160 w70 h25 gValidateComplement Default, OK
Gui Show, h200 W460, % tituloComp
SendInput {Tab}
SendInput {Tab}
WinWaitClose % tituloComp
return Complement

ValidateComplement:
	Gui Submit
	Gui destroy
	StringUpper, Complement, Complement
return


OnChangeMyComplement:
    Gui, Submit, NoHide ; Get the info entered in the GUI
    NewComplement := RegExReplace(Complement, "[^0-9a-zA-Z' ,&""%/.]", "") ;Allow digits letters underscore_ apostrophe' space dash-
    If NewComplement != %Complement% ; Check if any invalid characters were removed
    {
        ControlGet, cursorPos, CurrentCol,, %Complement%, A ; Get current cursor position
        GuiControl, Text, Complement, %NewComplement% ; Write text back to Edit control
        cursorPos := cursorPos - 2
        SendMessage, 0xB1, cursorPos, cursorPos,, ahk_id %hComplement% ; EM_SETSEL ; Add hwndh to Edit control for this to work
    }
	StringUpper, NewComplement, NewComplement
	GuiControl, Text, TypedComplement, % "-> " . f_name . " " . NewComplement ;
return
}

GetGenericDescription(f_name, comp) {
global
tituloGen := "Escolha uma descricao generica"
Gui Add, Text,x10 y20, % "Digite a descrição genérica do produto baseada na descrição completa`n(máximo 60 caracteres dentre 0-9 a-z A-Z   ' / - ,)`n`n-> " . f_name . " " . comp
Gui Add, Edit, x10 y77 w440 hwndhGenericDescription vGenericDescription gOnChangeMyGenericDescription Limit60, %f_name% %comp%
Gui Add, Button, x190 y160 w70 h25 gValidateGenericDescription Default, OK
Gui Show, h200 W460, % tituloGen
WinWaitClose % tituloGen
return GenericDescription

ValidateGenericDescription:
	Gui Submit
	StringUpper, GenericDescription, GenericDescription
	Gui destroy
return

OnChangeMyGenericDescription:
    Gui, Submit, NoHide ; Get the info entered in the GUI
    NewGenericDescription := RegExReplace(GenericDescription, "[^0-9a-zA-Z' ,&%""/.]", "") ;Allow digits letters underscore_ apostrophe' space dash-
    If NewGenericDescription != %GenericDescription% ; Check if any invalid characters were removed
    {
        ControlGet, cursorPos, CurrentCol,, %GenericDescription%, A ; Get current cursor position
        GuiControl, Text, GenericDescription, %NewGenericDescription% ; Write text back to Edit control
        cursorPos := cursorPos - 2
        SendMessage, 0xB1, cursorPos, cursorPos,, ahk_id %hGenericDescription% ; EM_SETSEL ; Add hwndh to Edit control for this to work
    }
return
}

GetReducedDescription(f_name, comp, gen_desc) {
global
tituloRed := "Escolha a descricao reduzida"
StringMid, ReducedDescription, GenericDescription, 1, 24
Gui Add, Text,x10 y20, % "Digite a descrição reduzida do produto`n(máximo 24 caracteres dentre 0-9 a-z A-Z   ' / - ,)`n`n-> Descrição Completa: " . f_name . " " . comp . "`n-> Descrição Genérica: " . gen_desc
Gui Add, Edit, x10 y97 w440 hwndhReducedDescription vReducedDescription gOnChangeMyReducedDescription Limit24, %ReducedDescription%
Gui Add, Button, x190 y147 w70 h25 gValidateReducedDescription Default, OK
Gui Show, h200 W460, % tituloRed
WinWaitClose % tituloRed
return ReducedDescription

ValidateReducedDescription:
	Gui Submit
	StringUpper, ReducedDescription, ReducedDescription
	Gui destroy
return
	
OnChangeMyReducedDescription:
    Gui, Submit, NoHide ; Get the info entered in the GUI
    NewReducedDescription := RegExReplace(ReducedDescription, "[^0-9a-zA-Z' "",&%/.]", "") ;Allow digits letters underscore_ apostrophe' space dash-
    If NewReducedDescription != %ReducedDescription% ; Check if any invalid characters were removed
    {
        ControlGet, cursorPos, CurrentCol,, %ReducedDescription%, A ; Get current cursor position
        GuiControl, Text, ReducedDescription, %ReducedDescription% ; Write text back to Edit control
        cursorPos := cursorPos - 2
        SendMessage, 0xB1, cursorPos, cursorPos,, ahk_id %hReducedDescription% ; EM_SETSEL ; Add hwndh to Edit control for this to work
    }
return
}

GetTributacaoMaisComum(familia) {
EANList := ""
NCMList := ""
for i, e in familia {
	for j, f in e.emb {
		barra_len := Strlen(f.1)
		if (barra_len >= 8 and barra_len <= 13)
			EANList := EANList . f.1 . "-"
	}
	NCMList := NCMList . e.ncm . "-"
}

StringTrimRight, EANList, EANList, 1
StringTrimRight, NCMList, NCMList, 1

UrlTrib := "http://172.16.0.60/trans-consinco/tributacao-mais-comum.php?barras=" . EANList . "&ncm=" . NCMList
UrlDownloadToFile, %UrlTrib%, C:\Users\usr121\Desktop\trans-tributacao.txt
FileRead, CodTributacao, C:\Users\usr121\Desktop\trans-tributacao.txt
return CodTributacao
}

GetBarraUltimaVenda(product) {
UrlTrib := "http://172.16.0.60/trans-consinco/barra-ultima-venda.php?produto=" . product
UrlDownloadToFile, %UrlTrib%, C:\Users\usr121\Desktop\trans-barra-ult-venda.txt
FileRead, UltBarra, C:\Users\usr121\Desktop\trans-barra-ult-venda.txt
return UltBarra
}	

Check:
MouseGetPos, xx, yy
Tooltip %xx%`, %yy%
return

Escape::
MsgBox O script foi encerrado pela tecla Esc.
ExitApp
Return


