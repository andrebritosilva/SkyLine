#include "protheus.ch" 
#Include "RWMAKE.CH" 

//-------------------------------------------------------------------
/*/{Protheus.doc} ETX_001
Extrato Bancário - SkyLine

@author André Brito
@since 26/06/2019
@version P12
/*/
//-------------------------------------------------------------------
           
User Function ETX_001()     

Local aPWiz       := {}
Local aRetWiz     := {}
Local cTmpRng     := ""

Private cFilDe    := ""
Private cFilAte   := "" 
Private cBanco    := ""
Private cAgencia  := "" 
Private cMark     := GetMark()
Private aRotina	  := MenuDef()
Private dDtDe      
Private dDtAte                    
Private oMark 
Private oProcess
Private aSelFil   := AdmGetFil()
Private cRng      := ""
Private cPrefixos := ""
Private cBancos   := ""
Private lXConti   := .F.
Private nTotSal   := 0
Private cTotSal   := ""
Private cAgens    := ""
Private cCtas     := ""
Private nXValTot  := 0

aAdd(aPWiz,{ 1,"Data de: "               ,Ctod("")                        ,"","",""   ,  ,60  ,.T.})
aAdd(aPWiz,{ 1,"Data ate: "              ,Ctod("")                        ,"","",""   ,  ,60  ,.T.})

aAdd(aRetWiz,Ctod(""))
aAdd(aRetWiz,Ctod(""))

ParamBox(aPWiz,"***** Extrato Bancário - SKYLINE *****",@aRetWiz) 

dDtDe      := aRetWiz[1] 
dDtAte     := aRetWiz[2] 

cRng := GetRngFil(aSelFil,"SE5",.T.,@cTmpRng)

cPrefixos := xBusPre()

cBancos := xBusBan()

oProcess := MsNewProcess():New( { || XProcMan(dDtDe, dDtAte, cPrefixos)})
oProcess:Activate()

//-------------------------------------------------------------------
/*/{Protheus.doc} XProcMan
Montagem de tabela temporária de acordo com os parametros informados,
usando o MSSelect.

@author André Brito
@since 26/06/2019
@version P12
/*/
//-------------------------------------------------------------------

Static Function XProcMan(dDtDe, dDtAte, cPrefixos)
 
Local aCpoBro     := {} 
Local aCampos     := {} 
Local aCores      := {}
Local aSize       := {}
Local cArqMov     := GetNextAlias()
Local cAliMov     := GetNextAlias() 
Local cQuery      := ""
Local lCheck      := .F.
Local aStru		  := SE5->(DbStruct())
Local aColumns	  := {}
Local nX          := 0
Local nVal        := 0
Local oPanel 
Local oDlgLocal
Local oSay1	 
Local _oConMan
Local oCheck1 
Local oChk

Private oBrowse

AADD(aCampos,{"E5_DATA"        ,"C",TamSX3("E5_DATA")[1]   ,0})
AADD(aCampos,{"E5_BENEF"       ,"C",TamSX3("E5_BENEF")[1]  ,0})
AADD(aCampos,{"E5_BANCO"       ,"C",80,0})
AADD(aCampos,{"E5_CONTA"       ,"C",TamSX3("E5_CONTA")[1]  ,0})
AADD(aCampos,{"E5_AGENCIA"     ,"C",TamSX3("E5_AGENCIA")[1]  ,0})
AADD(aCampos,{"E5_HISTOR"      ,"C",TamSX3("E5_HISTOR")[1] ,0})
AADD(aCampos,{"E5_NATUREZ"     ,"C",80,0})
AADD(aCampos,{"E5_FILIAL"      ,"C",40,0})
AADD(aCampos,{"E5_VALOR"       ,"N",TamSX3("E5_VALOR")[1]  ,2})
AADD(aCampos,{"E5_RECPAG"      ,"C",TamSX3("E5_RECPAG")[1]  ,0})
AADD(aCampos,{"E5_SALDO"       ,"C",26,0})
AADD(aCampos,{"E5_DTDISPO"     ,"C",TamSX3("E5_DTDISPO")[1]  ,0})
AADD(aCampos,{"E5RECNO"        ,"N",08  ,0})

cQuery := "SELECT R_E_C_N_O_, * FROM "
cQuery += RetSqlName("SE5") + " SE5 "
cQuery += " WHERE "
cQuery += " E5_FILIAL " + cRng + " "
cQuery += " AND E5_TIPODOC NOT IN ('BA','DC','JR','MT','CM','D2','J2','M2','C2','V2','CP','TL','VA')
cQuery += " AND E5_TIPODOC != 'TR' AND E5_TIPODOC != 'TE' "
cQuery += " AND E5_BANCO IN (" + cBancos + ")"
cQuery += " AND E5_AGENCIA IN (" + cAgens + ")"
cQuery += " AND E5_CONTA IN (" + cCtas + ")"
cQuery += " AND ( E5_PREFIXO IN (" + cPrefixos + ")"
cQuery += " OR E5_PREFIXO = ' ' )"
cQuery += " AND E5_DATA   >= '" + Dtos(dDtDe)  + "' AND E5_DATA <= '" + Dtos(dDtAte) + "' " 
cQuery += " AND D_E_L_E_T_ = ' ' "
	
cQuery := ChangeQuery(cQuery) 
 
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliMov,.T.,.T.)

If _oConMan <> Nil
	_oConMan:Delete() 
	_oConMan := Nil
EndIf

// Criando o objeto do arquivo temporário
_oConMan := FwTemporaryTable():New("cArqMov")

// Criando a estrutura do objeto  
_oConMan:SetFields(aCampos)

// Criando o indice da tabela
_oConMan:AddIndex("1",{"E5RECNO"})
//_oConMan:AddIndex("1",{"E5_FILIAL","E5_DATA","E5_BANCO", "E5_AGENCIA", "E5_CONTA"})
//_oConMan:AddIndex("1",{"E5_FILIAL","E5_DATA","E5_BANCO", "E5_AGENCIA", "E5_CONTA"})

//_oConMan:AddIndex("1", {"E5_DTDISPO", "E5_BANCO", "E5_AGENCIA", "E5_CONTA"})

_oConMan:Create()

(cAliMov)->(dbGoTop())

Do While (cAliMov)->(!Eof())
	
	RecLock("cArqMov",.T.)

	cArqMov->E5_DATA     := (cAliMov)->E5_DATA
	cArqMov->E5_BENEF    := (cAliMov)->E5_BENEF
	cArqMov->E5_BANCO    :=  xBusCta((cAliMov)->E5_FILIAL, (cAliMov)->E5_BANCO,(cAliMov)->E5_AGENCIA, (cAliMov)->E5_CONTA)
	cArqMov->E5_CONTA    := (cAliMov)->E5_CONTA
	cArqMov->E5_AGENCIA  := (cAliMov)->E5_AGENCIA
	cArqMov->E5_HISTOR   := (cAliMov)->E5_HISTOR 
	cArqMov->E5_NATUREZ  := xBusNat((cAliMov)->E5_NATUREZ,(cAliMov)->E5_FILIAL)
	cArqMov->E5_FILIAL   := xBusFil((cAliMov)->E5_FILIAL)
	cArqMov->E5_VALOR    := (cAliMov)->E5_VALOR
	cArqMov->E5_RECPAG   := (cAliMov)->E5_RECPAG
	cArqMov->E5_SALDO    := xBusSald((cAliMov)->E5_RECPAG, (cAliMov)->E5_VALOR )
	cArqMov->E5RECNO     := (cAliMov)->R_E_C_N_O_
	
	MsUnLock()
	
	(cAliMov)->(DbSkip())
		
EndDo

DbGoTop() 


aCpoBro     := {{ "E5_DATA"      ,, "Data"             ,PesqPict("SE5","E5_DATA")},;                
               {  "E5_BENEF"     ,, "Favorecido"       ,PesqPict("SE5","E5_BENEF")},; 
               {  "E5_BANCO"     ,, "Banco"            ,PesqPict("SE5","E5_BANCO")},;             
               {  "E5_CONTA"     ,, "Conta"            ,PesqPict("SE5","E5_CONTA")},;
               {  "E5_AGENCIA"   ,, "Agencia"          ,PesqPict("SE5","E5_AGENCIA")},;
               {  "E5_HISTOR"    ,, "Memo"             ,PesqPict("SE5","E5_HISTOR")},;
               {  "E5_NATUREZ"   ,, "Categoria"        ,PesqPict("SED","ED_DESCRIC")},;
               {  "E5_FILIAL"    ,, "Empresa"          ,PesqPict("SE5","E5_FILIAL")},; 
               {  "E5_VALOR"     ,, "Montante"         ,PesqPict("SE5","E5_VALOR")}}
               
For nX := 1 To Len(aStru)
	If	aStru[nX][1] $ "E5_FILIAL|E5_DATA|E5_BENEF|E5_BANCO|E5_CONTA|E5_HISTOR|E5_NATUREZ|E5_VALOR"
		AAdd(aColumns,FWBrwColumn():New())
		aColumns[Len(aColumns)]:SetData( &("{||"+aStru[nX][1]+"}") )
		aColumns[Len(aColumns)]:SetTitle(RetTitle(aStru[nX][1])) 
		If	aStru[nX][1] $ "E5_HISTOR"
			aColumns[Len(aColumns)]:SetSize(60)
		Else
			aColumns[Len(aColumns)]:SetSize(aStru[nX][3])
		EndIf
		aColumns[Len(aColumns)]:SetDecimal(aStru[nX][4])
		aColumns[Len(aColumns)]:SetPicture(PesqPict("SE5",aStru[nX][1]))
	EndIf
Next nX

oBrowse := FWmBrowse():New()
oBrowse:SetAlias( "cArqMov" )
oBrowse:SetDescription( "Extrato Personalizado - SKYLINE" )
oBrowse:AddLegend("E5_RECPAG == 'R'","GREEN" 	,"Receber")
oBrowse:AddLegend("E5_RECPAG == 'P'","RED"  	,"A Pagar")
oBrowse:SetColumns(aColumns)

oBrowse:Activate()

If _oConMan <> Nil
	_oConMan:Delete() 
	_oConMan := Nil
EndIf
 

Return 

//-------------------------------------------------------------------
/*/{Protheus.doc} xBusNat
Busca descrição da natureza

@author André Brito
@since 26/06/2019
@version P12
/*/
//-------------------------------------------------------------------

Static Function xBusNat(cNat, cFil)

Local aArea   := GetArea()
Local cNatPai := ""
Local cNatTot := ""
Local cEspaco := ": "
Default cNat  := ""
Default cFil  := ""

DbSelectArea("SED")
SED->(dbSetOrder(1))	// ED_FILIAL+ED_CODIGO
SED->(dbSeek(xFilial("SED") + cNat))

cNat := SED->ED_DESCRIC
cNat := Alltrim(cNat)
cNatPai := SED->ED_PAI
 
If SED->(dbSeek(xFilial("SED") + cNatPai))
	cNatPai := Alltrim(SED->ED_DESCRIC)
EndIf

cNatTot := cNatPai + cEspaco + cNat

RestArea(aArea)

Return cNatTot

//-------------------------------------------------------------------
/*/{Protheus.doc} xBusCta
Busca descrição da natureza

@author André Brito
@since 26/06/2019
@version P12
/*/
//-------------------------------------------------------------------

Static Function xBusCta(cFil, cBanco, cAgencia, cConta)

Local aArea   := GetArea()
Local cEspaco := ""

Default cBanco   := ""
Default cAgencia := ""
Default cConta   := ""
Default cFil     := ""

DbSelectArea( "SA6" )
SA6->(dbSetOrder(1))
If DbSeek(cFil + cBanco + cAgencia + cConta)
	cBanco := Alltrim(SA6->A6_NOME) + " Cod.: " + Alltrim(SA6->A6_COD) + " Ag.: " + Alltrim(SA6->A6_AGENCIA) + " Num. Cta.: " + Alltrim(SA6->A6_NUMCON)
Else
	cBanco := ""
EndIf

RestArea(aArea)

Return cBanco


//-------------------------------------------------------------------
/*/{Protheus.doc} xBusFil
Busca descrição da Filial

@author André Brito
@since 26/06/2019
@version P12
/*/
//-------------------------------------------------------------------

Static Function xBusFil(cFil)

Local aArea   := GetArea()
Local cEspaco := ""
Local cEmp    := ""
Local cXFil   := ""
Local cNome   := ""
Local aRetSM0 := {}
Local nCont   := 0

aRetSM0	:= FWLoadSM0()

cEmp  := SubStr( cFil, 1, 2)//Codidgo da empresa 
cXFil := SubStr( cFil, 3, 4)//Codidgo da filial 

for nCont := 1 To Len(aRetSM0)
	If Alltrim(aRetSM0[nCont][2]) == cEmp + cXFil
		cNome := Alltrim(aRetSM0[nCont][7])
		Exit
	EndIf
Next
//cNome := Posicione("SM0",1,cEmp + cXFil,"M0_NOME") 

RestArea(aArea)

Return cNome

//-------------------------------------------------------------------
/*/{Protheus.doc} MenuDef
MenuDef

@author André Brito
@since 26/06/2019
@version P12
/*/
//-------------------------------------------------------------------

Static Function MenuDef()

Local aArea		:= GetArea()
Local aRotina 	:= {}

AADD(aRotina, {"Gerar Relatório"    , "U_xGerRel()"		, 0, 3, 0, Nil })
AADD(aRotina, {"Alterar"			, "U_EXEM992I"		, 0, 4, 0, Nil })
AADD(aRotina, {"Excluir"			, "U_EXEM992I"		, 0, 5, 3, Nil })
	
Return( aRotina )



User Function xGerRel()

MsgRun("Gerando extrato para conferência","Aguarde...",{|| xProcSky()})

Return

Static Function xProcSky()

Local aArea     := GetArea()
Local cArquivo  := "CONCI.XLS"
Local oExcelApp := Nil
Local cPath     := "C:\Conciliados"
Local aColunas  := {}
Local aLocais   := {}
Local cAba      := ""
Local cTabela   := {}
Local nValor    := 0
Local oExcel
Local oExcelApp
Local _oPlan

oBrush1  := TBrush():New(, RGB(193,205,205))

/*If !ApOleClient('MsExcel')

    MsgAlert("Falha ao abrir Excel!")
    Return

EndIf*/

oExcel  := FWMSExcel():New()
cAba    := "Extrato Bancário"
cTabela := "Saldo Somados dos Bancos Selecionados: " + cTotSal 


// Criação de nova aba 
oExcel:AddworkSheet(cAba)

// Criação de tabela
oExcel:AddTable (cAba,cTabela)

// Criação de colunas 
oExcel:AddColumn(cAba,cTabela,"DATA"        ,1,2,.F.) 
oExcel:AddColumn(cAba,cTabela,"FAVORECIDO"  ,3,2,.F.) 
oExcel:AddColumn(cAba,cTabela,"BANCO"  		,2,1,.F.) 
oExcel:AddColumn(cAba,cTabela,"MEMO"     	,3,2,.F.) 
oExcel:AddColumn(cAba,cTabela,"CATEGORIA"   ,3,2,.F.) 
oExcel:AddColumn(cAba,cTabela,"EMPRESA"     ,1,2,.F.) 
oExcel:AddColumn(cAba,cTabela,"MONTANTE"    ,2,1,.F.) 
oExcel:AddColumn(cAba,cTabela,"SALDO"       ,2,1,.F.) 

DbSelectArea("cArqMov")
DbGoTop() 

While !cArqMov->(Eof())
 	
    nValor := cArqMov->E5_VALOR
    If cArqMov->E5_RECPAG = 'P'
    	nValor := -1 * nValor
    Else
    	nValor := cArqMov->E5_VALOR
    EndIf
    
    oExcel:AddRow(cAba,cTabela, { cArqMov->E5_DATA,;
                                  cArqMov->E5_BENEF,; 
                                  cArqMov->E5_BANCO,; 
                                  cArqMov->E5_HISTOR,;
                                  cArqMov->E5_NATUREZ,;
                                  cArqMov->E5_FILIAL,;
                                  nValor,;
                                  cArqMov->E5_SALDO})

	cArqMov->(dbSkip())

End

If !Empty(oExcel:aWorkSheet)

    oExcel:Activate()
    oExcel:GetXMLFile(cArquivo)
 
    CpyS2T("\SYSTEM\"+cArquivo, cPath)

    oExcelApp := MsExcel():New()
    oExcelApp:WorkBooks:Open(cPath + "\" + cArquivo) // Abre a planilha
	oExcelApp:SetVisible(.T.)
	
EndIf

RestArea(aArea)

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} xBusPre
Busca Prefixos de Títulos

@author André Brito
@since 26/06/2019
@version P12
/*/
//-------------------------------------------------------------------

Static Function xBusPre()

Local aCampos   := {}
Local cQuery    := ""
Local cTabela   := '99'
Local lInverte  := .F.
Local cArqTrb   := GetNextAlias()
Local cAliAux   := GetNextAlias()
Local lCheck    := .F.
Local _oConPre

AADD(aCampos,{"X5_XOK"         ,"C",2,0})
AADD(aCampos,{"X5_FILIAL"      ,"C",TamSX3("X5_FILIAL")[1],0})
AADD(aCampos,{"X5_TABELA"      ,"C",TamSX3("X5_TABELA")[1],0})
AADD(aCampos,{"X5_CHAVE"       ,"C",TamSX3("X5_CHAVE")[1],0})
AADD(aCampos,{"X5_DESCRI"      ,"C",TamSX3("X5_DESCRI")[1],0})
AADD(aCampos,{"X5_RECNO"       ,"N",20  ,0})

cQuery := "SELECT R_E_C_N_O_, * FROM "
cQuery += RetSqlName("SX5") + " SX5 "
cQuery += " WHERE "
cQuery += " X5_TABELA = '" + cTabela + "' "  
cQuery += " AND D_E_L_E_T_ = ' ' "

cQuery := ChangeQuery(cQuery) 
 
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliAux,.T.,.T.)

If _oConPre <> Nil
	_oConPre:Delete() 
	_oConPre := Nil
EndIf

// Criando o objeto do arquivo temporário
_oConPre := FwTemporaryTable():New("cArqTrb")

// Criando a estrutura do objeto  
_oConPre:SetFields(aCampos)

// Criando o indice da tabela
_oConPre:AddIndex("1",{"X5_CHAVE"})

_oConPre:Create()

(cAliAux)->(dbGoTop())

Do While (cAliAux)->(!Eof())
	
	RecLock("cArqTrb",.T.) 
	
	cArqTrb->X5_XOK       := cMark
	cArqTrb->X5_FILIAL    := (cAliAux)->X5_FILIAL
	cArqTrb->X5_TABELA    := (cAliAux)->X5_TABELA
	cArqTrb->X5_CHAVE     := (cAliAux)->X5_CHAVE
	cArqTrb->X5_DESCRI    := (cAliAux)->X5_DESCRI
	cArqTrb->X5_RECNO     := (cAliAux)->R_E_C_N_O_
	
	MsUnLock()
	
	(cAliAux)->(DbSkip())
		
EndDo

DbGoTop()

aCpoBro     := {{ "X5_XOK"      ,, "Marcacao"        ,"@!"},;                
               {  "X5_FILIAL"   ,, "Filial"          ,PesqPict("SX5","X5_FILIAL")},;              
               {  "X5_TABELA"   ,, "Tabela"          ,PesqPict("SX5","X5_TABELA")},;
               {  "X5_CHAVE"    ,, "Chave"           ,PesqPict("SX5","X5_CHAVE")},;
               {  "X5_DESCRI"   ,, "Descricao"       ,PesqPict("SX5","X5_DESCRI")}}

aSize := MSADVSIZE()

DEFINE MSDIALOG oDlg TITLE "*** Seleção de Prefixos de Títulos - SkyLine ***" From 003,000 TO 340,417 PIXEL//From aSize[7],0 To aSize[6],aSize[5] OF oMainWnd PIXEL 

oPanel := TPanel():New(0,0,'',oDlg,, .T., .T.,, ,35,35,.T.,.T. )
oPanel:Align := CONTROL_ALIGN_TOP


@15,10 CHECKBOX oChk VAR lCheck PROMPT "Selecionar Todos" SIZE 60,007 PIXEL OF oPanel ON CLICK XConSx5(lCheck) 

@12,130 button "Confirmar" size 45,11 pixel of oPanel action {||xPrefix(),If(lXConti,oDlg:end(),lXConti := .F.)}

aCores := {} 

oMark := MsSelect():New("cArqTrb","X5_XOK","",aCpoBro,@lInverte,@cMark,{35,1,155,205},,,,,aCores) 
oMark:bMark := {| | xDisSx5(cMark)} 

ACTIVATE MSDIALOG oDlg CENTERED

Return cPrefixos

Static Function xPrefix()

Local cPrefix     := ""

DbSelectArea("cArqTrb") 
DbGotop()

Do While ("cArqTrb")->(!Eof()) 
	
	If !Empty(cArqTrb->X5_XOK)
		
		If Empty(cPrefix)	  
			cPrefix += "'" + Alltrim(cArqTrb->X5_CHAVE) + "'"
			lXConti := .T.
		Else
			cPrefix += "," + "'"+ Alltrim(cArqTrb->X5_CHAVE) + "'"
		EndIf
		
	EndIf
	
	("cArqTrb")->(DbSkip())
		
EndDo

("cArqTrb")->(dbCloseArea())

cPrefixos := cPrefix

Return .T.


//-------------------------------------------------------------------
/*/{Protheus.doc} xBusPre
Busca Prefixos de Títulos

@author André Brito
@since 26/06/2019
@version P12
/*/
//-------------------------------------------------------------------

Static Function xBusBan()

Local aCampos   := {}
Local cQuery    := ""
Local lInverte  := .F.
Local cArqTrb   := GetNextAlias()
Local cAliAux   := GetNextAlias()
Local lCheck    := .F.
Local _oConBan

AADD(aCampos,{"A6_XOK"         ,"C",2,0})
AADD(aCampos,{"A6_FILIAL"      ,"C",TamSX3("A6_FILIAL")[1],0})
AADD(aCampos,{"A6_COD"         ,"C",TamSX3("A6_COD")[1],0})
AADD(aCampos,{"A6_AGENCIA"     ,"C",TamSX3("A6_AGENCIA")[1],0})
AADD(aCampos,{"A6_NUMCON"      ,"C",TamSX3("A6_NUMCON")[1],0})
AADD(aCampos,{"A6_NOME"        ,"C",TamSX3("A6_NOME")[1],0})
AADD(aCampos,{"A6_SALATU"      ,"N",TamSX3("A6_SALATU")[1],2})
AADD(aCampos,{"A6_RECNO"       ,"N",20  ,0})

cQuery := "SELECT R_E_C_N_O_, * FROM "
cQuery += RetSqlName("SA6") + " SA6 "
cQuery += " WHERE "  
cQuery += " A6_FILIAL " + cRng + " "
cQuery += " AND D_E_L_E_T_ = ' ' ORDER BY A6_FILIAL"

cQuery := ChangeQuery(cQuery) 
 
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliAux,.T.,.T.)

If _oConBan <> Nil
	_oConBan:Delete() 
	_oConBan := Nil
EndIf

// Criando o objeto do arquivo temporário
_oConBan := FwTemporaryTable():New("cArqTrb")

// Criando a estrutura do objeto  
_oConBan:SetFields(aCampos)

// Criando o indice da tabela
_oConBan:AddIndex("1",{"A6_FILIAL", "A6_COD"})

_oConBan:Create()

(cAliAux)->(dbGoTop())

Do While (cAliAux)->(!Eof())
	
	RecLock("cArqTrb",.T.) 
	
	cArqTrb->A6_XOK       := cMark
	cArqTrb->A6_FILIAL    := (cAliAux)->A6_FILIAL
	cArqTrb->A6_COD       := (cAliAux)->A6_COD 
	cArqTrb->A6_AGENCIA   := (cAliAux)->A6_AGENCIA
	cArqTrb->A6_NUMCON    := (cAliAux)->A6_NUMCON
	cArqTrb->A6_NOME      := (cAliAux)->A6_NOME
	cArqTrb->A6_SALATU    := (cAliAux)->A6_SALATU
	cArqTrb->A6_RECNO     := (cAliAux)->R_E_C_N_O_
	
	MsUnLock()
	
	(cAliAux)->(DbSkip())
		
EndDo

DbGoTop()

aCpoBro     := {{ "A6_XOK"       ,, "Marcacao"       ,"@!"},;                
               {  "A6_FILIAL"    ,, "Filial"         ,PesqPict("SA6","A6_FILIAL")},;              
               {  "A6_COD"       ,, "Código"         ,PesqPict("SA6","A6_COD")},;
               {  "A6_AGENCIA"   ,, "Agencia"        ,PesqPict("SA6","A6_AGENCIA")},;
               {  "A6_NUMCON"    ,, "Numero Conta"   ,PesqPict("SA6","A6_NUMCON")},;
               {  "A6_NOME"      ,, "Nome"           ,PesqPict("SA6","A6_NOME")}}

aSize := MSADVSIZE()

DEFINE MSDIALOG oDlg TITLE "*** Seleção de Bancos - SkyLine ***" From 003,000 TO 400,477 PIXEL//From aSize[7],0 To aSize[6],aSize[5] OF oMainWnd PIXEL 

oPanel := TPanel():New(0,0,'',oDlg,, .T., .T.,, ,35,35,.T.,.T. )
oPanel:Align := CONTROL_ALIGN_TOP


@15,10 CHECKBOX oChk VAR lCheck PROMPT "Selecionar Todos" SIZE 60,007 PIXEL OF oPanel ON CLICK XConSa6(lCheck) 

@12,130 button "Confirmar" size 45,11 pixel of oPanel action {||xBancos(),If(lXConti,oDlg:end(),lXConti := .F.)}
aCores := {} 

oMark := MsSelect():New("cArqTrb","A6_XOK","",aCpoBro,@lInverte,@cMark,{35,1,175,230},,,,,aCores) 
oMark:bMark := {| | xDisSa6(cMark)} 

ACTIVATE MSDIALOG oDlg CENTERED

Return cBancos

Static Function xBancos()

Local cBanco       := ""
Local cAgencia     := ""
Local cConta       := ""

DbSelectArea("cArqTrb") 
DbGotop()

Do While ("cArqTrb")->(!Eof()) 
	
	If !Empty(cArqTrb->A6_XOK)
		
		If Empty(cBanco)	  
			cBanco += "'" + Alltrim(cArqTrb->A6_COD) + "'"
			lXConti := .T.
		Else
			cBanco += "," + "'"+ Alltrim(cArqTrb->A6_COD) + "'"
		EndIf
		
		If Empty(cAgencia)	  
			cAgencia += "'" + Alltrim(cArqTrb->A6_AGENCIA) + "'"
			lXConti := .T.
		Else
			cAgencia += "," + "'"+ Alltrim(cArqTrb->A6_AGENCIA) + "'"
		EndIf
		
		If Empty(cConta)	  
			cConta += "'" + Alltrim(cArqTrb->A6_NUMCON) + "'"
			lXConti := .T.
		Else
			cConta += "," + "'"+ Alltrim(cArqTrb->A6_NUMCON) + "'"
		EndIf
		
		nTotSal += cArqTrb->A6_SALATU
		
	EndIf
	
	("cArqTrb")->(DbSkip())
		
EndDo

("cArqTrb")->(dbCloseArea())

cBancos := cBanco
cAgens  := cAgencia
cCtas   := cConta

cTotSal := Transform(nTotSal, "@E 999,999,999,999.99" ) 

Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} XConSx5
Função que realiza a marcação/desmarcação de todos os registros da tabela temporária

@author André Brito
@since 14/06/2019
@version P12
/*/
//-------------------------------------------------------------------

Static Function  XConSx5(lCheck)

Local aArea := GetArea()

dbSelectArea( "cArqTrb" ) 
dbGotop() 

Do While !EoF()
 
    If lCheck
    
		If RecLock( "cArqTrb", .F. ) 
			
			If Empty(cArqTrb->X5_XOK)
			
				cArqTrb->X5_XOK  := cMark 
			
			EndIf
			
			MsUnLock() 
		
		EndIf 
	Else
	
		If RecLock( "cArqTrb", .F. ) 
			
			If !Empty(cArqTrb->X5_XOK)
			
				cArqTrb->X5_XOK  := ''
				
			EndIf 
			
			MsUnLock() 
		
		EndIf 

	EndIf
	
	dbSkip() 

EndDo 

RestArea(aArea)

oMark:oBrowse:Refresh() 

Return 

//-------------------------------------------------------------------
/*/{Protheus.doc} xDisAvc
Funcao executada ao Marcar/Desmarcar um registro.  

@author André Brito
@since 14/06/2019
@version P12
/*/
//-------------------------------------------------------------------


Static Function xDisSx5(cMark) 

Local cMarca := cMark
Local aArea  := GetArea()

RecLock("cArqTrb",.F.) 

If Marked("X5_XOK")    

	cArqTrb->X5_XOK := cMarca  
	
Else 

	cArqTrb->X5_XOK := "" 

EndIf

MsUnLock() 

RestArea(aArea)

oMark:oBrowse:Refresh() 

Return

//-------------------------------------------------------------------
/*/{Protheus.doc} XConSa6
Função que realiza a marcação/desmarcação de todos os registros da tabela temporária

@author André Brito
@since 14/06/2019
@version P12
/*/
//-------------------------------------------------------------------

Static Function  XConSa6(lCheck)

Local aArea := GetArea()

dbSelectArea( "cArqTrb" ) 
dbGotop() 

Do While !EoF()
 
    If lCheck
    
		If RecLock( "cArqTrb", .F. ) 
			
			If Empty(cArqTrb->A6_XOK)
			
				cArqTrb->A6_XOK  := cMark 
			
			EndIf
			
			MsUnLock() 
		
		EndIf 
	Else
	
		If RecLock( "cArqTrb", .F. ) 
			
			If !Empty(cArqTrb->A6_XOK)
			
				cArqTrb->A6_XOK  := ''
				
			EndIf 
			
			MsUnLock() 
		
		EndIf 

	EndIf
	
	dbSkip() 

EndDo 

RestArea(aArea)

oMark:oBrowse:Refresh() 

Return 

//-------------------------------------------------------------------
/*/{Protheus.doc} xDisSa6
Funcao executada ao Marcar/Desmarcar um registro.  

@author André Brito
@since 14/06/2019
@version P12
/*/
//-------------------------------------------------------------------


Static Function xDisSa6(cMark) 

Local cMarca := cMark
Local aArea  := GetArea()

RecLock("cArqTrb",.F.) 

If Marked("A6_XOK")    

	cArqTrb->A6_XOK := cMarca  
	
Else 

	cArqTrb->A6_XOK := "" 

EndIf

MsUnLock() 

RestArea(aArea)

oMark:oBrowse:Refresh() 

Return


//-------------------------------------------------------------------
/*/{Protheus.doc} xBusSald
Funcao executada ao Marcar/Desmarcar um registro.  

@author André Brito
@since 14/06/2019
@version P12
/*/
//-------------------------------------------------------------------


Static Function xBusSald(cRecPag, nValor) 

Local aArea  := GetArea()
Local nSaldo := 0
Local cSaldo := ""

/*DbSelectArea( "SA6" )
SA6->(dbSetOrder(1))
If DbSeek(cFil + cBanco + cAgencia + cConta)
	nSaldo := SA6->A6_SALATU
Else
	nSaldo := 0
EndIf*/

If cRecPag == 'R'
	nXValTot += nValor
ElseIf cRecPag == 'P'
	nXValTot -= nValor
EndIf

cSaldo := Transform(nXValTot, "@E 999,999,999,999.99" ) 

RestArea(aArea)

Return cSaldo

