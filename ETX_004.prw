#INCLUDE "PROTHEUS.CH"        
#INCLUDE "TOPCONN.CH" 
#include "rwmake.ch"  
#include "fileio.ch"    
#INCLUDE "FWPrintSetup.ch"
#INCLUDE "RPTDEF.CH"
#Include "DBTREE.CH"
#Include "HBUTTON.CH"
#Define XENTERX Chr(13)+Chr(10) 
//============================================================================================================================= 
User Function ETX_004()                        
//=============================================================================================================================
Local oButton1
Local oButton2
Local oButton3
Local oButton4
Local oGroup1                
Local nLinx := 280
Static  oDlgExt
Private aSize   := MSADVSIZE()
Private aRotina := Menudef() 
Private cTabela14 := ""
Private cPerg   := "ETX004"
Private nLin1        := 001  
Private nCol1        := 002
Private nAlt1        := 270
Private nComp1       := 680
Private nPosSair     := 630
Private aHeader04    := {}
Private aCols04      := {}


SX5->(DbSetOrder(1))
SX5->(MsSeek(xFilial("SX5")+"14"))
While SX5->(!Eof()) .And. SX5->X5_TABELA == "14"
	cTabela14 += (Alltrim(SX5->X5_CHAVE) + "/")
	SX5->(DbSkip())
End	
cTabela14 += If(cPaisLoc=="BRA","","/$ ")         

If !Pergunte(cPerg)
   Return
Endif   

DEFINE MSDIALOG oDlgExt TITLE "* Movimentações de "+DTOC(MV_PAR03)+" à "+DTOC(MV_PAR04) From aSize[7],0 To aSize[6],aSize[5] OF oMainWnd PIXEL
       fMSNewGe1()          
       @ nLinx-6, 002 GROUP oGroup1 TO nLinx-17, nComp1-10 OF oDlgExt COLOR 0, 16777215 PIXEL
       @ nLinx, 007      BUTTON oButton1 PROMPT "Relatório" SIZE 037, 012 OF oDlgExt ACTION fRelat(aHeader04,aCols04) PIXEL
       @ nLinx, 049      BUTTON oButton2 PROMPT "Excel"     SIZE 037, 012 OF oDlgExt ACTION  fExcel(aHeader04,aCols04) PIXEL
       //@ nLinx, 091      BUTTON oButton3 PROMPT "HTML"      SIZE 037, 012 OF oDlgExt ACTION fHtml() PIXEL        
       @ nLinx, nPosSair BUTTON oButton4 PROMPT "Sair"      SIZE 037, 012 OF oDlgExt ACTION oDlgExt:End() PIXEL  
ACTIVATE MSDIALOG oDlgExt CENTERED
Return

//=============================================================================================================================
Static Function fMSNewGe1()
//=============================================================================================================================
Local nX           := 0
Local aFieldFill   := {}
Local aAlterFields := {}  
Local aDados       := {}
Local aItens       := {}
Local aDadosAux    := {}
Local nSaldo       := 0
Local dData        := dDataBase
Local nPosDat      := 0
Local lAchou       := .F.
Local nSaldoAtu    := 0
Local nSaldoIni    := 0
Local nDebito      := 0
Local nCredito    := 0

Local cNomeRem     := ""
Local cIERem       := ""
Local cEndRem      := ""
Local cMunRem      := ""
Local cUFRem       := ""
Local cCNPJRem     := ""     
//Local aFields      := {"E2_FILORIG","A3_NOME"    ,"E2_VENCREA","E2_NUM","E2_FORNECE"  ,"E2_LOJA","A2_NOME"   ,"E2_TIPO","A1_NREDUZ","E2_NATUREZ","ED_DESCRIC","E2_HIST"  ,"E1_VALOR","E2_VALOR","D1_TOTAL"}
Local aFields      := {"E2_FILORIG","A3_NOME"    ,"E2_VENCREA","E2_NUM","A2_NOME"   ,"A3_NREDUZ","ED_DESCRIC","E2_HIST"  ,"E1_VALOR","E2_VALOR","D1_TOTAL"}
Local aTitulo      := {"Filial"    ,"Nome Filial","Vencimento","Número","Favorecido","Tp Titulo","Natureza Financeira" ,"Histórico","Crédito" ,"Débito"  ,"Total"}
Local aSaldos      := {}
Local cxFilial     := "ZZ"
Static oMSNewGe1  

aHeader04    := {}
aCols04      := {}

// 005, 009, 202, 451             

DbSelectArea("SX3")
SX3->(DbSetOrder(2))
For nX := 1 to Len(aFields)
    If SX3->(DbSeek(aFields[nX]))
      aAdd(aHeader04,{aTitulo[nX],;
                      SX3->X3_CAMPO,;
                      If(SX3->X3_TIPO="N","@E 9,999,999.99",SX3->X3_PICTURE),;
                      If(ALLTRIM(SX3->X3_TIPO)="D",10,If(ALLTRIM(SX3->X3_TIPO)="N".AND.LEFT(SX3->X3_PICTURE,2)="@E",12,SX3->X3_TAMANHO)),;
                      SX3->X3_DECIMAL,;
                      SX3->X3_VALID,;
                      SX3->X3_USADO,;
                      SX3->X3_TIPO,;
                      SX3->X3_F3,;
                      SX3->X3_CONTEXT,;
                      SX3->X3_CBOX,;
                      SX3->X3_RELACAO})
    Endif
Next nX
nSaldoAtu := 0
If select("RS_TT") > 0
   dbselectarea("RS_TT") 
   dbCloseArea("RS_TT") 
Endif 
cQuery := ""
If MV_PAR05 = 1 .OR. MV_PAR05 = 3
   cQuery := " SELECT 'P' E2_TIPOX,E2_FILORIG,' ' E3_NOME,E2_VENCREA,E2_NUM,E2_FORNECE,E2_LOJA,A2_NOME,E2_TIPO,LEFT(X5_DESCRI,20) A3_NREDUZ,E2_NATUREZ,ED_DESCRIC,E2_HIST,E2_VALOR"  
   cQuery += " FROM "+RetSqlName("SE2")+" E2" 
   cQuery += " LEFT JOIN "+RetSqlName("SA2")+" AS A2 ON A2.D_E_L_E_T_ = ' ' AND A2_COD = E2_FORNECE AND A2_LOJA = E2_LOJA"
   cQuery += " LEFT JOIN "+RetSqlName("SED")+" AS ED ON ED.D_E_L_E_T_ = ' ' AND ED_CODIGO = E2_NATUREZ"
   cQuery += " LEFT JOIN "+RetSqlName("SX5")+" AS X5 ON X5.D_E_L_E_T_ = ' ' AND X5_TABELA = '05' AND X5_CHAVE = E2_TIPO"
   cQuery += " WHERE E2.D_E_L_E_T_ = '' "
   cQuery += " AND E2_BAIXA = '' "
   cQuery += " AND E2_FILORIG BETWEEN '"+MV_PAR01+"' AND '"+MV_PAR02+"'"
   cQuery += " AND E2_VENCREA BETWEEN '"+DTOS(MV_PAR03)+"' AND '"+DTOS(MV_PAR04)+"'"  
Endif
If MV_PAR05 = 1 
   cQuery += " UNION ALL"
Endif
If MV_PAR05 = 1 .OR. MV_PAR05 = 2   
   cQuery += " SELECT 'R' E2_TIPOX, E1_FILORIG E2_FILORIG,' ' E3_NOME,E1_VENCREA E2_VENCREA,E1_NUM E2_NUM,E1_CLIENTE E2_FORNECE,E1_LOJA E2_LOJA,A1_NOME A2_NOME,E1_TIPO E2_TIPO,LEFT(X5_DESCRI,20) A3_NREDUZ ,E1_NATUREZ E2_NATUREZ,ED_DESCRIC,E1_HIST E2_HIST,E1_VALOR E2_VALOR "  
   cQuery += " FROM "+RetSqlName("SE1")+" E1" 
   cQuery += " LEFT JOIN "+RetSqlName("SA1")+" AS A1 ON A1.D_E_L_E_T_ = ' ' AND A1_COD = E1_CLIENTE AND A1_LOJA = E1_LOJA "
   cQuery += " LEFT JOIN "+RetSqlName("SED")+" AS ED ON ED.D_E_L_E_T_ = ' ' AND ED_CODIGO = E1_NATUREZ"
   cQuery += " LEFT JOIN "+RetSqlName("SX5")+" AS X5 ON X5.D_E_L_E_T_ = ' ' AND X5_TABELA = '05' AND X5_CHAVE = E1_TIPO"
   cQuery += " WHERE E1.D_E_L_E_T_ = '' "
   cQuery += " AND E1_BAIXA = '' "
   cQuery += " AND E1_FILORIG BETWEEN '"+MV_PAR01+"' AND '"+MV_PAR02+"'"
   cQuery += " AND E1_VENCREA BETWEEN '"+DTOS(MV_PAR03)+"' AND '"+DTOS(MV_PAR04)+"'"
Endif
If MV_PAR05 = 1 .OR. MV_PAR05 = 3
   cQuery += " ORDER BY E2_VENCREA,E2_FILORIG,E2_NUM"
Endif
If MV_PAR05 = 2   
   cQuery += " ORDER BY E1_VENCREA,E1_FILORIG,E1_NUM"
Endif
dbUseArea(.T., "TOPCONN", TCGenQry(,,cQuery), "RS_TT", .F., .T.) 
dbselectarea("RS_TT")  
DbGoTop()  
While RS_TT->(!Eof())  
      dbSelectArea("SM0")
      dbSetOrder(1)
      dbGoTop()
      If dbSeek(cEmpAnt+RS_TT->E2_FILORIG)
         cNomeRem := LEFT(SM0->M0_FILIAL,20)
         cCNPJRem := STRZERO(VAL(SM0->M0_CGC),14)
         cIERem   := SM0->M0_INSC
         cEndRem  := SM0->M0_ENDENT
         cMunRem  := SM0->M0_CIDENT
         cUFRem   := SM0->M0_ESTENT    
         cCNPJRem := SUBSTR(CCNPJREM,1,2)+"."+SUBSTR(CCNPJREM,3,3)+"."+SUBSTR(CCNPJREM,6,3)+"/"+SUBSTR(CCNPJREM,9,4)+"-"+SUBSTR(CCNPJREM,13,2)     
      Endif	    
      If RS_TT->E2_TIPOX = "P"
         nDebito += RS_TT->E2_VALOR  
         nSaldoAtu := nSaldoAtu - RS_TT->E2_VALOR              
         Else      
         nSaldoAtu := nSaldoAtu + RS_TT->E2_VALOR  
         nCredito  += RS_TT->E2_VALOR                           
      Endif         
      aAdd(aCols04, {})  
      For nX := 1 to LEN(aHeader04)-3   
          aAdd(aCols04[Len(aCols04)],If(ALLTRIM(aHeader04[nX,2])="A3_NOME",cNomeRem,If(ALLTRIM(aHeader04[nX,8])="D",STOD(RS_TT->&(aHeader04[nX,2])),RS_TT->&(aHeader04[nX,2]))))
      Next nX
      If RS_TT->E2_TIPOX = "P"
         aAdd(aCols04[Len(aCols04)],0)
         aAdd(aCols04[Len(aCols04)],RS_TT->E2_VALOR)   
      Else
         aAdd(aCols04[Len(aCols04)],RS_TT->E2_VALOR)     
         aAdd(aCols04[Len(aCols04)],0)        
      Endif      
      aAdd(aCols04[Len(aCols04)],nSaldoAtu)
      aAdd(aCols04[Len(aCols04)],.F.)                                                       
      RS_TT->(DbSkip())        		              
Enddo  
aAdd(aCols04, {})  
For nX := 1 to LEN(aHeader04)-3    
    aAdd(aCols04[Len(aCols04)],If(ALLTRIM(aHeader04[nX,8])="D",CTOD("  /  /  "),""))
Next nX
aAdd(aCols04[Len(aCols04)],nCredito)
aAdd(aCols04[Len(aCols04)],nDebito)
aAdd(aCols04[Len(aCols04)],nSaldoAtu)    
aCols04[LEN(aCols04),2] := "TOTAL GERAL"
aAdd(aCols04[Len(aCols04)],.F.)       
oMSNewGe1 := MsNewGetDados():New(nLin1,nCol1,nAlt1,nComp1,GD_INSERT+GD_DELETE+GD_UPDATE,"AllwaysTrue","AllwaysTrue","+Field1+Field2",aAlterFields,,999,"AllwaysTrue","", "AllwaysTrue",oDlgExt,aHeader04,aCols04)
Return

//=============================================================================================================================
Static Function fRelat(aHeader04,aCols04) 
//=============================================================================================================================
Local aTitulo := {}

aAdd(aTitulo,{"Movimentação de Contas a Receber / Pagar"})   
aAdd(aTitulo,{""})     
aAdd(aTitulo,{"REFERÊNCIA [ "+DTOC(MV_PAR03)+" a "+DTOC(MV_PAR04)+" ]"})   
nX := 1     

U_ETX_RELT(1,"ETX_004","Movimentação de Contas a Receber / Pagar","SE5",aTitulo,aHeader04,aCols04,nX)  
Return

//=============================================================================================================================
Static Function fExcel(aHeader04,aCols04)
//=============================================================================================================================
Local aCabec    := {}
Local aPlanilha := {}

aAdd(aCabec,aHeader04)        
aAdd(aPlanilha,{aCabec,aCols04,"Movimentações","Movimentações emitido em: ["+DTOC(dDataBase)+"] Período: ["+DTOC(MV_PAR03)+" / "+DTOC(MV_PAR04)+"]","1"})
ExecBlock("ETX_EXCEL",.F.,.F.,{Len(aHeader04),"Movimentações",aPlanilha})                       
Return 
//=============================================================================================================================
Static Function fFluxo() 
//=============================================================================================================================
ALERT("Fluxo")
Return 

//=============================================================================================================================
Static Function MenuDef()
//=============================================================================================================================
Local axRotina := {}

aAdd(axRotina,{"Pesquisa","AxPesqui"    ,0,1})
aAdd(axRotina,{"Visualizar","AxVisual"  ,0,2})
aAdd(axRotina,{"Relatório","U_XRELAT()",0,3})
aAdd(axRotina,{"Excel","U_xexcel()"  ,0,4})
aAdd(axRotina,{"Imprimir","U_PAGLF003()" ,0,7})
aAdd(axRotina,{"HTML","U_XHTML()" ,0,8})
Return axRotina