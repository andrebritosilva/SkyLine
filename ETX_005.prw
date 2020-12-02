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
//Programa............: ETX_005()
//Autor...............: Paulo César (PC) 
//Data................: 12/06/2019
//Descricao / Objetivo: Extrato Bancário  
//Cliente             : ETHOS X / SKYLINE
//============================================================================================================================= 
User Function ETX_005()                        
//=============================================================================================================================
Local oButton1
Local oButton2
Local oButton3
Local oButton4
Local oGroup1                
Local nLinx := 280
Static  oDlgExt
Private aSize     := MSADVSIZE()
Private aRotina   := Menudef() 
Private cTabela14 := ""
Private cPerg     := "ETX005"
Private nLin1     := 001  
Private nCol1     := 002
Private nAlt1     := 270
Private nComp1    := 680
Private nPosSair  := 630
Private aFields   := {"E5_RECONC","E5_RECPAG","E5_NUMERO","E5_DTDISPO","A2_NOME"   ,"A6_NOME","X5_DESCRI","E5_HISTOR","ED_DESCRIC","E5_FILORIG","E5_VALOR","E5_VALOR","E5_VALOR"}
Private aTitulo   := {"Conc"     ,"R/P"       ,"Número"    ,"Data"      ,"Favorecido","Conta"  ,"Tipo"     ,"Histórico","Categoria" ,"Empresa"   ,"Crédito" ,"Débito"  ,"Saldo Atual"}                                                                                             
Private aHeaderEx := {}   
Private aColsEx      := {}

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

DEFINE MSDIALOG oDlgExt TITLE "* Extrato Bancário" /*+DTOC(MV_PAR09)+" à "+DTOC(MV_PAR10)*/ From aSize[7],0 To aSize[6],aSize[5] OF oMainWnd PIXEL
       fMSNewGE5()          
       @ nLinx-6, 002 GROUP oGroup1 TO nLinx-17, nComp1-10 OF oDlgExt COLOR 0, 16777215 PIXEL
       @ nLinx, 007      BUTTON oButton1 PROMPT "Relatório" SIZE 037, 012 OF oDlgExt ACTION fRelat(aHeaderEx,aColsEX) PIXEL
       @ nLinx, 049      BUTTON oButton2 PROMPT "Excel"     SIZE 037, 012 OF oDlgExt ACTION fExcel(aHeaderEx,aColsEX) PIXEL
       @ nLinx, 091      BUTTON oButton3 PROMPT "Títulos abertos"      SIZE 050, 012 OF oDlgExt ACTION fFluxo() PIXEL          
       @ nLinx, nPosSair BUTTON oButton4 PROMPT "Sair"      SIZE 037, 012 OF oDlgExt ACTION oDlgExt:End() PIXEL  
ACTIVATE MSDIALOG oDlgExt CENTERED
Return

//=============================================================================================================================
Static Function fMSNewGE5()
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
Local cNomeRem     := ""
Local cIERem       := ""
Local cEndRem      := ""
Local cMunRem      := ""
Local cUFRem       := ""
Local cCNPJRem     := ""  
Local aSaldos      := {}
Local cxFilial     := "ZZ"
Static oMSNewGE5  

aHeaderEx    := {}

DbSelectArea("SX3")
SX3->(DbSetOrder(2))
For nX := 1 to Len(aFields)
    If SX3->(DbSeek(aFields[nX]))
      aAdd(aHeaderEx,{aTitulo[nX],;
                      SX3->X3_CAMPO,;
                      If(SX3->X3_TIPO="N","@E 9,999,999.99",SX3->X3_PICTURE),;
                      If(ALLTRIM(SX3->X3_CAMPO)="X5_DESCRI",20,If(ALLTRIM(SX3->X3_TIPO)="D",10,SX3->X3_TAMANHO)),;
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
If select("RS_TT") > 0
   dbselectarea("RS_TT") 
   dbCloseArea("RS_TT") 
Endif 
cQuery := " SELECT E5_FILORIG,E5_BANCO,A6_NOME,E5_AGENCIA,E5_CONTA "  
cQuery += " FROM "+RetSqlName("SE5")+" E5" 
cQuery += " LEFT JOIN "+RetSqlName("SA6")+" AS A6 ON A6.D_E_L_E_T_ = ' ' AND A6_FILIAL = E5_FILORIG AND A6_COD = E5_BANCO AND A6_AGENCIA = E5_AGENCIA AND A6_NUMCON = E5_CONTA"
cQuery += " WHERE E5.D_E_L_E_T_ = '' "
cQuery += " AND E5_BANCO <> '' "
cQuery += " AND E5_FILORIG BETWEEN '"+MV_PAR01+"' AND '"+MV_PAR02+"'"
cQuery += " AND LTRIM(RTRIM(E5_BANCO))+LTRIM(RTRIM(E5_AGENCIA))+LTRIM(RTRIM(E5_CONTA)) BETWEEN '"+ALLTRIM(MV_PAR03)+ALLTRIM(MV_PAR04)+ALLTRIM(MV_PAR05)+"' AND '"+ALLTRIM(MV_PAR06)+ALLTRIM(MV_PAR07)+ALLTRIM(MV_PAR08)+"'"
cQuery += " AND E5_DTDISPO BETWEEN '"+DTOS(MV_PAR06)+"' AND '"+DTOS(MV_PAR07)+"'"
cQuery += " GROUP BY E5_FILORIG,E5_BANCO,A6_NOME,E5_AGENCIA,E5_CONTA"
cQuery += " ORDER BY E5_FILORIG,E5_BANCO,A6_NOME,E5_AGENCIA,E5_CONTA"
dbUseArea(.T., "TOPCONN", TCGenQry(,,cQuery), "RS_TT", .F., .T.) 
dbselectarea("RS_TT")  
DbGoTop()  
While RS_TT->(!Eof())  
      nSaldoAtu := 0
      nSaldoIni := 0                                                                    
    dbSelectArea("SM0")
    dbSetOrder(1)
    dbGoTop()
    If dbSeek(cEmpAnt+RS_TT->E5_FILORIG)
       cNomeRem := SM0->M0_FILIAL
       cCNPJRem := STRZERO(VAL(SM0->M0_CGC),14)
       cIERem   := SM0->M0_INSC
       cEndRem  := SM0->M0_ENDENT
       cMunRem  := SM0->M0_CIDENT
       cUFRem   := SM0->M0_ESTENT    
       cCNPJRem := SUBSTR(CCNPJREM,1,2)+"."+SUBSTR(CCNPJREM,3,3)+"."+SUBSTR(CCNPJREM,6,3)+"/"+SUBSTR(CCNPJREM,9,4)+"-"+SUBSTR(CCNPJREM,13,2)     
    Endif	             
    aSaldos := fLerSE8(RS_TT->E5_FILORIG,RS_TT->E5_BANCO,RS_TT->E5_AGENCIA,RS_TT->E5_CONTA) 
    nSaldoIni := aSaldos[1]
    nSaldoAtu := aSaldos[2]    
    aItens := {}
    aItens := fLerSE5(RS_TT->E5_FILORIG,cNomeRem,RS_TT->E5_BANCO,RS_TT->A6_NOME,RS_TT->E5_AGENCIA,RS_TT->E5_CONTA,nSaldoIni,nSaldoAtu)
    aAdd(aDadosAux,aItens)    
    RS_TT->(DbSkip())        		              
Enddo  
aColsEx := {} 
For nX := 1 to  LEN(aDadosAux)  
    For nPosDat := 1 to LEN(aDadosAux[nX])
        aAdd(aColsEx,{})
        For nSaldo := 1 to LEN(aHeaderEx)     
            If LEN(aDadosAux[nX,nPosDat]) = LEN(aHeaderEx) //aHeaderEx[nSaldo,2]  
               aAdd(aColsEx[Len(aColsEx)],If(ALLTRIM(aHeaderEx[nSaldo,8])="D",STOD(aDadosAux[nX,nPosDat,nSaldo]),aDadosAux[nX,nPosDat,nSaldo]))      
            Else    
               ALERT("JJJ")
            Endif   
        Next nSaldo
        aAdd(aColsEx[Len(aColsEx)],.F.)                            
    Next nPosDat    
Next nX  
oMSNewGE5 := MsNewGetDados():New(nLin1,nCol1,nAlt1,nComp1,GD_INSERT+GD_DELETE+GD_UPDATE,"AllwaysTrue","AllwaysTrue","+Field1+Field2",aAlterFields,,999,"AllwaysTrue","", "AllwaysTrue",oDlgExt,aHeaderEx,aColsEx)
Return

//=============================================================================================================================
Static Function fLerSE5(pFilial,pNomFil,pBanco,pNomBco,pAgencia,pConta,pSaldoIni,pSaldoAtu)
//=============================================================================================================================
Local aArray   := {}
Local aLinha  := {}
Local nCredito := 0
Local nDebito  := 0    
Local nTCredito := 0
Local nTDebito  := 0 
Local nSaldoIni := pSaldoIni  
Local nSaldoAtu := pSaldoIni  
Local nX        := 0
For nX := 1 to LEN(aFields)-3       
    aAdd(aLinha,If(ALLTRIM(aHeaderEx[nX,8])="D",DTOS(MV_PAR09)," "))      
Next nX      
aLinha[5] := "SALDO INICIAL"
aLinha[6] := ALLTRIM(pBanco)+"-"+ALLTRIM(pAgencia)+"-"+ALLTRIM(pConta)+"-"+ALLTRIM(pNomBco)
aAdd(aLinha,0)
aAdd(aLinha,0)                          
aAdd(aLinha,nSaldoIni)      
aAdd(aArray,aLinha)

If select("RS_OK") > 0
   dbselectarea("RS_OK") 
   dbCloseArea("RS_OK") 
Endif 
cQuery := " SELECT E5_RECONC,E5_RECPAG,E5_TIPODOC,E5_NUMERO,E5_DTDISPO,A2_NOME,A6_NOME,X5_DESCRI,E5_HISTOR,ED_DESCRIC,E5_FILORIG,E5_VALOR"   
cQuery += " FROM "+RetSqlName("SE5")+" E5"
cQuery += " LEFT JOIN "+RetSqlName("SA6")+" A6 ON A6.D_E_L_E_T_ = ' ' AND A6_FILIAL = E5_FILORIG AND A6_COD = E5_BANCO AND A6_AGENCIA = E5_AGENCIA AND A6_NUMCON = E5_CONTA"
cQuery += " LEFT JOIN "+RetSqlName("SED")+" ED ON ED.D_E_L_E_T_ = ' ' AND ED_FILIAL = '"+XFILIAL("SED")+"' AND ED_CODIGO = E5_NATUREZ"
cQuery += " LEFT JOIN "+RetSqlName("SA2")+" A2 ON A2.D_E_L_E_T_ = ' ' AND A2_FILIAL = '"+XFILIAL("SA2")+"' AND A2_COD = E5_CLIFOR AND A2_LOJA = A2_LOJA"
cQuery += " LEFT JOIN "+RetSqlName("SX5")+" X5 ON X5.D_E_L_E_T_ = ' ' AND X5_FILIAL = '"+XFILIAL("SX5")+"' AND X5_TABELA = '14' AND X5_CHAVE = E5_MOEDA"
cQuery += " WHERE E5.D_E_L_E_T_ = '' "
cQuery += " AND E5_RECPAG = 'P' "
cQuery += " AND E5_FILORIG = '"+pFilial+"'"
cQuery += " AND E5_BANCO = '"+pBanco+"'"
cQuery += " AND E5_AGENCIA = '"+pAgencia+"'"
cQuery += " AND E5_CONTA = '"+pConta+"'"  
cQuery += " AND E5_DTDISPO BETWEEN '"+DTOS(MV_PAR06)+"' AND '"+DTOS(MV_PAR07)+"'"
cQuery += " UNION ALL"  
cQuery += " SELECT E5_RECONC,E5_RECPAG,E5_TIPODOC,E5_NUMERO,E5_DTDISPO,A1_NOME,A6_NOME,X5_DESCRI,E5_HISTOR,ED_DESCRIC,E5_FILORIG,E5_VALOR"   
cQuery += " FROM "+RetSqlName("SE5")+" E5"
cQuery += " LEFT JOIN "+RetSqlName("SA6")+" A6 ON A6.D_E_L_E_T_ = ' ' AND A6_FILIAL = E5_FILORIG AND A6_COD = E5_BANCO AND A6_AGENCIA = E5_AGENCIA AND A6_NUMCON = E5_CONTA"
cQuery += " LEFT JOIN "+RetSqlName("SED")+" ED ON ED.D_E_L_E_T_ = ' ' AND ED_FILIAL = '"+XFILIAL("SED")+"' AND ED_CODIGO = E5_NATUREZ"
cQuery += " LEFT JOIN "+RetSqlName("SA1")+" A1 ON A1.D_E_L_E_T_ = ' ' AND A1_FILIAL = '"+XFILIAL("SA1")+"' AND A1_COD = E5_CLIFOR AND A1_LOJA = E5_LOJA"
cQuery += " LEFT JOIN "+RetSqlName("SX5")+" X5 ON X5.D_E_L_E_T_ = ' ' AND X5_FILIAL = '"+XFILIAL("SX5")+"' AND X5_TABELA = '14' AND X5_CHAVE = E5_MOEDA"
cQuery += " WHERE E5.D_E_L_E_T_ = '' "
cQuery += " AND E5_RECPAG = 'R' "
cQuery += " AND E5_FILORIG = '"+pFilial+"'"
cQuery += " AND E5_BANCO = '"+pBanco+"'"
cQuery += " AND E5_AGENCIA = '"+pAgencia+"'"
cQuery += " AND E5_CONTA = '"+pConta+"'"  
cQuery += " AND E5_DTDISPO BETWEEN '"+DTOS(MV_PAR09)+"' AND '"+DTOS(MV_PAR10)+"'"
cQuery += " ORDER BY E5_DTDISPO,A2_NOME,A6_NOME"
dbUseArea(.T., "TOPCONN", TCGenQry(,,cQuery), "RS_OK", .F., .T.) 
dbselectarea("RS_OK")  
DbGoTop()  
While RS_OK->(!Eof())  
      aLinha   := {}    
      nDebito  := 0   
      nCredito := 0
      //If  ALLTRIM(RS_OK->E5_RECONC) = ""
         If RS_OK->E5_RECPAG = "P"
            If RS_OK->E5_TIPODOC = "ES"
               nDebito  := nDebito - RS_OK->E5_VALOR          
            Else   
               nDebito  := nDebito + RS_OK->E5_VALOR                      
            Endif   
         Endif  
         If RS_OK->E5_RECPAG = "R"
            If RS_OK->E5_TIPODOC = "ES"
               nCredito  := nCredito - RS_OK->E5_VALOR      
            Else   
               nCredito := nCredito + RS_OK->E5_VALOR                              
            Endif   
         Endif  
      //Endif   
      nTCredito += nCredito          
      nTDebito  += nDebito 
      nSaldoAtu := nSaldoAtu + nCredito - nDebito     
      For nX := 1 to LEN(aFields)-3       
          aAdd(aLinha,If(ALLTRIM(aFields[nX])="A6_NOME",ALLTRIM(pBanco)+"-"+ALLTRIM(pAgencia)+"-"+ALLTRIM(pConta)+"-"+ALLTRIM(pNomBco),RS_OK->&(aFields[nX])))      
      Next nX      
      aAdd(aLinha,nCredito)
      aAdd(aLinha,nDebito)                          
      aAdd(aLinha,nSaldoAtu)      
      aAdd(aArray,aLinha)  // len(aArray)      
      RS_OK->(DbSkip())        		              
Enddo  
aLinha := {}
For nX := 1 to LEN(aFields)-3       
    aAdd(aLinha,If(ALLTRIM(aHeaderEx[nX,8])="D",DTOS(MV_PAR10)," "))      
Next nX      
aLinha[5] := "SALDO FINAL"   
aLinha[6] := ALLTRIM(pBanco)+"-"+ALLTRIM(pAgencia)+"-"+ALLTRIM(pConta)+"-"+ALLTRIM(pNomBco)
aAdd(aLinha,nTCredito)
aAdd(aLinha,nTDebito)                          
aAdd(aLinha,nSaldoIni+nTCredito-nTDebito)      
aAdd(aArray,aLinha)
//aLinha[5] := " "   
//aLinha[6] := " "
//aAdd(aLinha,0)
//aAdd(aLinha,0)                          
//aAdd(aLinha,0)        
//aAdd(aArray,aLinha)
Return (aArray)   

//=============================================================================================================================
Static Function fLerSE8(pFilial,pBanco,pAgencia,pConta)
//=============================================================================================================================
Local aArray   := {0,0}
Local nCredito := 0
Local nDebito   := 0

If select("RS_OK") > 0
   dbselectarea("RS_OK") 
   dbCloseArea("RS_OK") 
Endif 
cQuery := " SELECT TOP 1 * FROM "+RetSqlName("SE8")
cQuery += " WHERE D_E_L_E_T_ = '' "
cQuery += " AND E8_FILIAL = '"+pFilial+"'"
cQuery += " AND E8_BANCO = '"+pBanco+"'"
cQuery += " AND E8_AGENCIA = '"+pAgencia+"'"
cQuery += " AND E8_CONTA = '"+pConta+"'"
cQuery += " AND E8_DTSALAT < '"+DTOS(MV_PAR09)+"'"
cQuery += " ORDER BY E8_DTSALAT DESC"
dbUseArea(.T., "TOPCONN", TCGenQry(,,cQuery), "RS_OK", .F., .T.) 
dbselectarea("RS_OK")  
DbGoTop()  
While RS_OK->(!Eof())  
      If RS_OK->E8_DTSALAT = DTOS(MV_PAR09)
         aArray[1] := 0
      Else             
         aArray[1] := RS_OK->E8_SALATUA
      Endif      
      RS_OK->(DbSkip())        		              
Enddo  
If select("RS_OK") > 0
   dbselectarea("RS_OK") 
   dbCloseArea("RS_OK") 
Endif 
cQuery := " SELECT TOP 1 * FROM "+RetSqlName("SE8")
cQuery += " WHERE D_E_L_E_T_ = '' "
cQuery += " AND E8_FILIAL = '"+pFilial+"'"
cQuery += " AND E8_BANCO = '"+pBanco+"'"
cQuery += " AND E8_AGENCIA = '"+pAgencia+"'"
cQuery += " AND E8_CONTA = '"+pConta+"'"
cQuery += " AND E8_DTSALAT = '"+DTOS(MV_PAR09)+"'"
dbUseArea(.T., "TOPCONN", TCGenQry(,,cQuery), "RS_OK", .F., .T.) 
dbselectarea("RS_OK")  
DbGoTop()  
While RS_OK->(!Eof())  
      aArray[2] := RS_OK->E8_SALATUA
      RS_OK->(DbSkip())        		              
Enddo  
Return (aArray)   

//=============================================================================================================================
Static Function fRelat(aHeaderEx,aColsEX) 
//=============================================================================================================================
//fRelat(aHeaderEx,aColsEX)
Local aTitulo := {}

aAdd(aTitulo,{"Extrato Bancário"})   
aAdd(aTitulo,{""})     
aAdd(aTitulo,{"REFERÊNCIA [ "+DTOC(MV_PAR09)+" a "+DTOC(MV_PAR10)+" ]"})   
nX := 1     

//U_ETX_REL(1,"ETX","Movimentação de Contas a Receber / Pagar","SE5",aTitulo,aHeaderEx,aColsEX,nX)  
U_ETX_RELT(1,"ETX_005","Extrato Bancário","SE5",aTitulo,aHeaderEx,aColsEX,nX)    
Return

//=============================================================================================================================
Static Function fFluxo() 
//=============================================================================================================================
U_ETX_004()
Return     
       
//=============================================================================================================================
Static Function fExcel(aHeaderEx,aColsEX) 
//=============================================================================================================================
Local aCabec    := {}
Local aPlanilha := {}

aAdd(aCabec,aHeaderEx)        
aAdd(aPlanilha,{aCabec,aColsEX,"Extrato Bancário","Extrato Bancário emitido em: ["+DTOC(dDataBase)+"] Período: ["+DTOC(MV_PAR09)+" / "+DTOC(MV_PAR10)+"]","1"})
ExecBlock("ETX_EXCEL",.F.,.F.,{Len(aHeaderEx),"Extrato Bancário",aPlanilha})                       
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