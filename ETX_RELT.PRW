#INCLUDE "PROTHEUS.CH"        
#INCLUDE "TOPCONN.CH" 
#include "rwmake.ch"  
#include "fileio.ch"    
#INCLUDE "FWPrintSetup.ch"
#INCLUDE "RPTDEF.CH"
#Include "DBTREE.CH"
#Include "HBUTTON.CH"
#Define XENTERX Chr(13)+Chr(10) 
//==========================================================================================================================================================================================================================================================   
//Programa............: ETX_RELT()
//Autor...............: Paulo César (PC) 
//Data................: 17/05/2019
//Objetivo............: Função para Impressão Genérica de Relatórios
//==========================================================================================================================================================================================================================================================   
User Function ETX_RELT(pnCopias,pPrograma,pDescricao,pAlias,pTopo,pHeaderx,pItens,pTotalReg)
//==========================================================================================================================================================================================================================================================   
Local cDesc1         := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2         := "de acordo com os parametros informados pelo usuario."
Local cDesc3         := "1"
Local cPict          := ""
Local titulo         := ALLTRIM(pTopo[1,1])+" "+If(ALLTRIM(pTopo[2,1])<>"",ALLTRIM(pTopo[2,1]),"")+" "+If(ALLTRIM(pTopo[3,1])<>"",ALLTRIM(pTopo[3,1]),"")
Local nLin           := 80
Local Cabec1         := ""
Local Cabec2         := ""
Local imprime        := .T.
Local aOrd           := {}
Private lEnd         := .F.
Private lAbortPrint  := .F.
Private CbTxt        := ""
Private limite       := 220
Private tamanho      := "G"
Private nomeprog     := pPrograma // Coloque aqui o nome do programa para impressao no cabecalho
Private nTipo        := 18
Private aReturn      := { "Zebrado", 1, "Administracao", 2, 2, 1, "", 1}
Private nLastKey     := 0
Private cbtxt        := Space(10)
Private cbcont       := 00
Private CONTFL       := 01
Private m_pag        := 01
Private wnrel        := pPrograma // Coloque aqui o nome do arquivo usado para impressao em disco
Private cString      := "SE2"
Private aBox         := {}
Private aCabec       := {}
Private nX           := 0
Private nX1          := 0
Private nX2          := 0
Private nX3          := 0      
Private aDados       := pItens

For nX := 1 to Len(pHeaderx)
    If LEN(ALLTRIM(pHeaderx[nX,11])) > 0
       cLinha  := ""
       cLinha := "{'"+Replace(ALLTRIM(pHeaderx[nX,11]),";","','")+"'}" 
       aBox := aClone(&cLinha)            
       nX2 := pHeaderx[nX,4]          
       For nX1 := 1 to LEN(aBox)  
           If LEN(ALLTRIM(aBox[nX1])) > nX2
              nX2 := LEN(ALLTRIM(aBox[nX1]))
           Endif   
       Next nX1 
    Else   
       aBox := {}
   Endif
   aAdd(aCabec,{ALLTRIM(pHeaderx[nX,1]),;
                pHeaderx[nX,2],;
                pHeaderx[nX,3],;
                If(pHeaderx[nX,4]>=40,pHeaderx[nX,4]-10,pHeaderx[nX,4]),;
                pHeaderx[nX,5],;
                 pHeaderx[nX,6],;
                pHeaderx[nX,7],;
                pHeaderx[nX,8],;
                pHeaderx[nX,9],;
                pHeaderx[nX,10],;
                aBox,;
                pHeaderx[nX,12],;                             
                0})
   If LEN(ALLTRIM(pHeaderx[nX,11])) > 0 
      nX2 := pHeaderx[nX,4]          
      For nX1 := 1 to LEN(aBox)  
          If LEN(ALLTRIM(aBox[nX1])) > nX2
             nX2 := LEN(ALLTRIM(aBox[nX1]))
          Endif   
      Next nX1 
      aCabec[nX,4] := nX2
    Endif   
Next nX    
For nX := 1 to LEN(aCabec)     
    If ALLTRIM(aCabec[nX,8]) = "N"    
       If LEFT(aCabec[nX,3],2) = "@E"    
          nLin1 := LEN(ALLTRIM(SUBSTR(aCabec[nX,3],3,100)))
          If nLin1 > LEN(ALLTRIM(aCabec[nX,1])) 
             nLin1 := nLin1 - LEN(ALLTRIM(aCabec[nX,1]))  
             aCabec[nX,1] := SPACE(nLin1) + ALLTRIM(aCabec[nX,1])               
          Endif 
        Endif   
       aCabec[nX,4] := If(LEN(ALLTRIM(SUBSTR(aCabec[nX,3],3,100)))>aCabec[nX,4],LEN(ALLTRIM(SUBSTR(aCabec[nX,3],3,100))),aCabec[nX,4])
    Endif
Next nX
nX2 := 0
Cabec1 := ""
Cabec2 := ""
For nX := 1 to LEN(aCabec)
    aCabec[nX,13] := nX2
    Cabec1 += LEFT(aCabec[nX,1]+SPACE(100),aCabec[nX,4])+"#"
    Cabec2 += LEFT(Replicate("-",1000),aCabec[nX,4])+"#"  
     nX2 += aCabec[nX,4] + 1
Next nX    
Cabec1 := Replace(Cabec1,"#"," ")
Cabec2 := Replace(Cabec2,"#"," ")
wnrel := SetPrint(cString,NomeProg,"",@titulo,cDesc1,cDesc2,cDesc3,.T.,aOrd,.T.,Tamanho,,.T.)
If nLastKey == 27
	Return
Endif
SetDefault(aReturn,cString)
If nLastKey == 27
   Return
Endif
nTipo := If(aReturn[4]==1,15,18)
RptStatus({|| RunReport(Cabec1,Cabec2,Titulo,nLin) },Titulo)
Return

//==========================================================================================================================================================================================================================================================   
Static Function RunReport(Cabec1,Cabec2,Titulo,nLin)
//==========================================================================================================================================================================================================================================================   
Local nOrdem

nX  := 0
nX3 :=  LEN(aDados)
While nX <> nX3
    If lAbortPrint
       @nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
       Exit
    Endif
    nX += 1
    If nLin > 55 // Salto de Página. Neste caso o formulario tem 55 linhas...
       Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
       nLin := 8
    Endif       
    nLin := nLin + 1 // Avanca a linha de impressao         
    nX2 := 0
    For nX1 := 1 to LEN(aCabec)    //aCabec[nX1,2]
        If aCabec[nX1,8]="C"
           @nLin,aCabec[nX1,13] PSAY LEFT(ALLTRIM(aDados[nX,nX1])+SPACE(1000),aCabec[nX1,4])
        Endif   
        If aCabec[nX1,8]="N"
            @nLin,aCabec[nX1,13] PSAY aDados[nX,nX1] Picture aCabec[nX1,3]
        Endif   
        If aCabec[nX1,8]="D"
            @nLin,aCabec[nX1,13] PSAY aDados[nX,nX1] 
        Endif             
     Next nX1
Enddo
SET DEVICE TO SCREEN
If aReturn[5]==1
   dbCommitAll()
   SET PRINTER TO
   OurSpool(wnrel)
Endif
MS_FLUSH()
Return