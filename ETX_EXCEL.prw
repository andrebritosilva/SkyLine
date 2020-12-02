#Include "Protheus.ch"
#Include "RWMake.ch"  
#Include "DBTREE.CH"   
#Include "HBUTTON.CH"    
//=============================================================================================================================   
//Programa............: ETX_EXCEL()
//Autor...............: Paulo César (PC) 
//Data................: 12/06/2019
//Descricao / Objetivo: Geração de Planila Excel 
//Cliente             : ETHOS X / SKYLINE
//=============================================================================================================================
User Function ETX_EXCEL()  //aHeaderEx,aItens,cArquivo,cTitulo     
//=============================================================================================================================
Private nCols      := PARAMIXB[1]  
Private cArquivo   := PARAMIXB[2]  
Private aHeaderEx  := {}
Private aItens     := {} 
Private cTitulo    := ""
Private cTipo      := "" 
Private aExceles   := {}
Private nExceles   := 0  
Private cWorkSheet := ""
Private oFWMSExcel  

For nExceles := 1 to Len(PARAMIXB[3]) //aAdd(aExceles,PARAMIXB[2,1])  aAdd(aExceles,PARAMIXB[2,2])  aAdd(aExceles,PARAMIXB[2,3])  aAdd(aExceles,PARAMIXB[2,4])  aAdd(aExceles,PARAMIXB[2,5) 
    aAdd(aExceles,PARAMIXB[3,nExceles])  
Next nExceles    
cArquivo := UPPER(ALLTRIM(cUSERNAME)+"-"+ALLTRIM(cArquivo)+"-"+REPLACE(DTOC(DDATABASE),"/","-")+"-"+REPLACE(TIME(),":","-")+".XML" ) 
cPathTmp := cGetFile( '', 'Selecione Diretório onde os arquivos serão gravados',0,,.F.,GETF_LOCALHARD+GETF_RETDIRECTORY+GETF_NETWORKDRIVE)
If ALLTRIM(cPathTmp) = "" 
   MSGALERT("Escolha uma Pasta para gravar o arquivo a ser gerado")
   Return
Endif 
oFWMSExcel := FWMSExcel():New()
For nExceles := 1 to Len(aExceles)    
    If Len(aExceles[nExceles]) > 0
       aHeaderEx  := aExceles[nExceles,1]
       aItens     := aExceles[nExceles,2]
       cWorkSheet := aExceles[nExceles,3]
       cTable     := aExceles[nExceles,4]
       cTipo      := aExceles[nExceles,5]
      f_Excel(aHeaderEx,aItens,cTable,cWorkSheet,cTipo)          
    Endif  
Next nExceles    
oFWMSExcel:Activate()
oFWMSExcel:GetXMLFile(ALLTRIM(cPathTmp)+ALLTRIM(cArquivo))
oFWMSExcel:= FreeObj(oFWMSExcel)
oFWMSExcel := MsExcel():New()           //Abre uma nova conexão com Excel
oFWMSExcel:WorkBooks:Open(ALLTRIM(cPathTmp)+ALLTRIM(cArquivo)) //Abre uma planilha
oFWMSExcel:SetVisible(.T.)              //Visualiza a planilha
oFWMSExcel:Destroy()                        //Encerra o processo do gerenciador de tarefas 
Return      

//=============================================================================================================================
Static Function f_Excel(aHeaderEx,aItens,cTable,cWorkSheet,cTipo)
//=============================================================================================================================
Local cTexto  := ""
Local nId     := 0
Local axCombo := {}
Private nXXX := 0    
//Private aHeaderEx  := p_aHeaderEx
//Private aItens     := p_aItens   
//Private cTable     := p_cTable   
//Private cWorkSheet := p_cWorkSheet
//Private cTipo      := p_cTipo
Private aCabec     := {}
Private aKey       := {}
Private cxNome     := {}
Private lTotalize  := .F.
Private lPicture   := .F.
Private cType      := "" 
Private nAux       := 0
Private cColumn    := ""
Private cTexto     := ""
Private aCells     := {}
Private cFile      := ""
Private cFileTMP   := ""
Private cPicture   := ""
Private lTotal     := .F.
Private nRow       := 0
Private nRows      := 0
Private nField     := 0
Private nFields    := 0
Private nAlign     := ""
Private nFormat    := 0                                
Private uCell      := ""
Private cPathTmp   := "" 
Private aExcel1    := {} 
Private oMsExcel    

aExcel1 := {}
For nRow := 1 to len(aHeaderEx) 
    //For nRows := 1 to len(aHeaderEx[nRow]) 
        aAdd(aExcel1,aHeaderEx[nRow])//,aHeaderEx[nRow,nRows,8],If(ALLTRIM(cTexto)="",aHeaderEx[nRow,nRows,4],100),aHeaderEx[nRow,nRows,5]})  
   // Next nRows 
Next nRow
f_Dados(cWorkSheet,cTable,aExcel1,aItens)  
Return

//=============================================================================================================================
// gravando o arquivo XML  =À gerando dados das linhas das abas do arquivo XML
//=============================================================================================================================
Static Function f_Dados(cWorkSheet,cTable,aExcel1,aItens)                                                      
//=============================================================================================================================
Local cTexto  := ""
Local nId     := 0
Local axCombo := {}
Private nQtde := 0
//aItens

IncProc("===> Registros do Arquivo XML-"+cTable)
oFWMSExcel:AddworkSheet(cWorkSheet)
oFWMSExcel:AddTable(cWorkSheet,cTable)
For nField := 1 To 1  
    For nQtde := 1 to Len(aExcel1[nField])        
        cType := aExcel1[nField,nQtde,8]
        nAlign := IF(cType=="C",1,IF(cType=="N",3,2))
        nFormat := IF(cType=="D",4,IF(cType=="N",2,1)) 
        cColumn := ALLTRIM(aExcel1[nField,nQtde,1])
        lTotal := ( lTotalize .and. cType == "N")
        oFWMSExcel:AddColumn(@cWorkSheet,@cTable,@cColumn,@nAlign,@nFormat,@lTotal)
    Next nQtde    
Next nField 
For nField := 2 to LEN(aExcel1)    
    nFields := LEN(aExcel1[1])
    aCells  := Array(nFields)
    For nQtde := 1 to Len(aExcel1[nField]) 
        cTexto := If(nQtde<nCols,"",aExcel1[nField,nQtde,1])
        uCell := cTexto
        aCells[nQtde] := uCell  
    Next nQtde    
    oFWMSExcel:AddRow(@cWorkSheet,@cTable,aClone(aCells))    
Next nField

nFields := LEN(aExcel1[1])
aCells  := Array(nFields)
nRows   := Len(aItens)
nQtde   := 1
nFields := LEN(aExcel1[1])
aCells := Array(nFields)

For nRow := nQtde To Len(aItens)
    For nField := 1 To nFields
       cTexto := ""   
       axCombo:= {}     
       cPicture := ""
        DbSelectArea("SX3")
        SX3->(DbSetOrder(2))
        If SX3->(DbSeek(aExcel1[1,nField,2]))
           cPicture := SX3->X3_PICTURE   
           If ALLTRIM(aExcel1[1,nField,11]) <> ""
              cTexto := '{"'+ALLTRIM(aExcel1[1,nField,11])+'"}' 
              cTexto := StrTran(cTexto,';','","')    
              axCombo := &(cTexto) 
              nId := aScan(axCombo,LEFT(ALLTRIM(aItens[nRow,nField]),2)+"=")
              If nId > 0   
                 cTexto := ALLTRIM(SUBSTR(axCombo[nId],3,100))
              Endif   
           Endif   
        Endif    
        uCell := If(ALLTRIM(cTexto)<>"",ALLTRIM(cTexto),aItens[nRow][nField])
        IF .NOT.(Empty(cPicture))
            uCell := Transform(uCell,cPicture)
        EndIF
       aCells[nField] := uCell
    Next nField
    oFWMSExcel:AddRow(@cWorkSheet,@cTable,aClone(aCells))
Next nRow
Return