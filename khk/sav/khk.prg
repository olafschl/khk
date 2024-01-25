**
** khk.fxp
**
 PARAMETER usEmemberfield
 PRIVATE deNd
 PRIVATE dsTart
 PRIVATE ncHoice
 PRIVATE cfIle
 PRIVATE nhPorder
 PRIVATE naRea
 PRIVATE cwRitebuffer
 PRIVATE nbLockcount
 PRIVATE nlAstblockposition
 PRIVATE naBrnr
 IF  .NOT. chKlicense()
      RETURN
 ENDIF
 usEmemberfield = (PARAMETERS()<>0)
 naRea = SELECT()
 cwRitebuffer = ""
 nbLockcount = 0
 nlAstblockposition = 0
 naBrnr = 0
 nhPorder = ORDER("HistPost")
 dsTart = DATE()
 deNd = DATE()
 ncHoice = 1
 = dv_window(0)
 @ 1, 2 SAY geTlangtext("DATEV","TXT_START_DATE") FONT 'ARIAL', 10 SIZE 20, 22
 @ 3, 2 SAY geTlangtext("DATEV","TXT_END_DATE") FONT 'ARIAL', 10 SIZE 20, 22
 @ 1, 19 GET dsTart FONT 'ARIAL', 10 SIZE 1, 12 PICTURE "@K" VALID  .NOT.  ;
   EMPTY(dsTart)
 @ 3, 19 GET deNd FONT 'ARIAL', 10 SIZE 1, 12 PICTURE "@K" VALID  .NOT.  ;
   EMPTY(deNd)
 @ 7, 2 GET ncHoice FONT 'ARIAL', 10 STYLE "N" SIZE nbUttonheight, 17  ;
   FUNCTION "*"+"H" PICTURE "\!"+geTlangtext("COMMON","TXT_OK")+";\?"+ ;
   geTlangtext("COMMON","TXT_CANCEL")
 READ CYCLE MODAL
 = dv_window(1)
 IF (ncHoice==1)
      SELECT hiStpost
      SET ORDER TO 2
      SET NEAR ON
      IF paRam.pa_version>=6.64
           SEEK dsTart
      ELSE
           SEEK DTOS(dsTart)
      ENDIF
      SET NEAR OFF
      IF (EOF() .OR. hiStpost.hp_date>deNd)
           = alErt(geTlangtext("DATEV","TXT_NO_DATA_FOUND"))
      ELSE
           IF (yeSno(geTlangtext("DATEV","TXT_ARE_YOU_SURE")))
                lsTop = .F.
                cfile=SYS(5)+SYS(2003)+"\export\KHKEXPORT.TXT"
                IF (EMPTY(cfIle))
                     lsTop = .T.
                ENDIF
                IF ( .NOT. lsTop)
                     IF ( .NOT. dv_dexxxfile(dsTart,deNd,cfIle))
                          = alErt(geTlangtext("DATEV", ;
                            "TXT_NO_DE_FILE_CREATED"))
                     ELSE
                      *    IF ( .NOT. dv_dv01file(dsTart,deNd,cfIle))
                      *         = alErt(geTlangtext("DATEV", ;
                      *           "TXT_NO_DV01_FILE_CREATED"))
                      *    ELSE
                      *         = alErt(geTlangtext("DATEV", ;
                      *           "TXT_EXP_FILES_CREATED"))
                      *    ENDIF
                     ENDIF
                ENDIF
           ENDIF
      ENDIF
 ENDIF
 SET ORDER IN hiStpost TO nHpOrder
 SELECT (naRea)
 RETURN .T.
ENDFUNC
*
FUNCTION DV_Window
 PARAMETER naCtivate
 IF (naCtivate==0)
      DEFINE WINDOW wdAtev AT 0, 0 SIZE 10, 30 NOGROW FLOAT CLOSE NOZOOM  ;
             TITLE 'KHK Export V 2.01' ICON FILE "hotel.ico"
      MOVE WINDOW wdAtev CENTER
      ACTIVATE WINDOW wdAtev
      = paNelborder()
 ELSE
      DEACTIVATE WINDOW wdAtev
      RELEASE WINDOW wdAtev
      = chIldtitle("")
 ENDIF
 RETURN .T.
ENDFUNC
*
FUNCTION DV_DV01File
 PARAMETER dsTart, deNd, cdVfile
 PRIVATE nhAndle
 PRIVATE lsUccess
 cdVfile = SUBSTR(cfIle, 1, LEN(cfIle)-6)+"DV01"
 lsUccess = .F.
 nhAndle = FCREATE(cdVfile)
 IF (nhAndle==-1)
      = alErt(geTlangtext("DATEV","TXT_CREATE_ERROR"))
 ELSE
      *= FWRITE(nhAndle,"BUDAT;KTOSOLL;KTOHABEN;TEXT;RENR;REDAT;MWST;BETRAG"+CHR(13)+CHR(10))
      =FWRITE(nhandle,cKopf+CHR(13)+CHR(10))
      = FCLOSE(nhAndle)
      lsUccess = .T.
 ENDIF
 RETURN lsUccess
ENDFUNC
*
FUNCTION DV_DExxxFile
 PARAMETER dsTart, deNd, cfIle
 PRIVATE lsUccess
 PRIVATE nhAndle
 PRIVATE ntOtalamount
 PRIVATE cpAytext, cbIllnr
 PRIVATE csKiprooms
 PRIVATE cSatz,cKopf
 = dv_exportfile()
 lsUccess = .F.
 nhAndle = FCREATE(cfIle)
 IF (nhAndle==-1)
      = alErt(geTlangtext("DATEV","TXT_CREATE_ERROR"))
 ELSE
 	  *cbuffer="BUDAT;KTOSOLL;KTOHABEN;TEXT;RENR;REDAT;MWST;BETRAG"+CHR(13)+CHR(10)
      *cbUffer = f_Write(nhAndle,cbUffer)
      csKiprooms = ','+ALLTRIM(dv_param("Ausnahme","Zimmer","",80))+','
      cgEgenkonto = dv_param("Zahlung","Gegenkonto","87654321",8)
      ckOnto = dv_param("Artikel","Konto","12345678",8)
      cEArtikel=dv_param("Artikel","Ersatzkonto","12345678",8)
      cEZahlung=dv_param("Zahlung","Ersatzkonto","12345678",8)
      cKopf=getparam("Buchungssatz","Kopf","DATEV.INI")
      cSatz=getparam("Buchungssatz","Satz","DATEV.INI")
      caMacro = "Article.Ar_Lang"+g_Langnum
      cpMacro = "Paymetho.Pm_Lang"+g_Langnum
      cbuffer=cKopf+CHR(13)+CHR(10)
      * cbuffer=f_write(nhandle,cb)
      SELECT hiStpost
      SET ORDER TO 2
      SET NEAR ON
      IF paRam.pa_version>=6.64
           SEEK dsTart
      ELSE
           SEEK DTOS(dsTart)
      ENDIF
      SET NEAR OFF
      DO WHILE ( .NOT. EOF("HistPost") .AND. hiStpost.hp_date<=deNd)
           ddAte = hiStpost.hp_date
           WAIT WINDOW NOWAIT "1.21"+" "+DTOC(ddAte)
           DO WHILE ( .NOT. EOF("HistPost") .AND. hiStpost.hp_date==ddAte)
                SELECT hiStres
                SET ORDER TO 1
                IF paRam.pa_version>=6.64
                     SEEK hiStpost.hp_reserid
                ELSE
                     SEEK STR(hiStpost.hp_reserid, 12, 3)
                ENDIF
                SELECT hiStpost
                IF (','+TRIM(hiStres.hr_roomnum)+','$csKiprooms AND !EMPTY(histres.hr_roomnum));
                	OR INLIST(histres.hr_reserid,0.300,0.400)
                     SKIP 1 IN hiStpost
                     LOOP
                ENDIF
                DO CASE
                     CASE ( .NOT. EMPTY(hiStpost.hp_artinum) .AND.  ;
                          (EMPTY(hiStpost.hp_ratecod) .OR.  ;
                          hiStpost.hp_split) .AND. hiStpost.hp_reserid>0  ;
                          .AND.  .NOT. hiStpost.hp_cancel)
                          SELECT daTev
                          SET ORDER TO 1
                          IF ( .NOT. SEEK(DTOS(ddAte)+ ;
                             STR(hiStpost.hp_artinum, 4), "Datev"))
                               IF paRam.pa_version>=6.64
                                    = SEEK(hiStpost.hp_artinum, "Article")
                               ELSE
                                    = SEEK(hiStpost.hp_departm+ ;
                                      STR(hiStpost.hp_artinum, 4), "Article")
                               ENDIF
                               SELECT daTev
                               APPEND BLANK
                               REPLACE daTev.arTikel WITH hiStpost.hp_artinum
                               IF !empty(article.ar_user2)
	                               REPLACE daTev.geGenkonto WITH  ;
    	                                   ALLTRIM(arTicle.ar_user2)
    	                       ELSE
    	                       	   replace datev.gegenkonto WITH ALLTRIM(cEArtikel)
    	                       endif   
                               REPLACE daTev.daTum WITH ddAte
                               Replace Datev.Text       With &cAMacro
                               REPLACE daTev.koNto WITH ckOnto
                               replace datev.beleg1 WITH "Citadel-Export"
                               mwstges=histpost.hp_vat0+histpost.hp_vat1+histpost.hp_vat2+histpost.hp_vat3+histpost.hp_vat4+histpost.hp_vat5+;
                               	histpost.hp_vat6+histpost.hp_vat7+histpost.hp_vat8+histpost.hp_vat9
                               * replace datev.mwst	WITH IIF(mwstges>0,"1","0")
                               cmwst=getparam("MWST",STR(article.ar_vat,1,0),"DATEV.INI")
                               replace datev.mwst WITH IIF(EMPTY(cmwst),"0",ALLTRIM(cmwst))
                          ENDIF
                          REPLACE daTev.umSatz WITH daTev.umSatz+ ;
                                  hiStpost.hp_amount
                     CASE ( .NOT. EMPTY(hiStpost.hp_paynum) .AND.  ;
                          hiStpost.hp_reserid>0 .AND.  .NOT.  ;
                          hiStpost.hp_cancel)
                          SELECT daTev
                          SET ORDER TO 2
                          SELECT paYmetho
                          LOCATE FOR paYmetho.pm_paynum=hiStpost.hp_paynum
                          IF (usEmemberfield .AND. paYmetho.pm_paytyp==4)
                               SELECT adDress
                               IF ( .NOT. EMPTY(hiStres.hr_compid))
                                    LOCATE FOR ad_addrid=hiStres.hr_compid
                               ELSE
                                    LOCATE FOR ad_addrid=hiStres.hr_addrid
                               ENDIF
                               cbIllnr = IIF(hiStpost.hp_window=3,  ;
                                hiStres.hr_billnr3,  ;
                                IIF(hiStpost.hp_window=2,  ;
                                hiStres.hr_billnr2, hiStres.hr_billnr1))
                               IF AT('-', cbIllnr)>0
                                    cbIllnr = SUBSTR(cbIllnr, 1, 10)
                               ENDIF
                               cbEleg1 = RIGHT(cbIllnr, 10)
                               cbEleg2 = ''
                               cpAytext = IIF(EMPTY(histres.hr_lname),histres.hr_company,ALLTRIM(hiStres.hr_lname+" / "+histres.hr_company))
                               cpAykonto = ALLTRIM(STR(adDress.ad_member))
                               IF (EMPTY(VAL(cpAykonto)))
                               		IF !empty(paymetho.pm_user2)
	                                    cpAykonto = ALLTRIM(paYmetho.pm_user2)
	                                ELSE
	                                	cpaykonto = ALLTRIM(cEZahlung)
	                                endif
                               ENDIF
                          ELSE
                               cbEleg1 = ''
                               cbEleg2 = ''
                               cPayText = &cPMacro
                               cbIllnr = IIF(hiStpost.hp_window=3,  ;
                                hiStres.hr_billnr3,  ;
                                IIF(hiStpost.hp_window=2,  ;
                                hiStres.hr_billnr2, hiStres.hr_billnr1))
                               IF AT('-', cbIllnr)>0
                                    cbIllnr = SUBSTR(cbIllnr, 1, 10)
                               ENDIF
                               cbEleg1 = RIGHT(cbIllnr, 10)
                               cbEleg2 = ''
                               cpAytext = IIF(EMPTY(histres.hr_lname),histres.hr_company,ALLTRIM(hiStres.hr_lname+" / "+histres.hr_company))
							   IF histres.hr_reserid=0.100
							   		cpaytext="Passantenbuchung"
							   endif
                               IF !EMPTY(paymetho.pm_user2)
	                               cpAykonto = ALLTRIM(paYmetho.pm_user2)
	                           ELSE
	                           		cpaykonto = ALLTRIM(cEZahlung)
	                           endif
                          ENDIF
                          *IF  .NOT. SEEK(DTOS(ddAte)+ ;
                          *   STR(hiStpost.hp_paynum, 2)+cpAykonto,  ;
                          *    "Datev") .OR. (usEmemberfield .AND.  ;
                          *    paYmetho.pm_paytyp==4)
                          * WAIT WINDOW "Buchung Zahlung  "+cpaytext
                               SELECT daTev
                               APPEND BLANK
                               REPLACE daTev.zaHlung WITH hiStpost.hp_paynum
                               REPLACE daTev.geGenkonto WITH cgEgenkonto
                               REPLACE daTev.beLeg1 WITH IIF(EMPTY(cbeleg1),"Citadel Transfer",cbEleg1),  ;
                                       daTev.beLeg2 WITH cbEleg2
                               REPLACE daTev.daTum WITH ddAte
                               REPLACE daTev.teXt WITH cpAytext
                               REPLACE daTev.koNto WITH cpAykonto
                          *ENDIF
                          REPLACE daTev.umSatz WITH daTev.umSatz+ ;
                                  (hiStpost.hp_amount*-1)
                          *brow
                ENDCASE
                SKIP 1 IN hiStpost
           ENDDO
      ENDDO
      SELECT daTev
      DELETE ALL FOR umSatz=0
      SET ORDER TO 1
      GOTO TOP
      ntOtalamount = 0
      DO WHILE ( .NOT. EOF("Datev"))
           ntOtalamount = ntOtalamount+daTev.umSatz
           *cbuffer=cbuffer+DTOC(datev.datum)+";"+datev.konto+";"+datev.gegenkonto+";"+;
           	*datev.text+";"+datev.beleg1+";"+DTOC(datev.datum)+";"+datev.mwst+";"+;
           	*ALLTRIM(STR(datev.umsatz,12,2))+CHR(13)+CHR(10)
           cbuffer=cbuffer+&cSatz+CHR(13)+CHR(10)
           SELECT daTev
           SKIP 1
           * cbUffer = f_Write(nhAndle,cbUffer)
      ENDDO
      IF LEN(cbuffer)>0
      	=FWRITE(nhandle,cbuffer)
      endif
      * IF (LEN(cwRitebuffer)>0)
       *    nlAstblockposition = LEN(cwRitebuffer)-1
       *    cbUffer = f_Write(nhAndle,"")
      *ENDIF
      = FCLOSE(nhAndle)
      lsUccess = .T.
 ENDIF
 prntrep = MESSAGEBOX( ;
          "Soll ein Bericht gerdruckt werden?",  ;
          36, "Datev Export")
 IF prntrep = 6
     SELECT * FROM datev ORDER BY  ;
              datum, zahlung,  ;
              artikel INTO CURSOR  ;
              datevtmp
     REPORT FORM report\datev01  ;
            TO PRINTER PROMPT  ;
            NOCONSOLE
 ENDIF
 USE IN daTev
 WAIT CLEAR
 RETURN lsUccess
ENDFUNC
*
FUNCTION Binair
 PARAMETER nnUmber
 PRIVATE nhIgh1, nhIgh2
 PRIVATE nlOw1, nlOw2
 nhIgh1 = INT(nnUmber/4096)
 nhIgh2 = INT((nnUmber-(nhIgh1*4096))/256)
 nlOw1 = INT((nnUmber-(nhIgh1*4096)-(nhIgh2*256))/16)
 nlOw2 = nnUmber-(nhIgh1*4096)-(nhIgh2*256)-(nlOw1*16)
 RETURN CHR(nlOw1*16+nlOw2)+CHR(nhIgh1*16+nhIgh2)
ENDFUNC
*
FUNCTION DV_Param
 PARAMETER csEction, cpAram, cdEfault, nlEngth
 RETURN PADR(geTparam(csEction,cpAram,"DATEV.INI",cdEfault), nlEngth)
ENDFUNC
*
FUNCTION DV_Date
 PARAMETER ddAte
 RETURN stRzero(DAY(ddAte),2)+stRzero(MONTH(ddAte),2)+ ;
        SUBSTR(STR(YEAR(ddAte), 4), 3, 2)
ENDFUNC
*
FUNCTION DV_ExportFile
 CREATE CURSOR Datev (arTikel N (4, 0), zaHlung N (2, 0), umSatz N (12,  ;
        2), geGenkonto C (8), daTum D (8), koNto C (8), teXt C (30),  ;
        beLeg1 C (10), beLeg2 C (10), mwst C(1))
 SELECT daTev
 INDEX ON DTOS(daTum)+STR(arTikel, 4) TAG taG1
 INDEX ON DTOS(daTum)+STR(zaHlung, 2)+koNto TAG taG2
 SET ORDER TO 1
 RETURN .T.
ENDFUNC
*
FUNCTION DV_WriteParam
 PARAMETER csEction, cpArameter, cvAlue
 RETURN .T.
ENDFUNC
*
FUNCTION f_Write
 PARAMETER nhAndle, cdAta
 IF  .NOT. EMPTY(cdAta) && .AND. (256-(LEN(cwRitebuffer)+LEN(cdAta))>0))
      cwRitebuffer = cwRitebuffer+cdAta
 ELSE
      && = FWRITE(nhAndle, PADR(cwRitebuffer, 256, CHR(0)))
      cwRitebuffer = cdAta
      nbLockcount = nbLockcount+1
 ENDIF
 RETURN ""
ENDFUNC
*
FUNCTION ChkLicense
 PRIVATE ncC, ncOunt, ctMp, nlIc, lrEt
 ncC = 0
 ctMp = PADR(paRam.pa_hotel, 30)+"OsHk991"+PADR(paRam.pa_city, 30)+ ;
        PADR("ACCT", 8)
 FOR ncOunt = 1 TO LEN(ctMp)
      ncC = ncC+(ncOunt*ASC(SUBSTR(ctMp, ncOunt, 1)))
 ENDFOR
 nlIc = VAL(dv_param("License","Code","0",10))
 IF nlIc<>ncC
      = alErt( ;
        "Ungültiger Lizenskode für Datev Schnittstelle!;;Bitte tragen Sie den Lizenskode ein in die Datei 'DATEV.INI'." ;
        )
      lrEt = .F.
 ELSE
      lrEt = .T.
 ENDIF
 RETURN lrEt
ENDFUNC
*
