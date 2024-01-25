*** 
*** ReFox MMII (Win) #UK970148  SCHLINGMEIER  CITADEL [VFP70]
***
PARAMETER usememberfield
PRIVATE dend
PRIVATE dstart
PRIVATE nchoice
PRIVATE cfile
PRIVATE nhporder
PRIVATE narea
PRIVATE cwritebuffer
PRIVATE nblockcount
PRIVATE nlastblockposition
PRIVATE nabrnr
IF  .NOT. chklicense()
     RETURN
ENDIF
usememberfield = (PARAMETERS() <> 0)
narea = SELECT()
cwritebuffer = ""
nblockcount = 0
nlastblockposition = 0
nabrnr = 0
nhporder = ORDER("HistPost")
dstart = DATE()
dend = DATE()
nchoice = 1
= dv_window(0)
@ 1, 2 SAY getlangtext("DATEV","TXT_START_DATE") FONT 'ARIAL', 10 SIZE 20, 22
@ 3, 2 SAY getlangtext("DATEV","TXT_END_DATE") FONT 'ARIAL', 10 SIZE 20, 22
@ 1, 19 GET dstart FONT 'ARIAL', 10 SIZE 1, 12 PICTURE "@K" VALID  .NOT. EMPTY(dstart)
@ 3, 19 GET dend FONT 'ARIAL', 10 SIZE 1, 12 PICTURE "@K" VALID  .NOT. EMPTY(dend)
@ 7, 2 GET nchoice FONT 'ARIAL', 10 STYLE "N" SIZE nbuttonheight, 17 FUNCTION "*" + "H" PICTURE "\!" + getlangtext("COMMON","TXT_OK") + ";\?" + getlangtext("COMMON","TXT_CANCEL")
READ CYCLE MODAL
= dv_window(1)
IF (nchoice == 1)
     SELECT histpost
     SET ORDER TO 2
     SET NEAR ON
     IF param.pa_version >= 6.64 
          SEEK dstart 
     ELSE
          SEEK DTOS(dstart) 
     ENDIF
     SET NEAR OFF
     IF (EOF() .OR. histpost.hp_date > dend)
          = alert(getlangtext("DATEV","TXT_NO_DATA_FOUND"))
     ELSE
          IF (yesno(getlangtext("DATEV","TXT_ARE_YOU_SURE")))
               lstop = .F.
               cfile = SYS(5) + SYS(2003) + "\export\KHKEXPORT.TXT"
               IF (EMPTY(cfile))
                    lstop = .T.
               ENDIF
               IF ( .NOT. lstop)
                    IF ( .NOT. dv_dexxxfile(dstart,dend,cfile))
                         = alert(getlangtext("DATEV","TXT_NO_DE_FILE_CREATED"))
                    ELSE
                    ENDIF
               ENDIF
          ENDIF
     ENDIF
ENDIF
SET ORDER IN histpost TO nHpOrder
SELECT (narea)
RETURN .T.
ENDFUNC
*
FUNCTION DV_Window
PARAMETER nactivate
IF (nactivate == 0)
     DEFINE WINDOW wdatev AT 0, 0 SIZE 10, 30 NOGROW FLOAT CLOSE NOZOOM TITLE 'KHK Export V 2.01' ICON FILE "hotel.ico"
     MOVE WINDOW wdatev CENTER
     ACTIVATE WINDOW wdatev
     = panelborder()
ELSE
     DEACTIVATE WINDOW wdatev
     RELEASE WINDOW wdatev
     = childtitle("")
ENDIF
RETURN .T.
ENDFUNC
*
FUNCTION DV_DV01File
PARAMETER dstart, dend, cdvfile
PRIVATE nhandle
PRIVATE lsuccess
cdvfile = SUBSTR(cfile, 1, LEN(cfile) - 6) + "DV01"
lsuccess = .F.
nhandle = FCREATE(cdvfile)
IF (nhandle == -1)
     = alert(getlangtext("DATEV","TXT_CREATE_ERROR"))
ELSE
     = FWRITE(nhandle, ckopf + CHR(13) + CHR(10))
     = FCLOSE(nhandle)
     lsuccess = .T.
ENDIF
RETURN lsuccess
ENDFUNC
*
FUNCTION DV_DExxxFile
PARAMETER dstart, dend, cfile
PRIVATE lsuccess
PRIVATE nhandle
PRIVATE ntotalamount
PRIVATE cpaytext, cbillnr
PRIVATE cskiprooms
PRIVATE csatz, ckopf
= dv_exportfile()
lsuccess = .F.
nhandle = FCREATE(cfile)
IF (nhandle == -1)
     = alert(getlangtext("DATEV","TXT_CREATE_ERROR"))
ELSE
     cskiprooms = ',' + ALLTRIM(dv_param("Ausnahme","Zimmer","",80)) + ','
     cgegenkonto = dv_param("Zahlung","Gegenkonto","87654321",8)
     ckonto = dv_param("Artikel","Konto","12345678",8)
     ceartikel = dv_param("Artikel","Ersatzkonto","12345678",8)
     cezahlung = dv_param("Zahlung","Ersatzkonto","12345678",8)
     ckopf = getparam("Buchungssatz","Kopf","DATEV.INI")
     csatz = getparam("Buchungssatz","Satz","DATEV.INI")
     cintern=ALLTRIM(getparam("Artikel","Intern","",3))
     IF !INLIST(cintern,"N","J")
     	cintern="N"
     endif
     camacro = "Article.Ar_Lang" + g_langnum
     cpmacro = "Paymetho.Pm_Lang" + g_langnum
     cbuffer = ckopf + CHR(13) + CHR(10)
     SELECT histpost
     SET ORDER TO 2
     SET NEAR ON
     IF param.pa_version >= 6.64 
          SEEK dstart 
     ELSE
          SEEK DTOS(dstart) 
     ENDIF
     SET NEAR OFF
     DO WHILE ( .NOT. EOF("HistPost") .AND. histpost.hp_date <= dend)
          ddate = histpost.hp_date
          WAIT WINDOW NOWAIT "1.21" + " " + DTOC(ddate)
          DO WHILE ( .NOT. EOF("HistPost") .AND. histpost.hp_date == ddate)
               SELECT histres
               SET ORDER TO 1
               IF param.pa_version >= 6.64 
                    SEEK histpost.hp_reserid 
               ELSE
                    SEEK STR(histpost.hp_reserid, 12, 3) 
               ENDIF
               SELECT histpost
               IF (',' + TRIM(histres.hr_roomnum) + ',' $ cskiprooms .AND.  .NOT. EMPTY(histres.hr_roomnum)) .OR. INLIST(histres.hr_reserid, 0.300 , 0.400 )
                    SKIP 1 IN histpost
                    LOOP
               ENDIF
               DO CASE
                    CASE ( .NOT. EMPTY(histpost.hp_artinum) .AND. (EMPTY(histpost.hp_ratecod) .OR. histpost.hp_split) .AND. histpost.hp_reserid > 0 .AND.  .NOT. histpost.hp_cancel)
                         SELECT datev
                         SET ORDER TO 1
                         IF ( .NOT. SEEK(DTOS(ddate) + STR(histpost.hp_artinum, 4), "Datev"))
                              IF param.pa_version >= 6.64 
                                   = SEEK(histpost.hp_artinum, "Article")
                              ELSE
                                   = SEEK(histpost.hp_departm + STR(histpost.hp_artinum, 4), "Article")
                              ENDIF
                              IF INLIST(article.ar_artityp, 1, 2) OR (cintern="J" AND article.ar_artityp=3)
                                   SELECT datev
                                   APPEND BLANK
                                   REPLACE datev.artikel WITH histpost.hp_artinum
                                   IF  .NOT. EMPTY(article.ar_user2)
                                        REPLACE datev.gegenkonto WITH ALLTRIM(article.ar_user2)
                                   ELSE
                                        REPLACE datev.gegenkonto WITH ALLTRIM(ceartikel)
                                   ENDIF
                                   replace datev.kst WITH ALLTRIM(article.ar_user1)
                                   REPLACE datev.datum WITH ddate
                                   Replace Datev.Text       With &cAMacro
                                   REPLACE datev.konto WITH ckonto
                                   REPLACE datev.beleg1 WITH "Citadel-Export"
                                   mwstges = histpost.hp_vat0 + histpost.hp_vat1 + histpost.hp_vat2 + histpost.hp_vat3 + histpost.hp_vat4 + histpost.hp_vat5 + histpost.hp_vat6 + histpost.hp_vat7 + histpost.hp_vat8 + histpost.hp_vat9
                                   cmwst = getparam("MWST",STR(article.ar_vat, 1, 0),"DATEV.INI")
                                   REPLACE datev.mwst WITH IIF(EMPTY(cmwst), "0", ALLTRIM(cmwst))
                              ENDIF
                         ENDIF
                         REPLACE datev.umsatz WITH datev.umsatz + histpost.hp_amount
                    CASE ( .NOT. EMPTY(histpost.hp_paynum) .AND. histpost.hp_reserid > 0 .AND.  .NOT. histpost.hp_cancel)
                       SELECT datev
                       SET ORDER TO 2
                       SELECT paymetho
                       LOCATE FOR paymetho.pm_paynum = histpost.hp_paynum
                       IF paymetho.pm_paytyp<>8 OR (paymetho.pm_paytyp=8 AND cintern="J")
                         IF (usememberfield .AND. paymetho.pm_paytyp == 4)
                              SELECT address
                              IF ( .NOT. EMPTY(histres.hr_compid))
                                   LOCATE FOR ad_addrid = histres.hr_compid
                              ELSE
                                   LOCATE FOR ad_addrid = histres.hr_addrid
                              ENDIF
                              DO case
                              		CASE histpost.hp_window=1
                              			cbillnr=histres.hr_billnr1
                              		CASE histpost.hp_window=2
                              			cbillnr=histres.hr_billnr2
                              		CASE histpost.hp_window=3
                              			cbillnr=histres.hr_billnr3
                              		CASE histpost.hp_window=4
                              			cbillnr=histres.hr_billnr4
                              		CASE histpost.hp_window=5
                              			cbillnr=histres.hr_billnr5
                              		CASE histpost.hp_window=6
                              			cbillnr=histres.hr_billnr6
							  endcase 
                              IF AT('-', cbillnr) > 0
                                   cbillnr = SUBSTR(cbillnr, 1, 10)
                              ENDIF
                              cbeleg1 = RIGHT(cbillnr, 10)
                              cbeleg2 = ''
                              cpaytext = IIF(EMPTY(histres.hr_lname), histres.hr_company, ALLTRIM(histres.hr_lname + " / " + histres.hr_company))
                              IF address.ad_member>0
	                              cpaykonto = "D"+ALLTRIM(STR(address.ad_member))
	                          else
	                       	  	  cpaykonto = ""
	                       	  endif
                              IF (EMPTY(VAL(cpaykonto)))
                                   IF  .NOT. EMPTY(paymetho.pm_user2)
                                        cpaykonto = ALLTRIM(paymetho.pm_user2)
                                   ELSE
                                        cpaykonto = ALLTRIM(cezahlung)
                                   ENDIF
                              ENDIF
                         ELSE
                              cbeleg1 = ''
                              cbeleg2 = ''
                              cbillnr=''
                              cPayText = &cPMacro
                              DO case
                              		CASE histpost.hp_window=1
                              			cbillnr=histres.hr_billnr1
                              		CASE histpost.hp_window=2
                              			cbillnr=histres.hr_billnr2
                              		CASE histpost.hp_window=3
                              			cbillnr=histres.hr_billnr3
                              		CASE histpost.hp_window=4
                              			cbillnr=histres.hr_billnr4
                              		CASE histpost.hp_window=5
                              			cbillnr=histres.hr_billnr5
                              		CASE histpost.hp_window=6
                              			cbillnr=histres.hr_billnr6
							  endcase 
                              IF AT('-', cbillnr) > 0
                                   cbillnr = SUBSTR(cbillnr, 1, 10)
                              ENDIF
                              cbeleg1 = RIGHT(cbillnr, 10)
                              cbeleg2 = ''
                              cpaytext = IIF(EMPTY(histres.hr_lname), histres.hr_company, ALLTRIM(histres.hr_lname + " / " + histres.hr_company))
                              IF histres.hr_reserid = 0.100 
                              	IF histpost.hp_cashier=98
                              		IF histpost.hp_paynum<50
	                              		cpaytext="Rest-"+paymetho.pm_lang3
	                              	ELSE
	                              		cpaytext="Resi-"+paymetho.pm_lang3
	                              	endif
                              	else
                                	cpaytext = "-Passantenbuchung"
                                endif
                              ENDIF
                              IF  .NOT. EMPTY(paymetho.pm_user2)
                                   cpaykonto = ALLTRIM(paymetho.pm_user2)
                              ELSE
                                   cpaykonto = ALLTRIM(cezahlung)
                              ENDIF
                         ENDIF
                         SELECT datev
                         APPEND BLANK
                         REPLACE datev.zahlung WITH histpost.hp_paynum
                         REPLACE datev.gegenkonto WITH cgegenkonto
                         REPLACE datev.beleg1 WITH IIF(EMPTY(cbeleg1), "Citadel Transfer", cbeleg1), datev.beleg2 WITH cbeleg2
                         REPLACE datev.datum WITH ddate
                         REPLACE datev.text WITH cpaytext
                         REPLACE datev.konto WITH cpaykonto
                         REPLACE datev.umsatz WITH datev.umsatz + (histpost.hp_amount * -1)
                      endif
               ENDCASE
               SKIP 1 IN histpost
          ENDDO
     ENDDO
     SELECT datev
     DELETE ALL FOR umsatz = 0
     SET ORDER TO 1
     GOTO TOP
     ntotalamount = 0
     DO WHILE ( .NOT. EOF("Datev"))
          ntotalamount = ntotalamount + datev.umsatz
          cbuffer=cbuffer+&cSatz+CHR(13)+CHR(10)
          SELECT datev
          SKIP 1
     ENDDO
     IF LEN(cbuffer) > 0
          = FWRITE(nhandle, cbuffer)
     ENDIF
     = FCLOSE(nhandle)
     lsuccess = .T.
ENDIF
prntrep = MESSAGEBOX("Soll ein Bericht gerdruckt werden?", 36, "Datev Export")
IF prntrep = 6
     SELECT * FROM datev ORDER BY datum, zahlung, artikel INTO CURSOR datevtmp
     REPORT FORM report\datev01 TO PRINTER PROMPT NOCONSOLE
ENDIF
USE IN datev
WAIT CLEAR
RETURN lsuccess
ENDFUNC
*
FUNCTION Binair
PARAMETER nnumber
PRIVATE nhigh1, nhigh2
PRIVATE nlow1, nlow2
nhigh1 = INT(nnumber / 4096)
nhigh2 = INT((nnumber - (nhigh1 * 4096)) / 256)
nlow1 = INT((nnumber - (nhigh1 * 4096) - (nhigh2 * 256)) / 16)
nlow2 = nnumber - (nhigh1 * 4096) - (nhigh2 * 256) - (nlow1 * 16)
RETURN CHR(nlow1 * 16 + nlow2) + CHR(nhigh1 * 16 + nhigh2)
ENDFUNC
*
FUNCTION DV_Param
PARAMETER csection, cparam, cdefault, nlength
RETURN PADR(getparam(csection,cparam,"DATEV.INI",cdefault), nlength)
ENDFUNC
*
FUNCTION DV_Date
PARAMETER ddate
RETURN strzero(DAY(ddate),2) + strzero(MONTH(ddate),2) + SUBSTR(STR(YEAR(ddate), 4), 3, 2)
ENDFUNC
*
FUNCTION DV_ExportFile
CREATE CURSOR Datev (artikel N (4, 0), zahlung N (2, 0), umsatz N (12, 2), gegenkonto C (8), datum D (8), konto C (8), text C (30), beleg1 C (10), beleg2 C (10), mwst C (2), kst C(8))
SELECT datev
INDEX ON DTOS(datum) + STR(artikel, 4) TAG tag1
INDEX ON DTOS(datum) + STR(zahlung, 2) + konto TAG tag2
SET ORDER TO 1
RETURN .T.
ENDFUNC
*
FUNCTION DV_WriteParam
PARAMETER csection, cparameter, cvalue
RETURN .T.
ENDFUNC
*
FUNCTION f_Write
PARAMETER nhandle, cdata
IF  .NOT. EMPTY(cdata)
     cwritebuffer = cwritebuffer + cdata
ELSE
     cwritebuffer = cdata
     nblockcount = nblockcount + 1
ENDIF
RETURN ""
ENDFUNC
*
FUNCTION ChkLicense
PRIVATE ncc, ncount, ctmp, nlic, lret
ncc = 0
ctmp = PADR(param.pa_hotel, 30) + "OsHk991" + PADR(param.pa_city, 30) + PADR("ACCT", 8)
FOR ncount = 1 TO LEN(ctmp)
     ncc = ncc + (ncount * ASC(SUBSTR(ctmp, ncount, 1)))
ENDFOR
nlic = VAL(dv_param("License","Code","0",10))
IF nlic <> ncc
     = alert("Ungültiger Lizenskode für Datev Schnittstelle!;;Bitte tragen Sie den Lizenskode ein in die Datei 'DATEV.INI'.")
     lret = .F.
ELSE
     lret = .T.
ENDIF
RETURN lret
ENDFUNC
*
*** 
*** ReFox - retrace your steps ... 
***
