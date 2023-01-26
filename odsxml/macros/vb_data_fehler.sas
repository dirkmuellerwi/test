
/*
 *  Makro : vb_data_fehler    
 *
 *  Funktion :    Dieses Makro dient zusammen mit vb_ggf_fehler dazu, 
 *                nach einem fehlerhaften Data-Schritt eine Fehlermeldung
 *                auszugeben.
 * 
 *  Bemerkungen : Dies ist notwendig, da das Makro vb_fehler nicht innerhalb
 *                eines Data- oder Proc-IML-Schritts eingesetzt werden kann!
 *
 *                Anwendung:
 *                data ...;
 *                   Anweisung;
 *                   .
 *                   .
 *                   if ... then %vb_data_fehler(«Parameter»);
 *                   .
 *                   .
 *                   .                    <br>
 *                 run;                    <br>
 *                 %vb_ggf_fehler
 *
 *                Dabei sind dieselben Parameter wie bei vb_fehler</code></a> erlaubt
 *                und nur die Angabe von «Makroname» zwingend.
 *                Gibt man abbruch = 1 an, wird hier allerdings erst bei Erreichen der
 *                Zeile vb_ggf_fehler abgebrochen!
 *                Ergänzend oder alternativ zu abbruch = 1 kann man über den zusätzlichen optionalen 
 *                Parameter stop = 1 erreichen, dass der laufende Data-Schritt sofort  gestoppt wird.
 *                Achtung: Beim Erreichen der Zeile vb_ggf_fehler  wird immer nur die letzte durch
 *                vb_data_fehler erzeugte Meldung ausgegeben!
 *                Funktionsweise: vb_data_fehler speichert die Fehlermeldung in einigen globalen
 *                Makrovariablen, die von vb_ggf_fehler ausgelesen und ggf. zurückgesetzt werden.
 *
 *  Einsatz:      Verbund      
 *  Sprache:      ab SAS 9.1.3
 *  Autor;        Jobst Heitzig
 *
 *---------------------------------------------------------------------------------------
 *  Pflicht-Parameter :
 *    Makroname
 *                Name des abzubrechenden Makros ohne Prozentzeichen %
 *                Wird benötigt, um die Rückgabecode-Variable Makroname_rc
 *                zu setzen. Dieser Parameter kann auch in der Form «m = Makroname»
 *                übergeben werden und muss dann nicht unbedingt an erster Stelle stehen.
 *    Text        «Meldungstext»  Auszugebender Text. Standard-Zeilentrennzeichen ist §,
 *                dieses kann aber durch Setzen der globalen Variablen vb_meldung_split
 *                angepasst werden.
 *    Optionale Parameter :
 *    rc  =       Rückgabecode.
 *                Dieser wird in die globale Makro-Variable Makroname»_rc geschrieben
 *                und kann vom aufrufenden Makro ausgelesen werden.
 *
 *                Um zusätzlich zu «Meldungstext» einen für alle Fehlermeldungen
 *                einheitlichen <b>Fußtext</b> ins SAS-Log auszugeben, setzen Sie
 *                global oder im aufrufenden Kontext die Makro-Variable vb_fehler_fuss,
 *                etwa durch   %let vb_fehler_fuss = "«einheitlicher Fußtext»";
 *    abbruch =    Soll das Programm nach dem Data-Schritt abgebrochen werden?
 *    abbruch = 1  abbrechen
 *    abbruch = 0 (de4fault) nicht abbrechen
 *
 *    stop =      Soll der Data-Schritt hier gestoppt werden?
 *    stop = 1    stoppen
 *    stop = 0    (default) nicht stoppen
 *
 *
 *  Autor :   Jobst Heitzig        Mai 2003
*---------------------------------------------------------------------------------------
  *  Änderungen :  Oktober 2008    SAS-Code formatiert    Komenda    
 * 
 */



%macro vb_data_fehler(Makroname, text = "", rc = 1, stop = 0, abbruch = 0);
  /* Wichtig: alles in einen Block einschließen, da evtl. nach then oder else: */
  do;
    %let Makroname = %vbf_anfuehrungen_weg(&Makroname);
    %let rc = %vbf_anfuehrungen_weg(&rc);

    /* Werte für späteren Aufruf von %vb_ggf_fehler setzen: */
    call symput('vb_ggf_fehler_aufgetreten', "1");

    /* hier Anführungsstriche, da Makroname = ... ohne angegeben wird: */
    call symput('vb_ggf_fehler_m', "&Makroname");

    /* hier keine, da text=... auch ein Datastep-Ausdruck sein kann! : */
    call symput('vb_ggf_fehler_text', &text);
    /* hier aus demselben Grund auch keine: */ 
    %if &sysprocname=DATASTEP %then %do;
      call symput('vb_ggf_fehler_rc', trim(put(&rc, 10.0)));
    %end;
    %else %do;
      call symput('vb_ggf_fehler_rc', char(&rc, 10, 0));
    %end;

    /* späteren Abbruch anfordern: */
    if &abbruch=1 then do; 
      call symput('vb_ggf_fehler_abbruch',"1");
      stop;
    end;
    else do;
      call symput('vb_ggf_fehler_abbruch',"0");
    if &stop = 1 then stop;
    end;
  end;
%mend;
