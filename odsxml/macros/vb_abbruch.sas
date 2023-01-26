
/*
 *  Makro : vb_abbruch    
 *
 *  Funktion :    Makro zum fehlerbedingten Abbruch von Makros.
 *                unter Ausgabe einer Meldung ins SAS-Log und
 *                Setzen einer Rückgabecode-Variablen.
 * 
 *  Bemerkungen : Es wird empfohlen, stattdessen das funktionsgleiche Makro 
 *                    %vb_fehler mit dem Parameter abbruch = 1 zu verwenden!
 *                
 *                Nach Ausgabe der Meldung wird das gerade laufende und alle
 *                umgebenden Makros abgebrochen. Das SAS-System kehrt auf die
 *                Hauptebene zurück und fährt dort mit der Ausführung des
 *                nachfolgenden Codes fort.
 *
 *                Der Fuß kann z. B. den Namen des aufrufenden Makros und
 *                Kontaktinformationen enthalten.
 *
 *                Zeilentrennzeichen in «Meldungstext» und vb_fehler_fuss
 *                ist standardmäßig das Paragraphenzeichen §. Sie können es
 *                durch Belegung der Makro-Variablen vb_meldung_split ändern,
 *                z. B. durch %let vb_meldung_split = "*";
 *
 *                Zur Erzeugung von Leerzeilen ein einfaches Leerzeichen angeben (...§ §...).
 *
 *                Um eventuelle Abbrüche durch %vb_abbruch zu unterdrücken,
 *                können Sie global oder im aufrufenden Kontext die Makro-Variable
 *                vb_vermeide_abbruch auf den Wert 1 setzen.
 *                Um die Ausgabe der von %vb_abbruch, %vb_fehler und >%vb_erfolg
 *                erzeugten Meldungstexte zu unterdrücken, können Sie global oder im
 *                aufrufenden Kontext die Makro-Variable vb_vermeide_meldung auf den Wert 1 setzen.
 *
 *                Dieses Makro benötigt das Makro %vb_ggf_letglobal. 
 *
 *  Einsatz:      Verbund      
 *  Sprache:      ab SAS 9.1.3
 *  Autor:        Jobst Heitzig
 *---------------------------------------------------------------------------------------
 * 
 *  Pflicht-Parameter :
 *    Makroname
 *                Name des abzubrechenden Makros ohne Prozentzeichen %
 *                              Wird benötigt, um die Rückgabecode-Variable Makroname_rc
 *                zu setzen. Dieser Parameter kann auch in der Form «m = Makroname»
 *                übergeben werden und muss dann nicht unbedingt an erster Stelle stehen.
 *    Text        «Meldungstext»    Auszugebender Text. Standard-Zeilentrennzeichen ist §,
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
 *
 *
 *  Globale Variablen:
 *                vb_fehler_fuss      = Fußnote
 *                vb_meldung_split    = Zeichen für Zeilenumbruch
 *                vb_fehlerzahl       = Anzahl Fehler
 *                vb_vermeide_meldung = Fehlermeldung unterdrücken
 *                vb_vermeide_abbruch = Abbruch unterdrücken
 *                vb_rc               = Fehler-Code  
 *                vb_rt               = Fehler-Text
 *---------------------------------------------------------------------------------------
 *
 *  
 *  Änderungen :  September 2008    SAS-Code formatiert    Komenda    
 * 
 */


%macro vb_abbruch(Makroname, text = "Es ist ein Fehler aufgetreten.", rc=1);
    %local rc_var rt_var t notes_bak;

    %let Makroname = %vbf_anfuehrungen_weg(&Makroname);
    %let rc = %vbf_anfuehrungen_weg(&rc);
    %let t = "%vbf_anfuehrungen_weg(%bquote(&text))";

    /* defaults setzen: */
    %vb_ggf_letglobal(vb_fehler_fuss, neu = "")
    %vb_ggf_letglobal(vb_meldung_split, neu = "§")
    /* Name der Rückgabevariablen ermitteln: */

  %if &Makroname ne 
    %then %do;
      %let rc_var = &Makroname._rc;
      %let rt_var = &Makroname._rt;
  %end;

  %else %do;
    %let rc_var = vb_rc;
    %let rt_var = vb_rt;
  %end;

    /* als global deklarieren und setzen: */
    %global &rc_var &rt_var;
    %let &rc_var = &rc;
    %if %eval(%length(&t)-2) > 0 %then 
    %let &rt_var = %substr(&t, 2, %eval(%length(&t) -2));

    /* Fehlerzahl zurücksetzen: */
    %vb_ggf_letglobal(vb_fehlerzahl, wert = 0);

    %vb_ggf_letglobal(vb_vermeide_meldung, leer = 0);
    %vb_ggf_letglobal(vb_vermeide_abbruch, leer = 0);

    %let notes_bak = %sysfunc(getoption(notes));
    options nonotes;

  /* Meldung ausgeben, falls nicht unterdrückt: */
    %if &vb_vermeide_meldung ne 1 %then %do;

    data _null_;
      /* Gesamtmeldung zusammensetzen: */
            %if &vb_fehler_fuss = "" or &vb_fehler_fuss = %then %do;
        text = &t||&vb_meldung_split||"##";
      %end;

      %else %do;
        text = &t  ||&vb_meldung_split||
                         " " ||&vb_meldung_split||
                         &vb_fehler_fuss||&vb_meldung_split||"##";
            %end;

            /* erste Zeile mit ERROR bzw. FEHLER: beginnen: */
            zeile = scan(text, 1, &vb_meldung_split);
            s = repeat(' ', 71-length(zeile));
            put / / / "%vbf_error(:) *** " zeile s;

            /* restliche Zeilen ausgeben: */
            i = 2;
      zeile = scan(text, i, &vb_meldung_split);
      do while (zeile ne "##" and i < 100);
        s = repeat(' ', 71-length(zeile));
                put "%vbf_error(-) *** " zeile s;
                i = i+1;
        zeile = scan(text, i, &vb_meldung_split);
            end;
            text = "";
      zeile = "";
      i = .;
      s = "";
    run;
  %end;

    %if &vb_vermeide_abbruch ne 1 %then %do;
        options noprintmsglist;
        data _null_;
            /* Nun kommt der eigentliche Abbruch unter Ausgabe der Zeile */
            /* ERROR: Execution terminated by an ABORT statement at line ... column ... */
      /* put / "%vbf_warning(-) SAS springt jetzt ins Hauptprogramm zurück";
             put   "%vbf_warning(-) und führt ggf. dort nachfolgende Anweisungen aus!"; */
      put / "%vbf_warning(-) Die nachfolgenden zwei Zeilen können Sie IGNORIEREN:";
            options &notes_bak;
            abort;
        run; 
  %end;
    %else %put;
    options &notes_bak;

  /* 
  ** NOCH ZU TUN: 
  ** WAS PASSIERT IN DEN AUFRUF-SITUATIONEN
  ** TELNET-SESSION, APPLICATION DISPATCHER, SAS/CONNECT-SESSION
  ** ETC. ??? 
  */
%mend;

