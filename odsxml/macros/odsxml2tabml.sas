/*
 *  Makro : OdsXML2TabML    
 *
 *  Funktion :    Das Makro erstellt aus einer in SAS erzeugten OdsXML-Datei
 *                eine TabML-Datei.
 * 
 *  Bemerkungen : 
 *  
 *  Einsatz:      Destatis      
 *  Sprache:      ab SAS 9.1.3
 *  Autor:        Komenda, Edwin
 *---------------------------------------------------------------------------------------
 * 
 *  Pflicht-Parameter :
 *    OdsXMLDatei        Die erzeugte XML-Datei mit kompletten Pfad
 *    TabMLDatei         Ausgabe der TabML-Datei mit kompletten Pfad
 *
 *  Optionale Parameter : 
 *    Spaltenbreiten =   Relative Spaltenbreiten durch / getrennt.
 *                       z.B.: 12/8/8/8/6 für 5 Spalten.
 *                       Wiederholungen können mit * abgekürzt werden.
 *                       Das gleiche Beispiel: 12/3*8/6 .
 *    Zeilenumbruch  =   Anzahl der Zeilen pro Seite.
 *    Spaltenumbruch =   Spaltennummer, die der Trennung einer Seite in
 *                       eine linke und rechte Hälfte dient.
 *                       Alle folgenden Spalten, einschließlich der
 *                       angegebenen Spalte, werden auf die rechte
 *                       Hälfte ausgegeben.
 *    Spaltentrennzeichen = # (default)
 *                       Zeichen für Trennung der Überschriften, 
 *                       Texte des Tabellenkopfes und Fußnoten
 *                       beim Spaltenumbruch.
 *    Ausgabe_Splitten = Nein (default) Ausgabe aller Tabellen in eine Datei
 *                     = Ja   Jede Tabelle seperat in eine Datei ausgeben?
 *    ZwUeberschriften = 0 (default)  Anzahl der Vorspalten,
 *                       die als Zwischenüberschriften ausgegeben werden.
 *                       Es können höchstens (Anzahl der Vorspalten - 1)
 *                       Zwischenüberschriften ausgegeben werden.
 *    ZwUeberschriften_ausrichten = center (default) Horizontale Ausrichtung
 *                                = left
 *                                = right
 *    Dezimalzeichen = Komma (default) als Dezimalzeichen formatiert.
 *                           Einstellung ist nur bei Ausgabe im Excel-Format sinnvoll.
 *                   = Punkt
 *
 *---------------------------------------------------------------------------------------
 *
 *  
 *  Änderungen : Umstellung auf java-Object Mai 2010
 * 
 */

 %macro OdsXML2TabML(OdsXMLDatei,
                     TabMLDatei,
                     TabellenNr = ,
                     Spaltenbreiten = , 
                     Zeilenumbruch = ,
                     Spaltenumbruch =, 
                     Spaltentrennzeichen =,
                     Ausgabe_Splitten =,
                     ZwUeberschriften =,
                     ZwUeberschriften_ausrichten =,
                     Dezimalzeichen = Komma,
                     Log = ,
                     Debug = );

  %local par_XmlDateiDatei;
  %local par_Tabml;
  %local par_Tabnr;
  %local par_SpBr;
  %local par_Zeilenumbr;
  %local par_Spaltenumbr;
  %local par_Trennzeichen;
  %local par_AusgabeSplitten;
  %local laenge;
  %local par_ZwUeb;
  %local par_Ausrichten;
  %local par_log;
  %local fehlertext;
  %local par_Dezimal;
  %local fc;
  %local rgtext;
  %local fehler;

  %let fehler = nein;


  %let par_XmlDatei = %vbf_anfuehrungen_weg(&odsxmldatei);

  %let par_log = nein;
  %let Log = %vbf_anfuehrungen_weg(&Log);
  %let laenge = %length(&log);

  %if &laenge > 0 %then %do;
    %let par_log = %lowcase(&log);
    %if &par_log = ja %then %do;
      %let par_log = Debug;
    %end; 
  %end;

  %let par_Tabml        = %vbf_anfuehrungen_weg(&tabmldatei);
  %let par_SpBr         = %vbf_anfuehrungen_weg(&spaltenbreiten);
  %let par_Zeilenumbr   = %vbf_anfuehrungen_weg(&zeilenumbruch);
  %let par_Spaltenumbr  = %vbf_anfuehrungen_weg(&spaltenumbruch);
  %let par_Trennzeichen = %vbf_anfuehrungen_weg(&Spaltentrennzeichen);

  %if &par_SpBr = %then %let par_SpBr = -;

  /* Wenn TabellenNr fehlt wird "-" eingesetzt,
   * sonst wird der Parameter auf Zahl geprüft
   */
  %let par_Tabnr = %vbf_anfuehrungen_weg(&TabellenNr);
  %if %length(&par_Tabnr) > 0 %then %do;
    %let rgtext = %scan(&par_Tabnr, 1, 0123456789);
    %if %length(&rgtext) > 0 %then %do;
      %let fehlertext = Die TabellenNr enthält keine Zahl.;
      %let fc = 8;
      %goto fehler;
    %end;
  %end;

  %else %do;
     %let par_Tabnr = -;
   %end; 

  %if &par_Zeilenumbr = %then %do;
    %let par_Zeilenumbr = 0;
  %end;

  %if &par_Spaltenumbr = %then %do;
    %let par_Spaltenumbr = 0;
  %end;

  %if %length(&par_Trennzeichen) = 0 %then %do;
    %let par_Trennzeichen = #;
  %end;

  %let par_AusgabeSplitten = nein;
  %let ausgabe_splitten = %vbf_anfuehrungen_weg(&ausgabe_splitten);
  %let laenge = %length(&ausgabe_splitten);

  %if  &laenge ne 0 %then %do;
    %let par_AusgabeSplitten = &ausgabe_splitten;
  %end;

  %let par_ZwUeb = 0;
  %let par_Ausrichten = center;

  %let zwueberschriften = %vbf_anfuehrungen_weg(&zwueberschriften);
  %let laenge = %length(&zwueberschriften);

  %if &laenge > 0 %then %do;
    %let par_ZwUeb = &zwueberschriften;
    %if &laenge > 1 or &par_ZwUeb < 0 or &par_ZwUeb > 9 %then %do;
      %let fehlertext = Ungültige Anzahl der Zwischenüberschriften.;
      %let fc = 8;
      %goto fehler;
    %end;
  %end;

  %let zwueberschriften_ausrichten = %vbf_anfuehrungen_weg(&zwueberschriften_ausrichten);
  %let laenge = %length(&zwueberschriften_ausrichten);

  %if &laenge ne 0 %then %do;
    %let par_Ausrichten = %lowcase(&zwueberschriften_ausrichten);
     %if &par_Ausrichten ne center and &par_Ausrichten ne left and &par_Ausrichten ne right %then %do;
      %let fehlertext = Ungültige Ausrichtung der Zwischenüberschriften.;
      %let fc = 8;
      %goto fehler;
    %end;
  %end;



  %let par_Dezimal = %vbf_anfuehrungen_weg(&Dezimalzeichen);
  %let par_Dezimal = %lowcase(&par_Dezimal);

  %if &par_Dezimal = punkt %then %do;
    %let par_Dezimal = .;
  %end;
  %else %do;
    %let par_Dezimal = ,;
  %end;



 /*
  * Java Aufrufrufe
  */

  data _null_;
    length rgw 8
           rgFehlertext $200;
    


    /*
     * Anlegen Objekt von Java
     */
    declare javaobj tabml("de/destatis/sas/odsxml/OdsXml2TabML");

    /*
     * Setzen der Parameter
     */
    %local path_sasuser;
    %let path_sasuser = %sysfunc(pathname(sasuser));

    tabml.callVoidMethod("setLogEinstellungen", "&par_log",
                         "&path_sasuser/sasodsxml.log", "nein");

    tabml.callVoidMethod("setXmlDatei", "&par_XmlDatei");
    tabml.callVoidMethod("setTabMLDatei", "&par_Tabml");
    tabml.callVoidMethod("setSpaltenBreiten", "&par_SpBr");
    tabml.callVoidMethod("setSpaltenUmbruchZahl", "&par_Spaltenumbr");
    tabml.callVoidMethod("setZeilenUmbruchZahl", "&par_ZeilenUmbr");
    tabml.callVoidMethod("setSpaltenUmbruchTrennzeichen", "&par_Trennzeichen");
    tabml.callVoidMethod("setAusgabeSplitten", "&par_AusgabeSplitten");
    tabml.callVoidMethod("setAnzahlSpaltenZwUeb", "&par_ZwUeb");
    tabml.callVoidMethod("setAusrichtungZwUeb", "&par_Ausrichten");
    tabml.callVoidMethod("setSasDezimalzeichen", "&par_Dezimal");
    tabml.callVoidMethod("setTabellenNr", "&par_Tabnr");

    /*
     * Aufruf Verarbeitung
     */
    tabml.callIntMethod("executeOdsXml", rgw);
 
    if rgw ne 0 then do;
      call symput("fehler", "ja");
      fehlerText = tabml.callStringMethod("getFehlertext", rgFehlertext);


      tabml.delete();
      %vb_data_fehler(OdsXML2TabML, Text = rgFehlertext,
                      rc = 1, abbruch = 1, stop = 1);
    end;
    
    tabml.delete();
  run;

  /*
   * Nach Java-Aufrufe
   */
  %if &fehler = ja %then %do;
    %vb_ggf_fehler;
    %goto finish;
  %end;

  %vb_erfolg(OdsXML2TabML ,text = "OdsXML2TabML ohne Fehler ausgeführt.§Die Ausgabe: &par_Tabml wurde erstellt.");
  %goto finish;

  %fehler:
     %vb_abbruch(OdsXML2TabML, text = "&fehlertext", rc = "&fc");

  %finish:  
 %mend;
