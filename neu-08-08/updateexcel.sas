/*
 *  Makro : UpdateExcel   
 *
 *  Funktion :  Eine bereits vorhandene Excel-Datei wird
 *              durch eine oder mehrere SAS-Dateien aktualisiert.
 * 
 *  Bemerkungen : Aufruf der Makros: %UpdateExcel_Zusammenfassen
 *                                   %UpdateExcel_ConvertMissing 
 *               
 *------------------------------------------------------------------------
 * 
 *  Pflicht-Parameter :
 *   SASDatei              = SASDatei(en)
 *   ExcelDatei            = Pfad Excel-Datei (Ergebnisdatei)
 *   Arbeitsblatt          = Eine oder mehrere Arbeitsblätter
 *   Adresse               = Eine oder mehrere Zelladressen     
 *
 *  Optionale Parameter : 
 *   Zeitreihen        = Nein (default) Ja = Erste Wertspalte ist ein Datumsformat
 *   Trim              = Nein (default) Ja Links Rechts
 *   ExcelVorlage      = Pfadname der Exceldatei (Vorlage wird in die Ergebnisdatei ausgegeben)
 *   FormelnNeuUpdate  = Nein (Default) 
 *                                 Ja = Die Formeln werden nach dem Update aktualisiert
 *   FormelnNeuOeffnen = Ja Sorgt dafür, dass nach dem Öffnen in Excel
 *                                    die Formeln aktualisiert werden.
 *------------------------------------------------------------------------
 *
 *  Autor : Komenda, Edwin
 *  
 *  Änderungen : Frenzel, Stephan
 *
 *    Vgl. Kommentare mit "SFrenzel"
 * 
 */

%macro  UpdateExcel(SASDatei          = ,
                    ExcelDatei        = ,
                    Arbeitsblatt      = ,
                    Adresse           = ,
                    Trim              = Nein,
                    ExcelVorlage      = ,
                    Zeitreihen        = Nein,
                    FormelnNeuUpdate  = Nein,
                    FormelnNeuOeffnen = ,
                    Log               = Info,
                    Version           = 1,
					LimitZeilen       = ,
					LimitSpalten      = ,
					LimitZellen       = ,
					LimitGroesse      = 
);

  /* Der Zugriff auf Umgebungsvariablen mit "sysfunc(sysget(...))" fuehrt zu einer Warning,
     falls Umgebungsvariablen nicht gesetzt sind; daher erfolgt der Zugriff hier im Datastep.
  */
  data _null_;
    if "&LimitZeilen." = "" then call symput ('LimitZeilen', sysget('UEROWCOUNTLIMIT'));
    if "&LimitSpalten." = "" then call symput ('LimitSpalten', sysget('UECOLUMNCOUNTLIMIT'));
    if "&LimitZellen." = "" then call symput ('LimitZellen', sysget('UECELLCOUNTLIMIT'));
    if "&LimitGroesse." = "" then call symput ('LimitGroesse', sysget('UEFILESIZELIMIT'));
  run;
  %put UpdateExcel LimitZeilen = &LimitZeilen.;
  %put UpdateExcel LimitSpalten = &LimitSpalten.;
  %put UpdateExcel LimitZellen = &LimitZellen.;
  %put UpdateExcel LimitGroesse = &LimitGroesse.;

  %local macroName;
  
  %if %UpCase(&Zeitreihen) = JA %then %do;
    %let macroName = UpdateExcelZeitreihen;
  %end;
  %else %do;
    %let macroName = UpdateExcel;
  %end;
 
  /*
   * Aufruf zählen
   */
  /*  %stba_aufruf_zaehlen("&macroname");  */

  %local bak_notes;
  %let bak_notes = %sysfunc(getoption(NOTES));

  %local fehler_text fehler_rc;
  option NOQUOTELENMAX;
  option NONOTES;

  /*
   * Evtl. uebergebene Anfuehrungszeichen entfernen.
   */
  %let SASDatei          = %vbf_anfuehrungen_weg(&SASDatei);
  %let ExcelDatei        = %vbf_anfuehrungen_weg(&ExcelDatei);
  %let Trim              = %vbf_anfuehrungen_weg(&Trim);
  %let FormelnNeuUpdate  = %vbf_anfuehrungen_weg(&FormelnNeuUpdate);
  %let FormelnNeuOeffnen = %vbf_anfuehrungen_weg(&FormelnNeuOeffnen);

  /*
   * Trim Optionen
   */
  %if %length(&Trim) = 0 %then
    %let Trim = nein;
  %else %if  &Trim = 1  or  %upcase(&Trim) = JA  %then 
    %let Trim = ja;
  %else %if  &Trim = 0  or  %upcase(&Trim) = NEIN  %then 
    %let Trim = nein;
  %else %if  %upcase(&Trim) = LINKS  or  %upcase(&Trim) = LEFT  %then 
    %let Trim = links;
  %else %if  %upcase(&Trim) = RECHTS  or  %upcase(&Trim) = RIGHT  %then 
    %let Trim = rechts;
  %else %do;
    %let text = Für den Parameter 'Trim' sind nur folgende Werte zulässig:;
    %let text = &text §   JA   (oder 1)  =  Bei alphanumerischen Werten werden führende Blanks bzw. Blanks am Ende entfernt.;
    %let text = &text §   NEIN (oder 0)  =  Alphanumerische Werte werden unverändert übernommen.;
    %let text = &text §   LINKS          =  Nur führende Blanks werden entfernt.;
    %let text = &text §   RECHTS         =  Nur Blanks am Ende werden entfernt.;

    option QUOTELENMAX;
    option &bak_notes;
    %vb_abbruch(&MacroName,
                rc   = 3,
                text = "&text");
    %goto exit;
  %end;

  %let ExcelVorlage      = %vbf_anfuehrungen_weg(&ExcelVorlage);
  %let Zeitreihen        = %vbf_anfuehrungen_weg(&Zeitreihen);

  %let Log               = %vbf_anfuehrungen_weg(&Log);

  /*
   * In den Parametern Arbeitsblatt und Adresse
   * die " " und ' '  entfernen
   */
  data _null_;
    length mit_hochkomma ohne_hochkomma $32767;

    mit_hochkomma   = symget("Arbeitsblatt");
    ohne_hochkomma  = translate(mit_hochkomma, ' ', '"');
    ohne_hochkomma  = translate(ohne_hochkomma, " ", "'");
    call symput("Arbeitsblatt", trim(left(ohne_hochkomma)));

    mit_hochkomma   = symget("Adresse");
    ohne_hochkomma  = translate(mit_hochkomma, ' ', '"');
    ohne_hochkomma  = translate(ohne_hochkomma, " ", "'");
    call symput("Adresse", trim(left(ohne_hochkomma)));
  run; 

  /*
   * Sind alle SASDateien vorhanden?
   */
  %local eintragSASDatei indx anzahl;

  %let anzahl = 0;
  %let indx = 1;

  %let eintragSASDatei = %scan(&SASDatei, &indx, %str( ));
  %do %while (%length(&eintragSASDatei) > 0);
    %if %vbf_dataset_exist(&eintragSASDatei) = 0 %then %do;
      option QUOTELENMAX;
      option &bak_notes;

      %vb_abbruch(&macroName,
                  rc   = 1,
                  text = "Die SAS-Datei &eintragSASDatei existiert nicht.");
      %goto exit;
    %end;

    %let anzahl = %eval(&anzahl + 1);
    %let indx = %eval(&indx + 1);
    %let eintragSASDatei = %scan(&SASDatei, &indx, %str( ));
  %end;

  %if &anzahl = 0 %then %do;
    option QUOTELENMAX;
    option &bak_notes;

    %vb_abbruch(&macroName,
                rc   = 1,
                text = "Es wurde keine SAS-Datei angegeben.");
    %goto exit;
  %end;

  /* SFrenzel: Die folgenden beiden Zeilen hier nach oben gezogen, s.u. */
  %local returnCode fehlerText;
  %let returnCode = 0;

  /* SFrenzel: Falls möglich neue Funktionalität "RangeCopier" aus POI-5 nutzen */
  %if %str (&Version.) >= 2 and
    %sysfunc (countw (&SASDatei., %str( ))) = 1 and
    %upcase(&Trim.) = NEIN and
    %upcase(&Zeitreihen.) = NEIN and
    %upcase(&FormelnNeuUpdate.) = NEIN and
    %length(&FormelnNeuOeffnen.) = 0
    %then %do; %* SAS-9 / POI-5-Modus;

    %local _SASDatei;
    %let _SASDatei = %scan (&SASDatei., 2, .);
    %if %length (&_SASDatei.) = 0 %then %let _SASDatei = &SASDatei.;

	%* Um eine SAS-Datei direkt in eine Excel-Datei zu schreiben ist die unter 
	   Windows und Unix gleichermassen verfügbare Methode der XLSX-Engine;

	%if %length (&ExcelVorlage.) = 0 %then %do;

      libname _UpdXLSX xlsx "&ExcelDatei.";

      data _UpdXLSX.&_SASDatei.;
	    set &SASDatei.;
	  run;

	  libname _UpdXLSX clear;

    %end; %* Einfaches Schreiben von SAS nach Excel ohne Vorlage;

	%else %do; %* Schreiben von SAS in eine Excel-Vorlage;

	  %local _NVar _NObs _SysTime;
	  data _null_;
	    /* Bestimmung Anzahl Zeilen und Spalten der Quelldatei */
	    dsid = open ("&SASDatei.");
	    call symput ("_NVar", strip (put (attrn (dsid,'NVAR'), best.)));
	    call symput ("_NObs", strip (put (attrn (dsid,'NObs'), best.)));
		rc = close (dsid);
		/* Systemzeit numerisch zur Bildung eines eindeutigen Namens der Datei
		   zur Übergabe von SAS-Daten im Excel-Format
		*/
	    call symput ("_SysTime", strip (put (round(datetime()), best.)));
	  run;
	  %put _NVar = &_NVar., _NObs = &_NObs., _SysTime = &_SysTime.;

	  libname _UpdXLSX xlsx "&ExcelDatei..&_SysTime..xlsx";

      data _UpdXLSX.&_SASDatei.;	
	    set &SASDatei.;
	  run;

	  libname _UpdXLSX clear;

      %CopyRange (
        /* Quelle */
        ExcelDatei			= %str (&ExcelVorlage.),
        ExcelDateiData		= %str (&ExcelDatei..&_SysTime..xlsx),
        Arbeitsblatt		= &_SASDatei.,				/* Arbeitsplatt, das gelesen werden soll */
        Zelle				= A2,						/* Zelle, ab der gelesen werden soll */
        AnzahlZeilen		= &_NObs.,
        AnzahlSpalten		= &_NVar.,
        /* Ziel */
        ExcelDateiZiel		= %str (&ExcelDatei.),
        ArbeitsblattZiel	= &Arbeitsblatt.,			/* Name des Blattes, in das geschrieben werden soll */
        ZelleZiel			= &Adresse.,				/* Zelle, ab der geschrieben werden soll */
	    Log  				= &Log.,
		LimitZeilen			= &LimitZeilen.,
		LimitSpalten		= &LimitSpalten.,
		LimitZellen			= &LimitZellen.,
		LimitGroesse		= &LimitGroesse.
      );

      %if %sysfunc(fileexist(&ExcelDatei..&_SysTime..xlsx)) ge 1 %then %do;
        %let rc=%sysfunc(filename(temp,&ExcelDatei..&_SysTime..xlsx)); %put &rc.;
        %let rc=%sysfunc(fdelete(&temp)); %put &rc.;
      %end; 
      %else %put Die Datei &ExcelDatei..&_SysTime..xlsx existierte nicht.;

    %end; %* Schreiben von SAS in eine Excel-Vorlage;

  %end; %* POI5-Modus 22.05;

  %else %do; %* Kompatibilitätsmodus 21.12;

  /*
   *  Alle SAS-Dateien zu einer Gesamt-SAS-Datei zusammenfassen
   */
  %local Zusammenfassung;
  %let Zusammenfassung = _UpdateExcel_Zusammenfassung_;

  %local NSASDateien  NVarsInsgesamt  NObsInsgesamt;
  %let NSASDateien    = ;
  %let NVarsInsgesamt = ;
  %let NObsInsgesamt  = ;

  %UpdateExcel_Zusammenfassen(&SASDatei, &Zusammenfassung);
  
  /* SFrenzel: Ausgabe der Struktur der Zusammenfassungsdatei */
  %if %upcase(&Log) = DEBUG %then %do;
    proc contents data = _UpdateExcel_Zusammenfassung_;
    run;
  %end;

  /*
   * Jetzt nur noch mit der Gesamt-Datei arbeiten
   * Ermitteln: Anzahl: Zeilen, Variablen, Variablen-Typen
   *
   */
  %local AnzahlZeilen;
  %let AnzahlZeilen = %vbf_dataset_attrn(&Zusammenfassung, NOBS); 

  %local  AnzahlVariablen;
  %let AnzahlVariablen = %vbf_dataset_attrn(&Zusammenfassung, NVAR); 

  %local VarTypen;
  %let VarTypen = %vbf_DatasetVarTypen(&Zusammenfassung);

  %local withNumeric  withCharacter ;
  %let withNumeric   = %eval( %index(&VarTypen, N) > 0 );
  %let withCharacter = %eval( %index(&VarTypen, C) > 0 );

  /*
   * Parameter Zeitreihe
   */
  %local isZeitreihe;
  %let isZeitreihe = 0;

  %if %upcase(&Zeitreihen) = JA %then %do;
    %let isZeitreihe = 1;
  %end;

  /* SFRenzel: Die folgenden beiden Zeilen nach oben */
  %*local returnCode fehlerText;
  %*let returnCode = 0;

  %local path_sasuser;
  %let path_sasuser = %sysfunc(pathname(sasuser));

  /*
   * Einstellung Formeln neu berechnen
   */
  %if %upcase(&FormelnNeuUpdate) = JA %then %let FormelnNeuUpdate = 1;
  %else %let FormelnNeuUpdate = 0;

  %if %length(&FormelnNeuOeffnen) = 0 %then %do;
    %if &FormelnNeuUpdate = 0 %then %do;
      %let FormelnNeuOeffnen = JA;
    %end; 
    %else %do;
      %let FormelnNeuOeffnen = Nein;
    %end; 
  %end;

  %if %upcase(&FormelnNeuOeffnen) = JA %then %let FormelnNeuOeffnen = 1;
  %else %let FormelnNeuOeffnen = 0;

   /*
   * Datastep mit Aufruf der Java-Methoden
   */ 

  data _null_;
  
    set  &Zusammenfassung end = istLetzterSatz;

    /*
     *  ACHTUNG:  array - Anweisungen müssen zuerst angegeben werden,
     *            sonst werden die zusätzlichen Variablen auch mit aufgenommen !
     */
    %if  &withNumeric  %then %do;
      array __numArray {*} _NUMERIC_;
    %end;

    %if  &withCharacter  %then %do;
      array __charArray {*} _CHARACTER_;
    %end;

	/* SFrenzel, 2022-04-13:
	   __NobsInsgesamt und __NObsInsgesamt reduziert von $32767 auf $8. SAS-
	   Dateien koennen max. 32767 Variablen haben. Um die Zahl 32767 als 
	   String abzulegen, reichen 5 Zeichen.
	   Desgleichen __Arbeitsblatt, __Adresse: Reduziert von $32767 auf $255.
	   __VarTypen bleibt $32767 - enthaelt Typen (C/N) aller Variablen.
	   Lesen aus Makro-Variablen und Uebergabe an Java erfolgt ausserdem mit
	   strip(), um Auffuellen mit Blanks zu vermeiden.
	*/
	/* OPTIMIZE: __NobsInsgesamt und __NObsInsgesamt (auch in Java) als
	   numerische Variablen anlegen - wie __NSASDateien.
	*/
    length __NSASDateien                         8;
    length __NVarsInsgesamt __NObsInsgesamt      $8;
    length __VarTypen                            $32767;
    length __ExcelErgebnisDatei __ExcelVorlage   $255;
    length __Arbeitsblatt                        $255;
    length __Adresse                             $255;
    length __isZeitreihe                         8;
    length __TrimOption                          $40;
    length __FormelnNeuUpdate                    8;
    length __FormelnNeuOeffnen                   8;
    length __indexSasDatei                       8
    length __fehlerText                          $500;
    length __returnCode                          8;
    __returnCode = 0;

    if _N_ = 1 then do;
	
      /*
      * Anlegen Java-Objekt und Setzen der Parameter
      */
      declare javaobj excel("de/destatis/sas/excel/UpdateExcel");

      __NSASDateien        = &NSASDateien;
      __NVarsInsgesamt     = strip(symget("NVarsInsgesamt"));
      __NObsInsgesamt      = strip(symget("NObsInsgesamt"));
      __VarTypen           = strip(symget("VarTypen"));

      __ExcelErgebnisDatei = strip("&ExcelDatei"); 
      __ExcelVorlage       = strip("&ExcelVorlage");
      __Arbeitsblatt       = strip(symget("Arbeitsblatt"));
      __Adresse            = strip(symget("Adresse"));
      __TrimOption         = strip("&Trim");
      __isZeitreihe        = &isZeitreihe;

      excel.callVoidMethod("setLogEinstellungen", "&log", "&path_sasuser/updateexcel.log", "nein");
   
      excel.callVoidMethod("setNSASDateien", __NSASDateien);
      excel.callVoidMethod("setNVarsInsgesamt", strip(__NVarsInsgesamt));
      excel.callVoidMethod("setNObsInsgesamt", strip(__NObsInsgesamt));
      excel.callVoidMethod("setVarTypen", strip(__VarTypen));
      excel.callVoidMethod("setExcelErgebnisDatei", strip(__ExcelErgebnisDatei));
      excel.callVoidMethod("setExcelVorlageDatei", strip(__ExcelVorlage));
      excel.callIntMethod("setArbeitsblatt", strip(__Arbeitsblatt), __returnCode);

      if __returnCode = 0 then do;
	  
        excel.callIntMethod("setAdresse", __Adresse, __returnCode);
        excel.callVoidMethod("setTrimOption", __TrimOption);
        excel.callVoidMethod("setIsZeitreihe", __isZeitreihe);

        __FormelnNeuUpdate = &FormelnNeuUpdate;

        excel.callVoidMethod("setAktualisierenFormeln", __FormelnNeuUpdate);
  
        __FormelnNeuOeffnen = &FormelnNeuOeffnen;
        excel.callVoidMethod("setSchalterAktualisieren", __FormelnNeuOeffnen);
		
      end;
	  
    end; /* if _N_ = 1  */

    /*
     *  Datensätze übergeben
     */
    if __returnCode = 0 then do;

      %if  &withNumeric  %then %do;
        %UpdateExcel_ConvertMissing(__numArray);
        excel.callVoidMethod("setNumericData", __numArray);
      %end;
      %if  &withCharacter  %then %do;
        excel.callVoidMethod("setCharacterData", __charArray);
      %end;
	  
    end;
    
    /*
     * Nach letztem Satz - Excel-Datei speichern
     */
    if  istLetzterSatz then do;
	
      if __returnCode = 0 then do; 
        excel.callIntMethod("update", __returnCode);
      end;

      if __returnCode > 0 then do;
        /*
         * Fehlertext holen
         */
        excel.callStringMethod ("getFehlertext", __fehlerText);
        call symput("fehlerText", trim(__fehlerText));
        call symput("returnCode", putn(__returnCode, "2.0"));
        stop;
      end;
	  
    end; /* Nach letztem Satz */
	
  run; /* Ende Datastep */

  /*
   *  Zusammenfassungs-Datei löschen
   */

  proc datasets library = work nolist;
    delete  &Zusammenfassung ;
  run;

  %end; %* Kompatibilitätsmodus 21.12;

  /*
   *  Ist Fehler aufgetreten?
   */
  %if  &returnCode = 0 %then %do;
    /*
     * OK!
     */
    option QUOTELENMAX;
    option &bak_notes;
    %vb_erfolg(&MacroName, 
                  text="Die Excel-Datei &ExcelDatei §wurde durch die SAS-Datei(en) § &SASDatei aktualisiert.")
  %end;

  %else %do;
    option QUOTELENMAX;
    option &bak_notes;
    %vb_abbruch(&MacroName,
                  rc    = &returnCode,
                  text  = "&fehlerText");
  %end;

%exit:  

%mend;
