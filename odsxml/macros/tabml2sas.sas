
/*
 *  Makro : Tabml2SAS    
 *
 *  Funktion :    Das Makro erstellt aus einer vorliegenden TabML-Datei eine SAS-Datei.
 *
 * 
 *  Bemerkungen : Die TabML-Datei kann z.B.: das Ergebnis einer Tabellierung
 *                in Statspez oder eine exportierte Ergebnistabelle aus Genesis sein.
 *  
 *  Einsatz:      Destatis      
 *  Sprache:      ab SAS 9.1.3
 *  Autor:        Komenda, Edwin
 *---------------------------------------------------------------------------------------
 * 
 *  Pflicht-Parameter :
 *    Eingabe   TabML-Datei mit kompletten Pfad.
 *    SASDatei  Name der SAS-Datei (in der Form  Bibliothek.Datei).
 *
 *  Optionale Parameter : 
 *    Dezimalzeichen  = Komma  Als Dezimalzeichen wird das Komma (',') verwendet.                                     als Dezimalzeichen formatiert,
 *                      Punkt  Als Dezimalzeichen wird der Punkt ('.') verwendet.
 *                   (default) Wurde die TabML-Datei mit %OdsXML2TabML erstellt,
 *                             ist das Dezimalzeichen Punkt (in SAS standard),
 *                             sonst Komma.
 *    Split = Minus (default)  Zeilenumbruch für die erzeugten Labels   
 *            Stern       * 
 *            Minus       -
 *            Plus        +
 *            At-Zeichen  @
 *            Und         &
 *---------------------------------------------------------------------------------------
 *
 *  
 *  Änderungen : Umstellung auf java-Object Mai 2010
 * 
 */


%macro Tabml2SAS(Eingabe,
                 SASDatei,
                 Dezimalzeichen =,
                 Split =,
                 Log =);
  %local verzeichnis_dataset_pfad;
  %local verzeichnis_dataset_lib;
  %local verzeichnis_dataset_dataset;
  %local par_Eingabe;
  %local par_Sasdatei;
  %local par_Dezimalzeichen;
  %local par_Umbruch;
  %local par_Log;
  %local laenge;
  
  %local anzahlDataSteps;

  %local fehler;
  %let fehler = nein;



  %*stba_aufruf_zaehlen("TabML2SAS");

  %let par_Eingabe = %vbf_anfuehrungen_weg(&eingabe);
  %let par_Sasdatei = %vbf_anfuehrungen_weg(&sasdatei);
  %let par_Dezimalzeichen = %vbf_anfuehrungen_weg(&Dezimalzeichen);

  %let laenge = %length(&par_Dezimalzeichen);
  %if &laenge = 0 %then %do; 
    %let par_Dezimalzeichen = -;
  %end;


  %let laenge = %length(&split);
    %if &laenge = 0 %Then %do;
    %let par_Umbruch = minus;
  %End;
  %Else %do;
    %let par_Umbruch = %vbf_anfuehrungen_weg(&split);
  %End;

  %let par_Log = nein;
  %let Log = %vbf_anfuehrungen_weg(&Log);

  %let laenge = %length(&log);
  %if &laenge > 0 %then %do;
    %let par_Log = &LOG;
  %end; 


  /*
   * Pfadanweisung für Libanweisung ermitteln
   * Parameter SASDatei wird getrennt in Library und Dataset
   */
  %let verzeichnis_dataset_pfad = %vbf_dataset2ordner(&sasdatei);
  %let verzeichnis_dataset_lib  = %vbf_dataset2bibliothek(&sasdatei);
  %let verzeichnis_dataset_dataset = %vbf_dataset2datei(&sasdatei);

  %if %length(&verzeichnis_dataset_pfad) = 0 %Then %do;
    %vb_abbruch(Tabml2SAS, text = "Die Library ist nicht vorhanden.", rc = 1);
    %goto finish;
  %End;



  data _null_;
    length rgw 8
           fehlerText $50
           datastep $32767
           anzahl 8;
    


    /*
     * Anlegen Objekt von Java
     */
    declare javaobj tabml("de/destatis/sas/tabml/TabML2Sas");

    /*
     * Setzen der Parameter
     */
    %local path_sasuser;
    %let path_sasuser = %sysfunc(pathname(sasuser));

    tabml.callVoidMethod("setLogEinstellungen", "&par_Log",
                         "&path_sasuser/sastabml.log", "nein"); 


    tabml.callVoidMethod("setTabMLDatei", cats("&par_Eingabe"));
    tabml.callVoidMethod("setDataSet", "&par_Sasdatei");
    tabml.callVoidMethod("setLabelUmbruchZeichen", "&par_Umbruch");
    tabml.callVoidMethod("setDezimalZeichen", "&par_Dezimalzeichen");

    tabml.callIntMethod("executeTabML", rgw);

    if rgw ne 0 then do;
      call symput("fehler", "ja");
      fehlerText = tabml.callStringMethod("getFehlertext", fehlerText);
      tabml.delete();
      %vb_data_fehler("Tabml2SAS", Text = fehlerText,
                      rc = 1, abbruch = 1, stop = 1);
    end;

    tabml.callIntMethod("getAnzahlTabellen", anzahl);
    call symput("anzahlDataSteps", anzahl);
    
    %local i;

    do i = 1 to anzahl;
      tabml.callStringMethod("getSasCode", i, datastep);
      call symput("datastep" || left(putn(i, "3.0")), datastep);
    end;


    tabml.delete();
  run;

  /*
   * Nach Java-Aufrufe
   */
  %if &fehler = ja %then %do;
    %vb_ggf_fehler;
  %end;

  /*
   * Aufruf der DataSteps
   */
  %do i = 1 %to &anzahlDataSteps;
    &&datastep&i;
  %end;

 %finish: ; 
 %mend;
