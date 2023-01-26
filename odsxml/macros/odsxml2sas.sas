
/*
 *  Makro : OdsXml2Sas    
 *
 *  Funktion :    Liest ODS-XML-Output in eine (oder evtl. mehrere) SAS-Dateien ein.
 *
 * 
 *  Bemerkungen : Dieses Makro kann insbesondere dazu verwendet werden,
 *                die Ergebnisse eines Proc Tabulate Aufrufs in eine
 *                SAS-Datei auszugeben und anschließend weiter zu verarbeiten.
 *  
 *  Einsatz:      Destatis      
 *  Sprache:      ab SAS 9.1.3
 *  Autor:        Komenda, Edwin
 *---------------------------------------------------------------------------------------
 * 
 *  Pflicht-Parameter :
 *    OdsXMLDatei   Name der ODS-XML-Datei mit kompletten Pfad
 *    SASDatei  Name der SAS-Datei (in der Form  Bibliothek.Datei)
 *
 *  Optionale Parameter : 
 *    Dezimalzeichen  = Punkt (default) Die auszugebenden Werte in SAS sind mit                                     als Dezimalzeichen formatiert,
 *                                      diesem Dezimalzeichen formatiert.
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


%macro OdsXml2Sas(OdsXMLDatei,
                  SASDatei,
                  Dezimalzeichen =,
                  Split =,
                  Log =);
  %local TabMLDatei;
  %let TabMLDatei = &home/tmp/_temp_.tabml;

  %*stba_aufruf_zaehlen("OdsXML2SAS");

  %odsxml2tabml(&OdsXMLDatei, "&TabMLDatei", Log = &Log);
  %tabml2sas(&TabMLDatei, &SASDatei, Dezimalzeichen = &Dezimalzeichen, Split = &Split, Log = &Log);
  
  
  /*
   *   temporäre Datei löschen
   */
  %local fref rc;  
  %let fref = tempfile;
  %let rc = %sysfunc(filename(fref, "&TabMLDatei"));
  %let rc = %sysfunc(fdelete(&fref)); 
 
 %exit:

%mend;