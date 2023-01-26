
/*
 *  Makro : vbf_Dataset2Datei    
 *
 *  Funktion :    Gibt den Dateinamen vom Dataset zurück.
 *
 * 
 *  Bemerkungen : Es erfolgt keine Prüfung, ob das
 *                Dataset existiert. 
 *
 *                Fehlt die Angabe des Dateinamens,
 *                so wird Leer zurückgegeben.
 *  
 *  Einsatz:      Verbund      
 *  Sprache:      ab SAS 9.1.3
 *  Autor:
 *---------------------------------------------------------------------------------------
 * 
 *  Parameter :
 *    Dataset  Name der SAS-Datei mit Bibliothek
 *
 *---------------------------------------------------------------------------------------
 *
 *  
 *  Änderungen :  August 2008    SAS-Code formatiert    Komenda    
 *                Mai    2011   Parameter geändert in Dataset 
 * 
 */
%macro vbf_Dataset2Datei(Dataset);
 
  %local datei; 
  %let Dataset = %vbf_anfuehrungen_weg(&Dataset);
   
  %let len = %length(&Dataset);
  %let pos = %index(&Dataset,.);

  %if &len = 0 %then %do;
    /* der Parameter ist leer */
    %let datei = ;
  %end;

  %if &pos = 0 %then %do;
    /* der Parameter enthält keine Library */
    %let datei = &Dataset;
  %end;

  %if &pos = 1 and &len = 1 %then %do;
    /* der Parameter enthält nur "."  */
    %let datei = ;
  %end;

  %if &pos = 1 and &len > 1 %then %do;
    /* der Parameter enthält ".datei" */
    %let datei = %substr(&Dataset, 2);;
  %end;


  %if &pos = &len %then %do;
    /* Der Parameter enthält "Library." */    
    %let datei = ;
  %end;

  %local posp1;
  %local subword;

  %if &pos > 1 and &pos < &len %then %do;
    /* der Parameter enthält Library und Datei */
    %let posp1 = %eval(&pos+1);
    %let subword = %substr(&Dataset,&posp1);

    %if %index(&subword,.) > 0 %then %do;
      /* der Dateiname enthält einen Punkt */
      /* und ist damit ungültig            */
      %let datei = ;
    %end;

    %else %do;
      /* Substring auf den Dateinamen */      
      %let datei = %substr(&Dataset,&posp1);
    %end;
  %end;

  &datei
%mend ;
