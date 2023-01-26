
/*
 *  Makro : vbf_Dataset2Datei    
 *
 *  Funktion :    Gibt den Dateinamen vom Dataset zur�ck.
 *
 * 
 *  Bemerkungen : Es erfolgt keine Pr�fung, ob das
 *                Dataset existiert. 
 *
 *                Fehlt die Angabe des Dateinamens,
 *                so wird Leer zur�ckgegeben.
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
 *  �nderungen :  August 2008    SAS-Code formatiert    Komenda    
 *                Mai    2011   Parameter ge�ndert in Dataset 
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
    /* der Parameter enth�lt keine Library */
    %let datei = &Dataset;
  %end;

  %if &pos = 1 and &len = 1 %then %do;
    /* der Parameter enth�lt nur "."  */
    %let datei = ;
  %end;

  %if &pos = 1 and &len > 1 %then %do;
    /* der Parameter enth�lt ".datei" */
    %let datei = %substr(&Dataset, 2);;
  %end;


  %if &pos = &len %then %do;
    /* Der Parameter enth�lt "Library." */    
    %let datei = ;
  %end;

  %local posp1;
  %local subword;

  %if &pos > 1 and &pos < &len %then %do;
    /* der Parameter enth�lt Library und Datei */
    %let posp1 = %eval(&pos+1);
    %let subword = %substr(&Dataset,&posp1);

    %if %index(&subword,.) > 0 %then %do;
      /* der Dateiname enth�lt einen Punkt */
      /* und ist damit ung�ltig            */
      %let datei = ;
    %end;

    %else %do;
      /* Substring auf den Dateinamen */      
      %let datei = %substr(&Dataset,&posp1);
    %end;
  %end;

  &datei
%mend ;
