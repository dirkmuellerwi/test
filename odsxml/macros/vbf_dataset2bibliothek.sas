/*
 *  Makro : vbf_Dataset2Bibliothek    
 *
 *  Funktion :    Das Makro gibt die Bibliothek des DataSets zur�ck.
 * 
 *  Bemerkungen : Fehlt die Angabe der Bibliothek wird Work zurueckgegeben.
 *                Ist die Angabe Bibliothek.Dateiename ung�ltig, so wird Leer
 *                zurueckgegeben.
 *
 *                Es erfolgt keine Pr�fung auf G�ltigkeit der Bibliothek 
 * 
 *  
 *  Einsatz:      Verbund      
 *  Sprache:      ab SAS 9.1.3
 *  Autor:
 *---------------------------------------------------------------------------------------
 * 
 *  Parameter :
 *    Dataset  Dataset im Format Bibliothek.Dateiname.
 *
 *
 *---------------------------------------------------------------------------------------
 *
 *  
 *  �nderungen :  Oktober 2008    SAS-Code formatiert    Komenda    
 *                Mai     2011   Parameter ge�ndert in Dataset 
 * 
 */

%macro vbf_Dataset2Bibliothek(Dataset);

  %let Dataset = %vbf_anfuehrungen_weg(&Dataset);

  %local len;
  %local pos;
  %local posp1;
  %local subword;
  %local posm1;
  %local lensub;


  %let len = %length(&Dataset);
  %let pos = %index(&Dataset,.);

  %local bibliothek;

  %if &len = 0 %then %do;
    /* Paramet�r ist leer */
    %let bibliothek = ;
  %end;

  %else %if &pos = 0 %then %do;
    /* Bibliothek fehlt */
    %let bibliothek = work;
  %end;

  %else %if &pos = 1 and &len = 1 %then %do;
    /* Bibliothek fehlt */
    %let bibliothek = work;
  %end;

  %else %if &pos = 1 and &len > 1 %then %do;
    /* Bibliothek fehlt */
    %let bibliothek = work;
  %end;


  %else %if &pos = &len %then %do;
    /* Bibliothek ohne Datei */
    %let lensub = %eval(&len-1);
    %let bibliothek = %substr(&Dataset, 1, &lensub);
  %end;

  %else %do;
    /* Bibliothek mit Datei */
    %let posp1 = %eval(&pos+1);
    %let subword = %substr(&Dataset,&posp1);
    
    %if %index(&subword,.) > 0 %then %do;
      /* Der Parameter enth�lt mindestens zwei Punkte */
      %let bibliothek = ;
    %end;
    %else %do;
      /* Substring auf die Bibliothek */
      %let posm1 = %eval(&pos-1);
      %let bibliothek = %substr(&Dataset,1,&posm1);
    %end;
  %end;
  &bibliothek
%mend ;
