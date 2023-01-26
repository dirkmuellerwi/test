/*
 *  Makro : vbf_dataset2ordner    
 *
 *  Funktion :    Gibt den vollständigen Pfad der Bibliothek
 *                vom Dataset zurück.
 * 
 *  Bemerkungen : Falls die Bibliothek auf mehrere Ordner verweist,
 *                wird der gültige Ordner zum Dataset zurückgegeben.
 *
 *                Existiert kein gültiger Ordner, wird Leer zurückgegeben.
 * 
 *  
 *  Einsatz:      Verbund      
 *  Sprache:      ab SAS 9.1.3
 *  Autor:
 *---------------------------------------------------------------------------------------
 * 
 *  Parameter :
 *    Dataset  Name der SAS-Bibliothek
 *
 *
 *---------------------------------------------------------------------------------------
 *
 *  
 *  Änderungen :  August 2008    SAS-Code formatiert    Komenda    
 *                Mai    2011   Parameter geändert in Dataset 
 * 
 */

%macro vbf_dataset2ordner(Dataset);  
  %local lib datei;
  %local pfad;
  %let pfad =;

  %let lib   = %vbf_dataset2bibliothek(&Dataset);
  %let datei = %vbf_dataset2datei(&Dataset);
     
  %if  %length(&lib) > 0  and   %length(&datei) > 0  %then
    %do;
      /* Patrameter enthält Library und Dateiname */ 
      %let pfad = %sysfunc(pathname(&lib));


      %if %length(%scan(&pfad, 2, ' ')) > 0 %then
        %do;
          /*
           *  Die Liste besteht aus mehr als einem Verzeichnis.
           *  Entfernen der Klammern ...
           */
          %let pfad = %substr(&pfad, 2, %length(&pfad) - 2);

          %local abbruch start verz;
          %let abbruch = 0;
          %let start = 1;

          /*
           * Schleife über alle Verzeichnisse
           */
          %do %while (&abbruch = 0);
            %let verz = %scan(&pfad, &start, ' ');

            %if %length(&verz) > 0 %then 
              %do;
                %let start = %eval(&start + 1);

                /* Temp Library auf das Verzeichnis */
                %let rc = %sysfunc(libname(_temp_, &verz));

                %if %sysfunc(libref(_temp_)) = 0 %then
                  %do;
                    %if %sysfunc(exist(_temp_.&datei)) %then 
                      %do;
                        /* Die Sas-Datei befindet sich im Verzeichnis */
                        %let pfad = &verz;
                        %let abbruch = 1; /* Beenden der while Schleife */
                      %end;
                    %let rc = %sysfunc(libname(_temp_));
              %end;
            %end;

            %else
              %do;
                /* Es wurde kein gültiges Verzeichnis gefunden */
                %let abbruch = 1;
                %let pfad =;
              %end;
          %end;
      %end;
    %end;

  &pfad
%mend;
