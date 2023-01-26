
/*
 *  Makro : vbf_anfuehrungen_weg    
 *
 *  Funktion :  Entfernt die einfachen und doppelten Anführungszeichen.
 *
 * 
 *  Bemerkungen : Entfernt die Anführungszeichen am Anfang
 *                und am Ende des Parameters.
 *                Bei geschachtelten Anführungen werden nur 
 *                die äußersten Anführungszeichen entfernt. 
 *
 *  Einsatz:      Verbund      
 *  Sprache:      ab SAS 9.1.3
 *  Autor:        Jobst Heitzig
 *  
 *---------------------------------------------------------------------------------------
 * 
 *  Parameter :
 *    Text
 *
 *---------------------------------------------------------------------------------------
 *
 *  
 *  Änderungen :  Juni 2008    SAS-Code formatiert    Komenda    
 * 
 */


%macro vbf_anfuehrungen_weg(text);

  %if %length(%bquote(&text)) > 1 %then
    %do; 
      %local laenge_m2; /* Länge des Textes - 2 */
      %local char1;     /* Erstes Zeichen des Textes */
      %local charn;     /* Letztes Zeichen des Textes */

      %let laenge_m2 = %eval(%length(%bquote(&text)) - 2);

      %let char1 = %bquote(%substr(%bquote(&text), 1,  1));
      %let charn = %bquote(%substr(%bquote(&text), %length(%bquote(&text)),  1));

      %if     (%bquote(&char1) = %str(%') or %bquote(&char1) = %str(%"))
          and (%bquote(&charn) = %str(%') or %bquote(&charn) = %str(%")) %then
        %do;
          %if %length(%bquote(&text)) > 2 %then
            %let text = %substr(%bquote(&text), 2, &laenge_m2);
          %else
            %let text =;
        %end;
    %end;

 &text
%mend;
