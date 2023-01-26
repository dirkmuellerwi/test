
/*
 *  Makro : vb_ggf_fehler   
 *
 *  Funktion : Dieses Makro dient zusammen mit %vb_data_fehler dazu,
 *             nach einem fehlerhaften Data-Schritt eine Fehlermeldung
 *             auszugeben. 
 *
 *                
 *
 * 
 *  Bemerkungen : Kann nicht innerhalb von Data-Schritten oder Proc-IML-Schritten
 *                aufgerufen werden. Dort stattdessen die Makros %vb_ggf_fehler
 *                und %vb_data_fehler verwenden! 
 *
 *  Einsatz:      Verbund      
 *  Sprache:      ab SAS 9.1.3
 *  Autor;        Jobst Heitzig
 *
 *---------------------------------------------------------------------------------------
 * 
 *  Pflichtarameter :
 *       
 *  Optionale Parameter:
 *---------------------------------------------------------------------------------------
 *  
 *  Änderungen :   Komenda September 2008 Layout  
 * 
 */
%macro vb_ggf_fehler;
  %vb_teste_makrovar(vb_ggf_fehler_aufgetreten);

  %if &vb_makrovar_existiert %then %do;
    %if &vb_ggf_fehler_aufgetreten %then %do;
      %let vb_ggf_fehler_aufgetreten = 0;
			%vb_fehler(&vb_ggf_fehler_m, text = "&vb_ggf_fehler_text",
                   rc = %trim(&vb_ggf_fehler_rc), abbruch = &vb_ggf_fehler_abbruch);
    %end;
  %end;
%mend;
