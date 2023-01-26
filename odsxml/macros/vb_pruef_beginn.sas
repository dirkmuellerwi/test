/* 
 * Programm:  vb_Pruef_Beginn
 * Version:   1.0
 * Einsatz:   Verbund
 * Sprache:   SAS 9.2
 * Autor:     C.Klein
 * Telefon:   4571
 *================================================================================================
 * Funktion: Sollen mehrere Prüfungsmakros hintereinander durchgeführt werden, ohne bei einem
 *           Fehler die Prüfungen abzubrechen, wird für diese Prüfungsreihe ein
 *           übergeordneter Makroname durch den Parameter &Makroname bestimmt und die beiden globalen
 *           Makrovariablen &Makroname._rc (Returncode) und &Makroname._rt (Fehlerbeschreibung)
 *           gesetzt, &Makroname._rc auf 0 und &Makroname._rt auf [leer].
 *
 *           Dieser Makro legt 2 globale Makrovariablen festen Namens an:
 *           _vb_Pruefreihenmakroname,     die den übergeordneten Makronamen enthält,
 *           _vb_PruefreihenSammelmeldung, die steuert, ob Fehlermeldungen direkt nach den
 *                                         einzelnen Prüfungen ausgegeben werden oder gesammelt
 *                                         am Ende der Prüfungsreihe.
 *                                         Vorgabe ist ja, und diese Vorgabe wird im die
 *                                         Fehlermeldung steuernden Makro vb_Pruef_Auswertung überprüft.
 *           Diese beiden Makrovariablen werden im Makro vb_Pruef_Ende gelöscht.
 *           Sollten bei Aufruf von vb_Pruef_Anfang diese Makrovariablen global vorhanden sein,
 *           wird eine nicht beendete Prüfungsreihe angenommen, und es folgt Programmabbruch.
 *           Programmabbruch folgt außerdem, wenn kein übergeordneter Makroname
 *           (Parameter &Makroname) angegeben wird.
 *================================================================================================
 * Parameter:                   Beschreibung:
 *------------------------------------------------------------------------------------------------
 * Makroname                    - Der übergeordnete Makroname für die Einzelprüfungen.
 *================================================================================================
 * Änderungsverlauf: erstellt am 27.IV.2011
 */

%macro vb_Pruef_Beginn 
        (Makroname     = ,
         Sammelmeldung = ja);
    %if %length(&Makroname) = 0 %then %do;
        %vb_Abbruch
            (Makroname = vb_Pruef_Beginn,
             Text      = Prüfungsreihe ohne Makroname);
    %end;

    %if     %symglobl(_vb_Pruefreihenmakroname)
        and %symglobl(_vb_PruefreihenSammelmeldung) %then %do;
        %vb_Abbruch
            (Makroname = vb_Pruef_Beginn,
             Text      = Nicht abgeschlossene Prüfreihe (&_vb_Pruefreihenmakroname) aktiv);
    %end;

    %global _vb_Pruefreihenmakroname;
    %let _vb_Pruefreihenmakroname = &Makroname;
    %global _vb_PruefreihenSammelmeldung;
    %let _vb_PruefreihenSammelmeldung = &Sammelmeldung;

    %global &Makroname._rc;
    %let &Makroname._rc = 0;

    %global &Makroname._rt;
    %let &Makroname._rt = ;
%mend;
