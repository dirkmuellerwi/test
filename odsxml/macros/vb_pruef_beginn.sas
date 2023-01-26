/* 
 * Programm:  vb_Pruef_Beginn
 * Version:   1.0
 * Einsatz:   Verbund
 * Sprache:   SAS 9.2
 * Autor:     C.Klein
 * Telefon:   4571
 *================================================================================================
 * Funktion: Sollen mehrere Pr�fungsmakros hintereinander durchgef�hrt werden, ohne bei einem
 *           Fehler die Pr�fungen abzubrechen, wird f�r diese Pr�fungsreihe ein
 *           �bergeordneter Makroname durch den Parameter &Makroname bestimmt und die beiden globalen
 *           Makrovariablen &Makroname._rc (Returncode) und &Makroname._rt (Fehlerbeschreibung)
 *           gesetzt, &Makroname._rc auf 0 und &Makroname._rt auf [leer].
 *
 *           Dieser Makro legt 2 globale Makrovariablen festen Namens an:
 *           _vb_Pruefreihenmakroname,     die den �bergeordneten Makronamen enth�lt,
 *           _vb_PruefreihenSammelmeldung, die steuert, ob Fehlermeldungen direkt nach den
 *                                         einzelnen Pr�fungen ausgegeben werden oder gesammelt
 *                                         am Ende der Pr�fungsreihe.
 *                                         Vorgabe ist ja, und diese Vorgabe wird im die
 *                                         Fehlermeldung steuernden Makro vb_Pruef_Auswertung �berpr�ft.
 *           Diese beiden Makrovariablen werden im Makro vb_Pruef_Ende gel�scht.
 *           Sollten bei Aufruf von vb_Pruef_Anfang diese Makrovariablen global vorhanden sein,
 *           wird eine nicht beendete Pr�fungsreihe angenommen, und es folgt Programmabbruch.
 *           Programmabbruch folgt au�erdem, wenn kein �bergeordneter Makroname
 *           (Parameter &Makroname) angegeben wird.
 *================================================================================================
 * Parameter:                   Beschreibung:
 *------------------------------------------------------------------------------------------------
 * Makroname                    - Der �bergeordnete Makroname f�r die Einzelpr�fungen.
 *================================================================================================
 * �nderungsverlauf: erstellt am 27.IV.2011
 */

%macro vb_Pruef_Beginn 
        (Makroname     = ,
         Sammelmeldung = ja);
    %if %length(&Makroname) = 0 %then %do;
        %vb_Abbruch
            (Makroname = vb_Pruef_Beginn,
             Text      = Pr�fungsreihe ohne Makroname);
    %end;

    %if     %symglobl(_vb_Pruefreihenmakroname)
        and %symglobl(_vb_PruefreihenSammelmeldung) %then %do;
        %vb_Abbruch
            (Makroname = vb_Pruef_Beginn,
             Text      = Nicht abgeschlossene Pr�freihe (&_vb_Pruefreihenmakroname) aktiv);
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
