%macro  CopyRange(
	/* Quelle */
	ExcelDatei			= ,
	ExcelDateiData		= ,
	Arbeitsblatt		= ,
	Zelle				= ,
	AnzahlZeilen		= ,
	AnzahlSpalten		= ,
	/* Ziel */
	ExcelDateiZiel		= ,
	ArbeitsblattZiel	= ,
	ZelleZiel			= ,
	Log					= ,
	LimitZeilen			= ,
	LimitSpalten		= ,
	LimitZellen			= ,
	LimitGroesse		= 
);

/*
* Datastep mit Aufruf der Java-Methoden
*/ 

data _null_;

	length
		__SourceFile			$	1000
		__SourceFileData		$	1000
		__SourceSheet			$	31
		__SourceCell			$ 	8
		__SourceRows				8
		__SourceCols				8
		__TargetFile			$	1000
		__TargetSheet			$	31
		__TargetCell			$	8
		__LimitZeilen				8
		__LimitSpalten				8
		__LimitZellen				8
		__LimitGroesse				8
	;

	__SourceFile 		= "&ExcelDatei.";
	__SourceFileData 	= "&ExcelDateiData.";
	__SourceSheet 		= "&Arbeitsblatt.";
	__SourceCell 		= "&Zelle.";
	__SourceRows 		= &AnzahlZeilen.;
	__SourceCols 		= &AnzahlSpalten.;
	__TargetFile 		= "&ExcelDateiZiel.";
	__TargetSheet 		= "&ArbeitsblattZiel.";
	__TargetCell 		= "&ZelleZiel.";
	/* Einlesen mit "input()" fuer den Fall, dass Variablen leer sind */
	__LimitZeilen		= input ("&LimitZeilen.", best.);
	__LimitSpalten		= input ("&LimitSpalten.", best.);
	__LimitZellen		= input ("&LimitZellen.", best.);
	__LimitGroesse		= input ("&LimitGroesse.", best.);

	/* Anlegen Java-Objekt und Setzen der Parameter */
	declare javaobj excel("de/destatis/sas/excel22/RangeCopyUtil");
	if __LimitZeilen then excel.callVoidMethod("setRowCountLimit", __LimitZeilen);
	if __LimitSpalten then excel.callVoidMethod("setColumnCountLimit", __LimitSpalten);
	if __LimitZellen then excel.callVoidMethod("setCellCountLimit", __LimitZellen);
	if __LimitGroesse then excel.callVoidMethod("setFileSizeLimit", __LimitGroesse);

	/* Todo: Dynamische Log-Dateien werden z.Zt. noch nicht unterst¸tzt. */
	/* excel.callVoidMethod("setLogEinstellungen", "&log", "&path_sasuser/updateexcel.log", "nein"); */

	excel.callVoidMethod("setTemplateFile", __SourceFile);
	%if %length(&ExcelDateiData.) > 0 %then %do;
	  excel.callVoidMethod("setSourceFile", __SourceFileData);
	%end;
	excel.callVoidMethod("setTargetFile", __TargetFile);
	excel.callVoidMethod("setSource", __SourceSheet, __SourceCell, __SourceRows, __SourceCols);
	excel.callVoidMethod("setTarget", __TargetSheet, __TargetCell);
	excel.callVoidMethod("doCopy");
	excel.callVoidMethod("save"); /* Save() schlieﬂt Close() ein */

	/* Exceptions abfangen und behandeln */
	length __errorMode __returnCode 8 __errorMessage $ 200;
	excel.callBooleanMethod("getErrorMode", __errorMode);

	if __errorMode then do;
		__returnCode = 8;
		/* Fehlertext holen */
		excel.callStringMethod ("getErrorMessage", __errorMessage);
		call symput("fehlerText", trim(__errorMessage));
		call symput("returnCode", putn(__returnCode, "2.0"));
	end;
	put _all_;

run; /* Ende Datastep */

%mend; %* %CopyRange;
