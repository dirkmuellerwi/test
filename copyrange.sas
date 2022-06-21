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
	Log					=
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

	/*
	* Anlegen Java-Objekt und Setzen der Parameter
	*/
	declare javaobj excel("de/destatis/sas/excel22/RangeCopyUtil");

	/* Todo: Dynamische Log-Dateien werden z.Zt. noch nicht unterstützt. */

	/* excel.callVoidMethod("setLogEinstellungen", "&log", "&path_sasuser/updateexcel.log", "nein"); */

	excel.callVoidMethod("setSourceFile", __SourceFile);
	%if %length(&ExcelDateiData.) > 0 %then %do;
	  excel.callVoidMethod("setSourceFileData", __SourceFileData);
	%end;
	excel.callVoidMethod("setTargetFile", __TargetFile);
	excel.callVoidMethod("setSource", __SourceSheet, __SourceCell, __SourceRows, __SourceCols);
	excel.callVoidMethod("setTarget", __TargetSheet, __TargetCell);
	excel.callVoidMethod("doCopy");
	excel.callVoidMethod("save");
	*excel.callVoidMethod("restart");
	*excel.callVoidMethod("close");

	/* Todo: Exceptions hier in SAS abfangen und behandeln */

run; /* Ende Datastep */

%mend; %* %CopyRange;
