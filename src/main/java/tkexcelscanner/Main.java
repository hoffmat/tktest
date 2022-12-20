package tkexcelscanner;

import java.io.IOException;
import java.nio.file.Path;

import java.util.*; 


/**
* <p>
* Hauptklasse der DataStore Excel-Parser Applikation
*</p>
* <p>
* Iteriert durch alle DataStore-Excels und instanziiert die Klassen {@link Init}, {@link Logging}, {@link ExcelParser} und {@link HubApi} Klassen
*</p>
*
* @version 1.0
* @author integration-factory
*/


public class Main {

	/**
	*Iteriert durch alle DataStore-Excels und instanziiert die Klassen {@link Init}, {@link Logging}, {@link ExcelParser} und {@link HubApi} Klassen
	*@param args String-Array fuer Applikationsparameter
	*/
	public static void main(String[] args) throws IOException {
		
		/** Intiliasierung der Applikation */
		 Init init = new Init();
		 
		 String excelFilePath = init.getParametervalue("quellverzeichnis");
		 
		 /** Erzeugung File-List fuer alle Datastore-Excels */
		 FileList xlsFilelist = new FileList(excelFilePath);
		 List<Path> fileList = xlsFilelist.getFileList();
			 
		 /** Erzeugung der Logging-Instanz */
		 Logging logging = new Logging(init);
		 
		 /** Erzeugung der HubApi-Instanz */
		 HubApi hubApi = new HubApi(init,logging); 
		 
		 
		 int excelCnt = 0;
		 /** Iteration über alle DataStore-Excel im angegebenen Quellverzeichnis*/
		 for (Iterator<Path> files = fileList.iterator(); files.hasNext(); ) {
	     	   String file = files.next().toString();
	     	  /** Erzeugung der einer Excel-Parser-Instanz pro Excel-File */
			   ExcelParser xlsParser = new ExcelParser(file, init, logging);
			   
			   /** Wenn das Excel-File valide ist, werden HUB-Metadaten erzeugt */		  
			   if(xlsParser.getExcelValidInd() == 1)
			   {
			   hubApi.generateHubFiles(xlsParser.getZusatzmappingHashMap()
					   , xlsParser.getVersVerfahrenHashMap()
					   , xlsParser.getObjektZuRessourceHashMap()
					   , xlsParser.getDynSheetsHashMap());
			   }   
			   excelCnt++;
		
		 }
	 			 
		 logging.closeLogFile();
		 System.out.println("Ende der Verarbeitung - "+excelCnt+" Excels geparst");
		 
		}

}

