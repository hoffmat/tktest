package tkexcelscanner;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
* <p>
* Klasse fuer das Parsen von DataStore-Excels
*</p>
* <p>
* Nutzt die Apache POI Java-Api fuer Microsoft Dokumente
*</p>
*
* @version 1.0
* @author integration-factory
*/

public class ExcelParser {
	/** Valid-Status fuer das Excel */
	private int excelValidInd =1;
	
	/** Whitelist fuer die Attribut-Metadaten */
	private String[] metadatenWhitelist = {"Metadaten_Einzelwert.Name",
			"Metadaten_Einzelwert.Geschäftsbezeichnung",
			"Metadaten_Einzelwert.Beschreibung",
			"Metadaten_Einzelwert.R+V-Anwendung",
			"Metadaten_Einzelwert.Status",
			"Metadaten_Einzelwert.URL",
			"Metadaten_Einzelwert.DataStore-Cluster",
			"Metadaten_Liste.Fachlicher Ansprechpartner",
			"Metadaten_Liste.Technischer Ansprechpartner",
			"Metadaten_Liste.Information Assets"
	};
	
	/** Whitelist fuer die Hierarchie-Metadaten */
	private String[] hierarchyWhitelist = {"Ressource_Hierarchie.Name des Versorgungsverfahrens",
			"Ressource_Hierarchie.Versorgungsart"
			
	};
	
	/** String-Array fuer Ecxel Spaltennummerierung */
	private String [] excelColumns = new String [] 
			{"A",
			"B",
			"C",
			"D",
			"E",
			"F",
			"G",
			"H",
			"I",
			"J",
			"K",
			"L",
			"M",
			"N",
			"O",
			"P",
			"Q",
			"R",
			"S",
			"T",
			"U",
			"V",
			"W",
			"X",
			"Y",
			"Z"};
	
	
	/** File-Name des Excel-Files */
	private String file;
	/** Name des Excel-Reiters in dem Zusatzmapping-Informationen erwartet werden */
	private String zusatzmappingSheet = "Informatica Zusatzmappings";
	/** Name des Excel-Reiters in dem Versorgungsverfahren-Informationen erwartet werden */
	private String versVerfahrenSheet = "Informatica Vers.verfahren";
	/** Name des Excel-Reiters in dem ObjektzuRessource-Informationen erwartet werden */
	private String objektZuRessourceSheet = "Informatica Objekt zu Ressource";
	/** Name des Excel-Reiters in dem Lineage-Informationen erwartet werden */
	private String lineageSheet = "Informatica Lineage";
	
	/** SheetArray fuer Reiter Zusatzmappings */
	private String[][] zusatzmappingSheetArray;
	/** SheetArray fuer Reiter Versorgungsverfahren */
	private String[][] versVerfahrenSheetArray;
	/** SheetArray fuer Reiter ObjektzuRessource */
	private String[][] objektZuRessourceSheetArray;
	/** SheetArray fuer Reiter Lineage */
	private String[][] lineageSheetArray;
	
	/** Workbook Objekt */
	private XSSFWorkbook workbook;

	/** Variable für Instanz der Klasse Logging */
	private Logging xlsParseLogging;
	/** Variable für Instanz der Klasse Init */
	private Init init;
	
	/** HashMap fuer Zusatzmappings */
	private Map<Integer, Map<String, String>> zusatzmappingHashMap = new HashMap<>();
	/** HashMap fuer Versorgungsverfahren */
	private Map<Integer, Map<String, String>> versVerfahrenHashMap = new HashMap<>();
	/** HashMap fuer ObjektzuRessource */
	private Map<Integer, Map<String, String>> objektZuRessourceHashMap = new HashMap<>();
	/** HashMap fuer Lineage */
	private Map<Integer, Map<String, String>> lineageHashMap = new HashMap<>();
	/** HashMap dynamische Sheets */
	private Map<Integer, Map<String, String>> dynSheetsHashMap = new HashMap<>(); 
	
	
	
	
/**
* Klassenkonstruktor - bei der Instanziierung wird der Excel-Filename und Klasseninstanzen fuer {@link Logging} und {@link Init} uebergeben
*@param file Name des Excel-Files
*@param init Instanz der Klasse Init
*@param logging Instanz der Klasse Logging  
*/	 	
public ExcelParser(String file, Init init, Logging logging) throws IOException{

		this.xlsParseLogging = logging;
		this.init = init;
		
		this.file = file;
		
		this.setWorkbook();
        this.zusatzmappingSheetArray = this.createSheetArray(zusatzmappingSheet);
        this.versVerfahrenSheetArray = this.createSheetArray(versVerfahrenSheet);
        this.objektZuRessourceSheetArray = this.createSheetArray(objektZuRessourceSheet);
        this.lineageSheetArray = this.createSheetArray(lineageSheet);
           
        
        if (this.getExcelValidInd() == 1)
        {	  	
       
          this.createHashMap(this.zusatzmappingSheetArray, this.zusatzmappingHashMap);
          this.createHashMap(this.versVerfahrenSheetArray, this.versVerfahrenHashMap);
          this.setMetadataDefaults(this.versVerfahrenHashMap);
          
          this.createHashMap(this.objektZuRessourceSheetArray, this.objektZuRessourceHashMap);
          this.createHashMap(this.lineageSheetArray, this.lineageHashMap);
          
        }
          
        if (this.getExcelValidInd() == 1)
        {	  	
      
        	this.createHashMapDynSheets();
        	
        }
        
        
        if (this.getExcelValidInd() == 1)
        {	  	
      
        	this.validateMetadata();
        	
        }
        
        
        
    }

/**
* Erzeugung einer Workbook-Instanz fuer das Excel-File 
* @return Workbook-Objekt fuer das Excel-File 
*/	
private XSSFWorkbook setWorkbook() throws IOException{
	FileInputStream file = new FileInputStream(new File(this.file));
	this.workbook = new XSSFWorkbook(file);
	return workbook;
}


/**
* Erzeugung eines 2-Dimensionalen String Arrays (Matrix) fuer die Speicherung der Inhalte eines Excel-Reiters
* @param sheetName Name des Excel-Reiters fuer das ein SheetArray erzeugt werden soll 
* @return Liefert das SheetArray zurueck 
*/
private String[][] createSheetArray(String sheetName)
{	 
	XSSFSheet sheet = this.workbook.getSheet(sheetName);
	
	int noof_rows;
	//int noof_columns  = sheet.getRow(0).getPhysicalNumberOfCells()+1;
	int noof_columns = 200;
	if (sheet != null){
		 noof_rows  = sheet.getLastRowNum()+1;
		
	}
	else
	{
		 noof_rows  = 0;
		 noof_columns  = 0;	
		 this.excelValidInd = 0;
		 String logMessage = this.file +",,,Fehler,"+ "Reiter "+sheetName+" konnte im Excel "+ this.file + " nicht gefunden werden";
		 this.xlsParseLogging.writeLog(logMessage);
	}
	
	String [][] sheetArray = new String [noof_rows][noof_columns];
	
	if (this.getExcelValidInd() == 1)
	{		
		 String merged_value ="";
		   
	   
	    
		 for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
          CellRangeAddress region = sheet.getMergedRegion(i); //Region of merged cells
          int colIndex = region.getFirstColumn(); //number of columns merged
          int rowNum = region.getFirstRow();      //number of rows merged
          merged_value = sheet.getRow(rowNum).getCell(colIndex).getStringCellValue();
          
        
          int merged_first_row_index = region.getFirstRow();
          int merged_last_row_index = region.getLastRow();
           
         		 
          for (int row_cnt = merged_first_row_index; row_cnt <= merged_last_row_index; row_cnt++) { 
         
       	   sheetArray[row_cnt][colIndex] = merged_value;
          	 
          }
          
 		               
		 }
		 
		 Iterator<Row> rowIterator = sheet.iterator();
		 int row_index = 0;
		 int cell_index = 0;
	     String column_value ="";	
	    	
		 while (rowIterator.hasNext()) {
		        Row row = rowIterator.next();

				
 	       	row_index = row.getRowNum();
		        
 	       		if (row_index == 0)
 	       		{
 	       		sheetArray[row_index][197] = "Excel_File_Name";
 	       	    sheetArray[row_index][198] = "Sheet_Name";
 	       	    sheetArray[row_index][199] = "Sheet_Rownumber";
 	       		}
 	       		else 
 	       		{
 	       		sheetArray[row_index][197] = this.file;
 	       	    sheetArray[row_index][198] = sheetName;
 	       	    sheetArray[row_index][199] = Integer.toString(row_index+1);
 	       		}
 	       	
		        Iterator<Cell> cellIterator = row.cellIterator();
		       
		        outer:
		        while (cellIterator.hasNext()) {
		        	
		            Cell cell = cellIterator.next();
		            cell_index = cell.getColumnIndex();
	             	column_value = cell.getStringCellValue();
	   	          
	             	
	             	if (sheetArray[row_index][cell_index] == null) {
	             		sheetArray[row_index][cell_index] = column_value;
		         
		            }
		        }        	
      
		 }
       return sheetArray;
		}
	
       else 
       { 
       	return null;    
       }
	
 }

/**
* Erzeugung einer Hash-Map aus einem 2-Dimensionalen Sheet-Array, um Excel-Daten als Key/Value-Pairs zu speichern  
* @param sheetArray SheetArray fuer das eine HashMap erzeugt werden soll
* @param hashMap zu erezeugende HashMap
*/
private  void createHashMap(String [][] sheetArray, Map<Integer, Map<String, String>>  hashMap)
{
	
		 Map<String, String> hashMapInsert = new HashMap<>();
	
	 String value; 
	 int row_cnt = sheetArray.length-1;
     int column_cnt = sheetArray[0].length-1;
     
     int currHashMapSize = hashMap.size(); 
     
	 for (int curr_row=1; curr_row <=row_cnt;curr_row++) 
      {
    	  for (int curr_column = 0;curr_column <=column_cnt;curr_column++) {
    		  value = sheetArray[curr_row][curr_column]; 
    		  if (sheetArray[0][curr_column] != "" && sheetArray[0][curr_column] != null)
    		  {
    		    	
    			   	 hashMapInsert.put(sheetArray[0][curr_column], value);
    			   	 
    			     hashMap.put(currHashMapSize+curr_row, hashMapInsert);
    			  
    		  }
			  
    	  }
    	   hashMapInsert = new HashMap<>();
				
         }


}

/**
* Iteriert durch die im Excel-Reiter Informatica_Lineage angegeben Excel-Reiter (Spalte Reiter) und erzeugt Sheet-Arrays und HashMaps
* <br>Die Anzahl und Namen der Reiter sind hierbei dynamisch   
*/
private  void createHashMapDynSheets()
{
      int hashMapSize = this.lineageHashMap.size();
      String[][] sheetArray;
      String sheetName;
		
      String processedSheets = ""; 

	 
	 
	 for (int curr_pos = 1; curr_pos < hashMapSize; curr_pos ++)
	 {
		 sheetName = this.lineageHashMap.get(curr_pos).get("Reiter");			 
		 
		 if (sheetName != "" && sheetName != null)
		 {
			 if (processedSheets.contains("|#|" + sheetName.toUpperCase()) == false) 
			 {
				 processedSheets = processedSheets + ("|#|" + sheetName.toUpperCase());
				 sheetArray = this.createSheetArray(sheetName);
				 	
				 
				 if (this.getExcelValidInd() == 1)
				 {
					 this.createDynHashMap(sheetName,sheetArray, this.dynSheetsHashMap);
				 }
				 
			 }
			 
		 }
		 
		
	 }
	
}


/**
* Erzeugung einer dynamischen HashMap aus einem 2-Dimensionalen Sheet-Array fuer die Mapping-Excel-Reiter,<br>
* in denen die DataStore Umsetzungen (mappings) definiert sind 
* <br>Die Mapping-Daten aus dem entsprechenden Excel-Reiter werden als Key/Value Pairs abgelegt 
* @param sheetName Name des Excel-Reiters
* @param sheetArray SheetArray fuer das eine dynamische HashMap erezugt werden soll
* @param hashMap zu erzeugende HashMap  
*/
private  void createDynHashMap(String sheetName, String [][] sheetArray, Map<Integer, Map<String, String>>  hashMap)
{		
   
     int lineageHashMapSize = this.lineageHashMap.size();
     String lineageSheetName = "";
     int quellObjektPos = -1;
     int zielObjektPos = -1;
     int quellAttributPos = -1;
     int zielAttributPos = -1;
     int mappingTypPos = -1;
     int mappingBeschrPos = -1;
     
     
     //Für jede Lineage-Angabe im Reiter Informatica Lineage
     for (int curr_pos = 1; curr_pos <= lineageHashMapSize; curr_pos ++)
	 {
    	 try {
    		 lineageSheetName = this.lineageHashMap.get(curr_pos).get("Reiter");
			 }
			 catch(Exception e)
			 {
			 lineageSheetName = "";
			 }
    	 
    	 if (lineageSheetName !="" && lineageSheetName != null && sheetName.equals(lineageSheetName))
    	 {
    	 
    		 quellObjektPos = Arrays.asList(this.excelColumns).indexOf(this.lineageHashMap.get(curr_pos).get("Quellspalte Objekt"));
    		 zielObjektPos = Arrays.asList(this.excelColumns).indexOf(this.lineageHashMap.get(curr_pos).get("Zielspalte Objekt"));
    		 
    		 quellAttributPos = Arrays.asList(this.excelColumns).indexOf(this.lineageHashMap.get(curr_pos).get("Quellspalte Attribut"));
    		 zielAttributPos = Arrays.asList(this.excelColumns).indexOf(this.lineageHashMap.get(curr_pos).get("Zielspalte Attribut"));
    		 
    		 if(sheetName.equals("Informatica Zusatzmappings"))
    		 {
    			 mappingTypPos = -1;
    			 mappingBeschrPos = -1;
    		 }
    		 else
    		 {
    		 mappingTypPos = Arrays.asList(this.excelColumns).indexOf(this.lineageHashMap.get(curr_pos).get("Mapping-Spalte Mappingtyp"));
    		 mappingBeschrPos = Arrays.asList(this.excelColumns).indexOf(this.lineageHashMap.get(curr_pos).get("Mapping-Spalte Beschreibung"));
    		 } 
    	
    	 }
		
	 }
     
     

     Map<String, String> hashMapInsert = new HashMap<>();
		
	 String value; 
	 int row_cnt = sheetArray.length-1;
     int column_cnt = sheetArray[0].length-1;
     
     int currHashMapSize = hashMap.size(); 
     
	 for (int curr_row=1; curr_row <=row_cnt;curr_row++) 
      {
    	  for (int curr_column = 0;curr_column <=column_cnt;curr_column++) {
    		  value = sheetArray[curr_row][curr_column]; 
    		
    		  if (sheetArray[0][curr_column] != "" && sheetArray[0][curr_column] != null)
    		  {
    			    if (sheetArray[0][curr_column].equals("Excel_File_Name") 
    			    	| sheetArray[0][curr_column].equals("Sheet_Name")
    			    	| sheetArray[0][curr_column].equals("Sheet_Rownumber")
    			    	)
    			    {
    	    	     
    			   	 hashMapInsert.put(sheetArray[0][curr_column], value);
    			     hashMap.put(currHashMapSize+curr_row, hashMapInsert);
    			    }
    			    else 
    			    {
    			    	
    			    	if (quellObjektPos == curr_column)
    			    	{
    			    		hashMapInsert.put("Quellobjekt", value);
    			    	}
    			    	else if(quellAttributPos == curr_column)
    			    	{
    			    		hashMapInsert.put("Quellattribut", value);
    			    	}
    			    	else if(zielObjektPos == curr_column)
    			    	{
    			    		hashMapInsert.put("Zielobjekt", value);
    			    	}
    			    	else if(zielAttributPos == curr_column)
    			    	{
    			    		hashMapInsert.put("Zielattribut", value);
    			    	}
    			    	else if(mappingTypPos == curr_column)
    			    	{
    			    		hashMapInsert.put("Mappingtyp", value);
    			    	}
    			    	else if(mappingTypPos == -1)
    			    	{
    			    		hashMapInsert.put("Mappingtyp", "Zusatzmapping");
    			    	}
    			    	else if(mappingBeschrPos == curr_column)
    			    	{
    			    		hashMapInsert.put("Mappingbeschreibung", value);
    			    	}
    			    }
    			    
    		  }
			  
    	  }
    	   hashMapInsert = new HashMap<>();
    		 
      }
    
}

/**
* Prueft die aus dem Excel geparsten Metadaten auf Gueltigkeit <br>
* Es wird gegen Whitelists geprueft, ob die angegebenen Metadaten-Attribute bekannt sind 
* und es werden verschiedene Mindestanforderungen und Kombinationen fuer die Feldbelegung geprueft<br>
* Wenn Metadaten nicht valide sind, dann wird ein Fehler (Ausschluss des Excels) oder eine Warnung (Excel wird weiterverarbeitet) erzeugt<br>
* Alle erzeugten Meldungen sind in der logging.csv Datei einsehbar
*/
private void validateMetadata() throws IOException
{
	 String kategorieDerInformation ="";
	 String artDerInformation ="";
	 String allAttributes = "";
	 String attribute = "";
	 int attributeInWhitelist = 0;
	 int duplicateAttribute = 0;
	 
	 String allHierarchies = "";
	 int hierarchyInWhitelist = 0;
	 int duplicateHierarchy = 0;
	 
	 
	 int versVerfahrenHashMapSize = this.versVerfahrenHashMap.size();
	 int metadatenWhitelistSize =  this.metadatenWhitelist.length;
	 int hierarchyWhitelistSize =  this.hierarchyWhitelist.length;
		 
	 
	 for (int curr_pos = 1; curr_pos <= versVerfahrenHashMapSize; curr_pos ++)
	 {		 
		 kategorieDerInformation ="";
		 artDerInformation ="";
	 
		 try {
			 kategorieDerInformation = versVerfahrenHashMap.get(curr_pos).get("Kategorie der Information");
			 artDerInformation = versVerfahrenHashMap.get(curr_pos).get("Art der Information");
		 }
		 catch(Exception e)
		 {
			 
		 }
		 attribute = kategorieDerInformation+"."+artDerInformation;
		 
		 attributeInWhitelist = 0;
		 hierarchyInWhitelist = 0;

		 
		 if(kategorieDerInformation.equals("Metadaten_Einzelwert") | kategorieDerInformation.equals("Metadaten_Liste"))
		 {
		 
		 
			 for (int i = 0; i< metadatenWhitelistSize; i++)
			 {
				 if(this.metadatenWhitelist[i].equals(attribute))
				 {   attributeInWhitelist = 1;
					 if(allAttributes.contains(attribute+"#"))
					 {
						 duplicateAttribute = 1;
					 }
					 else
					 {
						 allAttributes = allAttributes + attribute+"#";
						 duplicateAttribute = 0;
					 }
				 }
				 
			 } 
			 
				 
			 if(attributeInWhitelist == 0)
			 {
				 String logMessage = this.file +","+this.versVerfahrenHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.versVerfahrenHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Warnung,"
				                     +"Unbekanntes Attribut -> ["+kategorieDerInformation+","+artDerInformation+"]";
						 
				 this.xlsParseLogging.writeLog(logMessage);
			 }
			 
			 if(duplicateAttribute == 1)
			 {
				 this.excelValidInd = 0;
				 String logMessage = this.file +","+this.versVerfahrenHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.versVerfahrenHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Fehler,"
				                     +"Attribut mehrfach definiert -> ["+kategorieDerInformation+","+artDerInformation+"]";
						 
				 this.xlsParseLogging.writeLog(logMessage);
				 
			 }
		 
	 }	 
		 
		 if(kategorieDerInformation.equals("Ressource_Hierarchie") )
		 {
		 
		 
			 for (int i = 0; i< hierarchyWhitelistSize; i++)
			 {
				 if(this.hierarchyWhitelist[i].equals(attribute))
				 {   hierarchyInWhitelist = 1;
					 if(allHierarchies.contains(attribute+"#"))
					 {
						 duplicateHierarchy = 1;
					 }
					 else
					 {
						 allHierarchies = allHierarchies + attribute+"#";
						 duplicateAttribute = 0;
					 }
				 }
				 
			 } 
			 
				 
			 if(hierarchyInWhitelist == 0)
			 {
				 String logMessage = this.file +","+this.versVerfahrenHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.versVerfahrenHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Warnung,"
				                     +"Unbekannte Ressource-Hierarchie -> ["+kategorieDerInformation+","+artDerInformation+"]";
						 
				 this.xlsParseLogging.writeLog(logMessage);
			 }
			 
			 if(duplicateHierarchy == 1)
			 {
				 this.excelValidInd = 0;
				 String logMessage = this.file +","+this.versVerfahrenHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.versVerfahrenHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Fehler,"
				                     +"Ressource-Hierarchie mehrfach definiert -> ["+kategorieDerInformation+","+artDerInformation+"]";
						 
				 this.xlsParseLogging.writeLog(logMessage);
				 
			 }
		 
		 }		 
	 }
	
	 
	//Prüfung Anzahl der Ressourcen-Hierachien 
    int noofHierarchies = 0;
	for (int i = 0; i< hierarchyWhitelistSize; i++)
		 {
		for (int curr_pos = 1; curr_pos <= versVerfahrenHashMapSize; curr_pos ++)
		 {
			try {
				 kategorieDerInformation = versVerfahrenHashMap.get(curr_pos).get("Kategorie der Information");
				 artDerInformation = versVerfahrenHashMap.get(curr_pos).get("Art der Information");
				 
			 }
			 catch(Exception e)
			 {
				 
			 }
			
			if(this.hierarchyWhitelist[i].equals(kategorieDerInformation+"."+artDerInformation))
			{
			noofHierarchies ++;	
			}
			
			
			
		 }
		
		 }
	
	if(noofHierarchies!=hierarchyWhitelistSize) 
	{
		 this.excelValidInd = 0;
		 String logMessage = this.file +","+this.versVerfahrenHashMap.get(1).get("Sheet_Name")+","
		                     +","
		                     +"Fehler,"
		                     +"Anzahl der Ressource-Hierarchie ist inkorrekt -> Erwartet:"+hierarchyWhitelistSize+" Gefunden:"+noofHierarchies;
				 
		 this.xlsParseLogging.writeLog(logMessage);
	}
	
	
	
	
	 
	 
	//Prüfung Mappings 
	 int dynSheetsHashMapSize = this.dynSheetsHashMap.size();
	 
	 String mappingTyp = "";
	 String quellobjekt = "";
	 String quellattribut = "";
	 String zielobjekt = "";
	 String zielattribut = "";
		 
	 
	 for (int curr_pos = 1; curr_pos <= dynSheetsHashMapSize; curr_pos ++)
	 {
		 mappingTyp = "";
		 
		 try {

			 mappingTyp = this.dynSheetsHashMap.get(curr_pos).get("Mappingtyp");
			 quellobjekt = this.dynSheetsHashMap.get(curr_pos).get("Quellobjekt");
			 quellattribut = this.dynSheetsHashMap.get(curr_pos).get("Quellattribut");
			 zielobjekt = this.dynSheetsHashMap.get(curr_pos).get("Zielobjekt");
			 zielattribut = this.dynSheetsHashMap.get(curr_pos).get("Zielattribut");
		 
			 
		 }
		 catch(Exception e)
		 {
			 mappingTyp = "";
		 }
		 
		 if ( mappingTyp != null)
		 {
		 // Invalider Mappingtyp
		 if ( mappingTyp != null &&  mappingTyp !="" &&  mappingTyp !="Zusatzmapping") 
		 {
			 if (!mappingTyp.equals("Eins_zu_Eins") && !mappingTyp.equals("Konstanter_Wert") && !mappingTyp.equals("Transformation"))
			 {
				 this.excelValidInd = 0;
				 String logMessage = this.file +","+this.dynSheetsHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.dynSheetsHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Fehler,"
				                     +"Unbekannter Mappinptyp -> "+mappingTyp;
						 
				 this.xlsParseLogging.writeLog(logMessage);  
			 }
		 }
		 
		 // Fehlendes Objekt / Attribut
		 if ( mappingTyp.equals("Eins_zu_Eins")  |  mappingTyp.equals("Transformation")) 
		 {
			 if (quellobjekt == null | quellobjekt == "" | quellobjekt == "nicht vorhanden" )
			 {
				 this.excelValidInd = 0;
				 String logMessage = this.file +","+this.dynSheetsHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.dynSheetsHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Fehler,"
				                     +"Quellobjekt für Mapping nicht angegeben";
						 
				 this.xlsParseLogging.writeLog(logMessage);  
			 }
	
			 if (quellattribut == null | quellattribut == "" | quellattribut == "nicht vorhanden")
			 {
				 this.excelValidInd = 0;
				 String logMessage = this.file +","+this.dynSheetsHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.dynSheetsHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Fehler,"
				                     +"Quellattribut für Mapping nicht angegeben";
						 
				 this.xlsParseLogging.writeLog(logMessage);  
			 }
		 
			 if (zielobjekt == null | zielobjekt == "" | zielobjekt == "nicht vorhanden")
			 {
				 this.excelValidInd = 0;
				 String logMessage = this.file +","+this.dynSheetsHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.dynSheetsHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Fehler,"
				                     +"Zielobjekt für Mapping nicht angegeben";
						 
				 this.xlsParseLogging.writeLog(logMessage);  
			 }
	
			 if (zielattribut == null | zielattribut == "" | zielattribut == "nicht vorhanden")
			 {
				 this.excelValidInd = 0;
				 String logMessage = this.file +","+this.dynSheetsHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.dynSheetsHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Fehler,"
				                     +"Zielattribut für Mapping nicht angegeben";
						 
				 this.xlsParseLogging.writeLog(logMessage);  
			 }
			 
		 }
		 
		 
		 if ( mappingTyp.equals("Konstanter_Wert") ) 
		 {
			 if (zielobjekt == null | zielobjekt == "" | zielobjekt == "nicht vorhanden")
			 {
				 this.excelValidInd = 0;
				 String logMessage = this.file +","+this.dynSheetsHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.dynSheetsHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Fehler,"
				                     +"Zielobjekt für Mapping nicht angegeben";
						 
				 this.xlsParseLogging.writeLog(logMessage);  
			 }
	
			 if (zielattribut == null | zielattribut == "" | zielattribut == "nicht vorhanden")
			 {
				 this.excelValidInd = 0;
				 String logMessage = this.file +","+this.dynSheetsHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.dynSheetsHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Fehler,"
				                     +"Zielattribut für Mapping nicht angegeben";
						 
				 this.xlsParseLogging.writeLog(logMessage);  
			 }
		 }
		 }
		 else
		 {
		 
		 if ( mappingTyp == null | mappingTyp == "" ) 
		 {
			 if (zielobjekt == null | zielobjekt == "" | zielobjekt == "nicht vorhanden")
			 {
				 this.excelValidInd = 0;
				 String logMessage = this.file +","+this.dynSheetsHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.dynSheetsHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Fehler,"
				                     +"Zielobjekt für Mapping nicht angegeben";
						 
				 this.xlsParseLogging.writeLog(logMessage);  
			 }
	
			 if ((zielattribut == null | zielattribut == "" | zielattribut == "nicht vorhanden") && !this.dynSheetsHashMap.get(curr_pos).get("Sheet_Name").equals("Informatica Zusatzmappings"))
			 {
				 this.excelValidInd = 0;
				 String logMessage = this.file +","+this.dynSheetsHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.dynSheetsHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Fehler,"
				                     +"Zielattribut für Mapping nicht angegeben";
						 
				 this.xlsParseLogging.writeLog(logMessage);  
			 }
		 }
		 
		 if ( mappingTyp == null | mappingTyp == "" ) 
		 {
			 if (quellobjekt == null | quellobjekt == "" | quellobjekt == "nicht vorhanden")
			 {
				 this.excelValidInd = 0;
				 String logMessage = this.file +","+this.dynSheetsHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.dynSheetsHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Fehler,"
				                     +"Quellobjekt für Mapping nicht angegeben";
						 
				 this.xlsParseLogging.writeLog(logMessage);  
			 }
	
			 if ((quellattribut == null | quellattribut == "" | quellattribut == "nicht vorhanden") && !this.dynSheetsHashMap.get(curr_pos).get("Sheet_Name").equals("Informatica Zusatzmappings"))
			 {
				 this.excelValidInd = 0;
				 String logMessage = this.file +","+this.dynSheetsHashMap.get(curr_pos).get("Sheet_Name")+","
				                     +this.dynSheetsHashMap.get(curr_pos).get("Sheet_Rownumber")
				                     +",Fehler,"
				                     +"Quellattribut für Mapping nicht angegeben";
						 
				 this.xlsParseLogging.writeLog(logMessage);  
			 }
		 }
		 }
	 }
}


/**
* Setzen von Metadaten-Default Werten, wenn diese nicht im Excel angegeben sind
* @param versVerfahrenHashMap HashMap fuer das Versorgungsverfahren 
* 
*/	
private void setMetadataDefaults(Map<Integer, Map<String, String>> versVerfahrenHashMap) {
	 int versVerfahrenHashMapSize = versVerfahrenHashMap.size();
	 	
	 String [] metadataDefaultKeys = {"Name",
	                                  "Geschäftsbezeichnung",
	                                  "Beschreibung",
	                                  "Fachlicher Ansprechpartner",
	                                  "Technischer Ansprechpartner",
	                                  "Information Assets",
	                                  "R+V-Anwendung",
	                                  "Status",
	                                  "URL",
	                                  "DataStore-Cluster"};
	 
	 String [] metadataDefaultValues = {"",
             "",
             "",
             "V_Programm_Omnichannel_Datastore_Kernteam",
             "V_Programm_Omnichannel_Datastore_Kernteam",
             "IA 9",
             "APP-4837",
             "aktiv",
             "",
             ""};

	 
	 String artDerInformation = "";
	 int metadataDefaultKeysSize = metadataDefaultKeys.length;
	 int metadataExist = 0;
	 int versVerfahrenHashMapCurrCnt = versVerfahrenHashMapSize; 
	 
	 Map<String, String> hashMapInsert = new HashMap<>();
	 for (int i=0; i < metadataDefaultKeysSize;i++)
	
		
	 {
	     metadataExist = 0; 
	 
	 
		 for (int curr_pos = 1; curr_pos <= versVerfahrenHashMapSize; curr_pos ++)
		 {
			 
			 try {
				 artDerInformation = versVerfahrenHashMap.get(curr_pos).get("Art der Information");
				 
			 }
			 catch(Exception e)
			 {
				 
			 }
			 
			 if (metadataDefaultKeys[i].equals(artDerInformation))
			 {metadataExist = 1;
					 break;
					 
					 
			 }
			
			 
		 }  
		 if (metadataExist == 0)
		 {
			 versVerfahrenHashMapCurrCnt++;	
			 hashMapInsert.put("Kategorie der Information", "Metadaten");
			 hashMapInsert.put("Art der Information", metadataDefaultKeys[i]);
			 hashMapInsert.put("Information", metadataDefaultValues[i]);
			 this.versVerfahrenHashMap.put(versVerfahrenHashMapCurrCnt, hashMapInsert);
			 hashMapInsert = new HashMap<>();
			 
		 }
	 }
}	
	
	
	
	

/**
* Methode fuer das Print-Out eines HashMap-Inhalts auf die Console 
* @param hashMap die ausgegeben werden soll
*/
public  void printHashMap( Map<Integer, Map<String, String>>  hashMap)
	 {	  
		 
		 int hashMapSize = hashMap.size();
		 
			 
		 for (int curr_pos = 1; curr_pos <= hashMapSize; curr_pos ++)
		 {
			 hashMap.get(curr_pos).forEach((k,v) -> System.out.println("Key "+k +" ##### Value "+v)); 
			 
		 }
				
		 
	}

	
	
	
	
/**
* Get-Methode fuer den Excel-Valid Status 
* @return Excel-Valid Status
*/	
public int getExcelValidInd() {
		 return this.excelValidInd;
}
	 

		
		
/**
* Get-Methode fuer die HashMap zusatzmappingHashMap
* @return  HashMap zusatzmappingHashMap
*/			
public Map<Integer, Map<String, String>> getZusatzmappingHashMap() {
			 return this.zusatzmappingHashMap;
}
	
/**
* Get-Methode fuer die HashMap versVerfahrenHashMap
* @return  HashMap versVerfahrenHashMap
*/	
public Map<Integer, Map<String, String>> getVersVerfahrenHashMap() {
			 return this.versVerfahrenHashMap;
}
  
/**
* Get-Methode fuer die HashMap objektZuRessourceHashMap
* @return  HashMap objektZuRessourceHashMap
*/	
public Map<Integer, Map<String, String>> getObjektZuRessourceHashMap() {
		 return this.objektZuRessourceHashMap;
}
 
/**
* Get-Methode fuer die HashMap lineageHashMap
* @return  HashMap lineageHashMap
*/	 
public Map<Integer, Map<String, String>> getLineageHashMap() {
		 return this.lineageHashMap;
}
     
/**
* Get-Methode fuer die HashMap dynSheetsHashMap
* @return  HashMap lineageHashMap
*/	 
public Map<Integer, Map<String, String>> getDynSheetsHashMap() {
		 return this.dynSheetsHashMap;
}
	 
     

/**
* Methode fuer das Print-Out eines SheetArray-Inhalts auf die Console 
* @param  sheetArray SheetArray das ausgegeben werden soll 
*/	 
public  void printSheetArray(String[][] sheetArray)
{	  
		 
		 int row_cnt = sheetArray.length-1;
	      int column_cnt = sheetArray[0].length-1;
		 
	      for (int curr_row=0; curr_row <=row_cnt;curr_row++) 
	      {
	    	  System.out.println("############ Zeilennummer  " + String. valueOf(curr_row) + "  #############################   ");
	    	  for (int curr_column = 0;curr_column <=column_cnt;curr_column++) {
	    		  System.out.println("Spalte  " + String. valueOf(curr_column) + " : " +  sheetArray[curr_row][curr_column]);
	             	
	    	  }
	      }
	      
	      
	    }
}
