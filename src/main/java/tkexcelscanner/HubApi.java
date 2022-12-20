package tkexcelscanner;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
* <p>
* Klasse fuer das Bereitstellen von Daten fuer den MDM-HUB 
* Die Basis stellen die in der Klasse {@link ExcelParser} berechneten HashMaps dar
* </p>
* @version 1.0
* @author integration-factory
*/

public class HubApi {
	/** Liste bekannter Objekte */
	private List objectsList = new ArrayList();
	/** Liste bekannter Links */
	private List linksList = new ArrayList();
	/** Liste bekannter Attribute */
	private List attributesList = new ArrayList();
	/** Variable fuer Versorgungsarten */
	private String versorgungsart = "";
	/** Variable fuer Versorgungsverfahren */
	private String versorgungsverfahren = "";
	/** Indikator, ob attribute.csv bereits initialisiert wurde */
	private int attributeFileInit = 0;
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
	/** Variable fuer Instanz der Klasse Logging */
	private Logging xlsParseLogging;
	/** Variable fuer Instanz der Klasse Init */
	private Init init;

	
/**
* Klassenkonstruktor - bei der Instanziierung werden die Klasseninstanzen fuer {@link Logging} und {@link Init} uebergeben
*@param init Instanz der Klasse Init
*@param logging Instanz der Logging  
*/		
public HubApi(Init init,Logging logging) throws IOException{
	    this.xlsParseLogging = logging;
	    this.init = init;
}


/**
* Laden von Metadaten in die 4 HUB-Dateien objects.csv, links.csv , lineage.csv und attributes.csv auf Basis entsprechender HashMaps 
* @param  zusatzmappingHashMap HashMap fuer Zusatzmapping
* @param  versVerfahrenHashMap HashMap fuer Versorgungsverfahren
* @param  objektZuRessourceHashMap HashMap fuer ObjektzuRessource
* @param  dynSheetsHashMap HashMap fuer dynamische SheetArrays
*/
public void generateHubFiles (Map<Integer, Map<String, String>> zusatzmappingHashMap
		, Map<Integer, Map<String, String>> versVerfahrenHashMap
		,Map<Integer, Map<String, String>>  objektZuRessourceHashMap 
		,Map<Integer, Map<String, String>>  dynSheetsHashMap) throws IOException
{
	this.generateObjectsCsv(versVerfahrenHashMap, dynSheetsHashMap);
	this.generateLinksCsv(dynSheetsHashMap);
	this.generateLineageCsv(objektZuRessourceHashMap,dynSheetsHashMap,zusatzmappingHashMap);
	this.generateAttributeCsv(versVerfahrenHashMap,dynSheetsHashMap);
	
}

/**
* Laden von Metadaten in die HUB-Datei objects.csv auf Basis der HashMaps edcVersVerfahrenHashMap und dynSheetsHashMap 
*@param versVerfahrenHashMap HashMap fuer das Versorgungsverfahren
*@param dynSheetsHashMap HashMap fuer dynamische SheetArrays
*/
private void generateObjectsCsv(  Map<Integer, Map<String, String>> versVerfahrenHashMap
		, Map<Integer, Map<String, String>> dynSheetsHashMap) throws IOException
{
	 String targetDirectory = this.init.getParametervalue("zielverzeichnis");
	
	 File objectsCsv = new File(targetDirectory+"datastore_objects.csv");
	 FileOutputStream objectsCsvFos = new FileOutputStream(objectsCsv,true);
		 
	 BufferedWriter objectsCsvBw = new BufferedWriter(new OutputStreamWriter(objectsCsvFos));
		
	
	 int versVerfahrenHashMapSize = versVerfahrenHashMap.size();
	
	 String modelName = "com.ldm.custom.Custom_Model_Versorgungsart";
	 String kategorieDerInformation = "";
	 String artDerInformation = "";
	 String information = "";
	 
	 String objClass = "";
	 String objIdentity = "";
	 String objCoreName = "";
	 String objCoreDescription = "";
	 
	 this.versorgungsart = "";
	 this.versorgungsverfahren = "";
	 String fileRow = "";
	 String rowHash = "";
	 
	 for (int curr_pos = 1; curr_pos <= versVerfahrenHashMapSize; curr_pos ++)
	 {
		 
		 try {
			 kategorieDerInformation = versVerfahrenHashMap.get(curr_pos).get("Kategorie der Information");
			 artDerInformation = versVerfahrenHashMap.get(curr_pos).get("Art der Information");
			 information = versVerfahrenHashMap.get(curr_pos).get("Information");
			 
			 
		 }
		 catch(Exception e)
		 {
			 
		 }
		 
		 if (kategorieDerInformation.equals("Ressource_Hierarchie"))  
				 {
			 	  if (artDerInformation.equals("Versorgungsart"))
			 	  {
			   		this.versorgungsart=information;
			 	  }
			 	  else if (artDerInformation.equals("Name des Versorgungsverfahrens"))
			 	  {
			 		 this.versorgungsverfahren =  information;
			 	  }
			 
				 }
		    			 
	 }

	 if(this.versorgungsart != "" && this.versorgungsverfahren !="" )
	 {
		 objClass = modelName+".Versorgungsart";
		 objIdentity = "Versorgungsart/"+this.versorgungsart;
		 objCoreName = this.versorgungsart;
		 objCoreDescription = "";
		 
		 
		 fileRow = objClass+","+objIdentity+","+objCoreName+","+objCoreDescription;
		 
		 rowHash = MD5.getMd5(fileRow);
		 
		 if (!this.objectsList.contains(rowHash))
		 {
			 objectsCsvBw.write(fileRow);
			 objectsCsvBw.newLine();
			 this.objectsList.add(rowHash);
		 }
		 
		 objClass = modelName+".Versorgungsverfahren";
		 objIdentity = "Versorgungsart/"+this.versorgungsart+"/"+this.versorgungsverfahren;
		 objCoreName = this.versorgungsverfahren;
		 objCoreDescription = "";
		 
		 fileRow = objClass+","+objIdentity+","+objCoreName+","+objCoreDescription;
		 
		 rowHash = MD5.getMd5(fileRow);
		 
		 if (!this.objectsList.contains(rowHash))
		 {
			 objectsCsvBw.write(fileRow);
			 objectsCsvBw.newLine();
			 this.objectsList.add(rowHash);
			 
		 }
		 
		 String mappingTyp = "";
		 String mappingBeschr = "";
		 int dynSheetsHashMapSize = dynSheetsHashMap.size();
		 int mappingCount = 0;
		 
		 for (int curr_pos = 1; curr_pos <= dynSheetsHashMapSize; curr_pos ++)
		 {
			 mappingTyp = "";
			 mappingBeschr = "";
			 
			 try {
				 mappingTyp = dynSheetsHashMap.get(curr_pos).get("Mappingtyp");
				 mappingBeschr = dynSheetsHashMap.get(curr_pos).get("Mappingbeschreibung");
				 
				
				 
			 }
			 catch(Exception e)
			 {
				 
			 }
			 
			 
			 if ( mappingTyp != null && (mappingTyp.equals("Eins_zu_Eins") | mappingTyp.equals("Konstanter_Wert")| mappingTyp.equals("Transformation")))
			 {
				 
			    	 	 
			 
				 mappingCount++;
				 
				 objClass = modelName+".Mapping";
				 objIdentity = "Versorgungsart/"+this.versorgungsart+"/"+this.versorgungsverfahren+"/"+mappingTyp+"_"+mappingCount;
				 objCoreName = mappingTyp;
				 objCoreDescription = "";
				 
				
				 
				 fileRow = objClass+","+objIdentity+","+objCoreName+","+objCoreDescription;
				 
				 rowHash = MD5.getMd5(fileRow);
				 
				 if (!this.objectsList.contains(rowHash))
				 {
					 objectsCsvBw.write(fileRow);
					 objectsCsvBw.newLine();
					 this.objectsList.add(rowHash);
					 
				 }
			 }
			 
			 
			
			 
		 }
	 
	 }
	 
	 
	 objectsCsvBw.close();

}

/**
* Laden von Metadaten in die HUB-Datei links.csv auf Basis der HashMaps dynSheetsHashMap 
* @param dynSheetsHashMap HashMap fuer dynamische Sheet Arrays
*/
private void generateLinksCsv(   Map<Integer, Map<String, String>> dynSheetsHashMap) throws IOException
{
	 String targetDirectory = this.init.getParametervalue("zielverzeichnis");
		
	 File linksCsv = new File(targetDirectory+"datastore_links.csv");
	 FileOutputStream linksCsvFos = new FileOutputStream(linksCsv,true);
		 
	 BufferedWriter linksCsvBw = new BufferedWriter(new OutputStreamWriter(linksCsvFos));
		
	
	 	
	
	 String modelName = "com.ldm.custom.Custom_Model_Versorgungsart";
	 String kategorieDerInformation = "";
	 String artDerInformation = "";
	 String information = "";
	 
	 String linkAssociation = "";
	 String linkFromObjectIdentity = "";
	 String linkToObjectIdentity = "";
	 
	 String fileRow = "";
	 String rowHash = "";
		 
		 String mappingTyp = "";
		 String mappingBeschr = "";
		 String mapinngName ="";
		 
		 linkAssociation = modelName+".VersorgungsartToVersorgungsverfahren";
		 linkFromObjectIdentity = "Versorgungsart/"+this.versorgungsart;
		 linkToObjectIdentity = "Versorgungsart/"+this.versorgungsart+"/"+this.versorgungsverfahren;
		 
		 
		 fileRow = linkAssociation+","+linkFromObjectIdentity+","+linkToObjectIdentity;
			
		 
		 rowHash = MD5.getMd5(fileRow);
		 
		 if (!this.objectsList.contains(rowHash))
		 {
			 linksCsvBw.write(fileRow);
			 linksCsvBw.newLine();
			 this.linksList.add(rowHash);
			 
		 }
		 
		 
		 int dynSheetsHashMapSize = dynSheetsHashMap.size();
		 int mappingCount = 0;
		 
		 for (int curr_pos = 1; curr_pos <= dynSheetsHashMapSize; curr_pos ++)
		 {
			 mappingTyp = "";
			 mappingBeschr = "";
			 
			 try {
				 mappingTyp = dynSheetsHashMap.get(curr_pos).get("Mappingtyp");
				 mappingBeschr = dynSheetsHashMap.get(curr_pos).get("Mappingbeschreibung");
				 
						 
			 }
			 catch(Exception e)
			 {
				 
			 }
			 
			 
			 if ( mappingTyp != null) 
			 {
			 if (mappingTyp.equals("Eins_zu_Eins") | mappingTyp.equals("Konstanter_Wert")| mappingTyp.equals("Transformation"))
			 {
			 
				 mappingCount++;
				 
				 linkAssociation = modelName+".VersorgungsverfahrenToMapping";
				 linkFromObjectIdentity = "Versorgungsart/"+this.versorgungsart+"/"+this.versorgungsverfahren;
				 linkToObjectIdentity = "Versorgungsart/"+this.versorgungsart+"/"+this.versorgungsverfahren+"/"+mappingTyp+"_"+mappingCount;
				 
				 
				 fileRow = linkAssociation+","+linkFromObjectIdentity+","+linkToObjectIdentity;
					
				 
				 rowHash = MD5.getMd5(fileRow);
				 
				 if (!this.objectsList.contains(rowHash))
				 {
					 linksCsvBw.write(fileRow);
					 linksCsvBw.newLine();
					 this.linksList.add(rowHash);
					 
				 }
			 }
			 }
		 }

 linksCsvBw.close();

}


/**
* Laden von Metadaten in die HUB-Datei lineage.csv auf Basis der HashMaps objektZuRessourceHashMap, dynSheetsHashMap und zusatzmappingHashMap
* @param objektZuRessourceHashMap hashMap fuer ObjektzuRessource
* @param dynSheetsHashMap HashMap für dynamische Sheet Arrays
* @param zusatzmappingHashMap HashMap fuer Zusatzmappings
*/
private void generateLineageCsv( Map<Integer, Map<String, String>> objektZuRessourceHashMap
		,Map<Integer, Map<String, String>> dynSheetsHashMap
		, Map<Integer, Map<String, String>> zusatzmappingHashMap) throws IOException
{
	String targetDirectory = this.init.getParametervalue("zielverzeichnis");
		
		
	File lineageCsv = new File(targetDirectory+"datastore_lineage.csv");
	FileOutputStream lineageCsvFos = new FileOutputStream(lineageCsv,true);
		 
	BufferedWriter lineageCsvBw = new BufferedWriter(new OutputStreamWriter(lineageCsvFos));
		
	String lineageAssociation = "";
	String lineageFromConnection = "";
	String lineageToConnection = "";
	String lineageFromObject = "";
	String lineageToObject = "";
	
	String quellobjekt = "";
	String quellobjektRes = "";
	String quellattribut = "";
	String zielobjekt = "";
	String zielobjektRes = "";
	String zielattribut = "";
	
	 String mappingTyp = "";
	 String mapinngName ="";
	 String fileRow = "";
	 String rowHash= "";
	 String resObjekt = "";
	 String ressource = "";
	 
	 int dynSheetsHashMapSize = dynSheetsHashMap.size();
	 int edcObjektZuRessourceHashMapSize = objektZuRessourceHashMap.size();
	 
	 int mappingCount = 0;
	 
	 for (int curr_pos = 1; curr_pos <= dynSheetsHashMapSize; curr_pos ++)
	 {
		 mappingTyp = "";
		 
		 try {
			 mappingTyp = dynSheetsHashMap.get(curr_pos).get("Mappingtyp");
			 
				 
		 }
		 catch(Exception e)
		 {
			 
		 }
		 
		 
		 
		 if ( mappingTyp != null) 
		 {
		 if (mappingTyp.equals("Zusatzmapping") | mappingTyp.equals("Eins_zu_Eins") | mappingTyp.equals("Konstanter_Wert")| mappingTyp.equals("Transformation"))
		 {
			 if(!mappingTyp.equals("Zusatzmapping"))
			 {
			 mappingCount++;
			 }
			 
			 quellobjekt = dynSheetsHashMap.get(curr_pos).get("Quellobjekt").replaceAll("'", "").trim();
		     quellattribut = dynSheetsHashMap.get(curr_pos).get("Quellattribut").replaceAll("'", "").trim();;
			 zielobjekt = dynSheetsHashMap.get(curr_pos).get("Zielobjekt").replaceAll("'", "").trim();
			 zielattribut = dynSheetsHashMap.get(curr_pos).get("Zielattribut").replaceAll("'", "").trim();
			 
			 
			 for (int curr_objektRessource_pos = 1; curr_objektRessource_pos <= edcObjektZuRessourceHashMapSize; curr_objektRessource_pos ++)
			 {
				 resObjekt = objektZuRessourceHashMap.get(curr_objektRessource_pos).get("Objekt").replaceAll("'", "").trim();
				 ressource = objektZuRessourceHashMap.get(curr_objektRessource_pos).get("Ressource").replaceAll("'", "").trim();
				 
				 if (resObjekt.equals(quellobjekt))
				 {
					 quellobjektRes = ressource;
				 }
				 
				 if (resObjekt.equals(zielobjekt))
				 {
					 zielobjektRes = ressource;
				 }
			     			
			 }	 
			
			 if (!quellobjekt.equals("nicht vorhanden") && !quellattribut.equals("nicht vorhanden") && !mappingTyp.equals("Zusatzmapping"))
			 {
			 //Quelle zum Mapping
			 lineageAssociation = "core.DirectionalDataFlow";
			 lineageFromConnection = "";
			 lineageToConnection = "";
			 lineageFromObject = "<"+quellobjektRes+">/"+quellobjekt+"/"+quellattribut;
			 lineageToObject = "Versorgungsart://"+this.versorgungsart+"/"+this.versorgungsverfahren+"/"+mappingTyp+"_"+mappingCount;
			 quellobjektRes = "";
			 
			 fileRow = lineageAssociation+","+lineageFromConnection+","+lineageToConnection+","+lineageFromObject+","+lineageToObject;
				
			 
			 lineageCsvBw.write(fileRow);
			 lineageCsvBw.newLine();
			 }
			 
				
			 
			 if (!zielobjekt.equals("nicht vorhanden") && !zielattribut.equals("nicht vorhanden") && !mappingTyp.equals("Zusatzmapping"))
			 {
			 //Quelle zum Mapping
			 lineageAssociation = "core.DirectionalDataFlow";
			 lineageFromConnection = "";
			 lineageToConnection = "";
			 lineageFromObject = "Versorgungsart://"+this.versorgungsart+"/"+this.versorgungsverfahren+"/"+mappingTyp+"_"+mappingCount;
			 lineageToObject =  "<"+zielobjektRes+">/"+zielobjekt+"/"+zielattribut;
			 zielobjektRes = "";
			 
			 fileRow = lineageAssociation+","+lineageFromConnection+","+lineageToConnection+","+lineageFromObject+","+lineageToObject;
				
			 
			 lineageCsvBw.write(fileRow);
			 lineageCsvBw.newLine();
			 }
			
			 if (mappingTyp.equals("Zusatzmapping"))
			 {
			 //Quelle zum Mapping
			 lineageAssociation = "core.DirectionalDataFlow";
			 lineageFromConnection = "";
			 lineageToConnection = "";
			 
			 if (quellattribut =="" | quellattribut == null )
			 {
				 lineageFromObject = "<"+quellobjektRes+">/"+quellobjekt;
			 }
			 else
			 {
				 lineageFromObject = "<"+quellobjektRes+">/"+quellobjekt+"/"+quellattribut;
			 }
			
			 if (zielattribut =="" | zielattribut == null )
			 {
				 lineageToObject = "<"+zielobjektRes+">/"+zielobjekt;
			 }
			 else
			 {
				 lineageToObject = "<"+zielobjektRes+">/"+zielobjekt+"/"+zielattribut;
			 }
			
			 
			 zielobjektRes = "";
			 
			 fileRow = lineageAssociation+","+lineageFromConnection+","+lineageToConnection+","+lineageFromObject+","+lineageToObject;
				
			 
			 lineageCsvBw.write(fileRow);
			 lineageCsvBw.newLine();
			 }
			
			 
			
		 }
		 }
	 }
		
	 lineageCsvBw.close();


}



/**
* Laden von Metadaten in die HUB-Datei attributes.csv auf Basis der HashMaps versVerfahrenHashMap und dynSheetsHashMap 
@param versVerfahrenHashMap HashMap fuer das Versorgungsverfahren
@param dynSheetsHashMap hashMap fuer dynamische SheetArrays
*/
private void generateAttributeCsv( Map<Integer, Map<String, String>> versVerfahrenHashMap,
		Map<Integer, Map<String, String>> dynSheetsHashMap
		) throws IOException
{
	 


	 int versVerfahrenHashMapSize = versVerfahrenHashMap.size();
	 	
	 String targetDirectory = this.init.getParametervalue("zielverzeichnis");
			
	 String kategorieDerInformation = "";
	 String artDerInformation = "";
	 String information = "";
	 String attributCsvHeader = "id,core.name,core.classType";
	 String attributCsvHeaderDyn = "";
	 String attributCsvVersorgungsart = "";
	 String attributCsvVersorgungsverfahren = "";
	 String attributCsvMapping = "";

	 String attributBeschreibung = "";		 
	 String attributCsvValues = "";		 
	 String id ="";
	 String coreName ="";
	 String classType ="";
	 String fileRow = "";
	 String rowHash = "";
	 String [] informationParts;
	 String informationPart;
	 int metadatenWhitelistSize =  this.metadatenWhitelist.length;
	 String  informationRaw = "";
	 
	 
	 for (int i = 0; i< metadatenWhitelistSize; i++)
	 {
		 for (int curr_pos = 1; curr_pos <= versVerfahrenHashMapSize; curr_pos ++)
		 {
			 
			 information = "";
			 try {
				 kategorieDerInformation = versVerfahrenHashMap.get(curr_pos).get("Kategorie der Information");
				 artDerInformation = versVerfahrenHashMap.get(curr_pos).get("Art der Information");
				 
				 informationRaw = versVerfahrenHashMap.get(curr_pos).get("Information");
				 
				 informationParts = informationRaw.split("\n");
				 
				 
				 if (informationParts.length > 1)
				 {
					 for (int x = 0; x < informationParts.length;x++)
					 {
						 informationPart = informationParts[x];
						 if (informationPart == "" | informationPart == null)
						 {
							 informationPart = "<br>";
						 }
						 
						 information = information+"<p>"+informationPart+"</p>"; 
					 }
					 
				 }
				 else
				 {
				  information =  informationRaw.replace("\n", " ").replace(";", "#");
				 }
				 
				 information = information.replace(",", " ");
				
			 }
			 catch(Exception e)
			 {
				 
			 }
			 
			 
			 if ((kategorieDerInformation.equals("Metadaten_Einzelwert")  | kategorieDerInformation.equals("Metadaten_Liste")) 
					 && this.metadatenWhitelist[i].equals(kategorieDerInformation+"."+artDerInformation))
			 {
				 if(!artDerInformation.equals("Beschreibung"))
					{
				 	  attributCsvHeaderDyn = 	attributCsvHeaderDyn+","+artDerInformation;
				 	  attributCsvValues = attributCsvValues+","+information;
				 	  //Versorgungsart
				 	  if (artDerInformation.equals("Fachlicher Ansprechpartner") | artDerInformation.equals("Technischer Ansprechpartner"))
				 	  {
				 		  attributCsvVersorgungsart = attributCsvVersorgungsart+","+information;
				 	  }
				 	  else
				 	  {
				 		 attributCsvVersorgungsart = attributCsvVersorgungsart+",";
				 	  }
				 	 //Versorgungsverfahren
				 	 attributCsvVersorgungsverfahren = attributCsvVersorgungsverfahren+","+information;
				 	 
				 	//Mapping
				 	  if (!artDerInformation.equals("DataStore-Cluster") && !artDerInformation.equals("Information Assets"))
				 	  {
				 		 attributCsvMapping = attributCsvMapping+","+information;
				 	  }
				 	  else
				 	  {
				 		 attributCsvMapping = attributCsvMapping+",";
				 	  }
				 	  
					 }
				 else {
					 attributBeschreibung = information;
				 }
				 
				 
				 break;
				 
			 }  			 
		 }
	 }
	 if(this.attributeFileInit == 0)
	 {
		 
		 File attributesCsv = new File(targetDirectory+"datastore_attributes.csv");
	 	 FileOutputStream attributesCsvFos = new FileOutputStream(attributesCsv);
	 	 
	 	
	 	 BufferedWriter attributesBw = new BufferedWriter(new OutputStreamWriter(attributesCsvFos));
	 	 
	 	
	 	attributesBw.write(attributCsvHeader+attributCsvHeaderDyn+",Beschreibung");
	 	attributesBw.newLine();
	 	attributesBw.close();
		 
	 	this.attributeFileInit = 1;
	 	 
	 }
	
	 //Versorgungsart
	 id = "Versorgungsart://"+this.versorgungsart;
	 coreName = this.versorgungsart;
	 classType = "Versorgungsart";
	
	 fileRow = id+","+coreName+","+classType+attributCsvVersorgungsart+",";
	 
	 rowHash = MD5.getMd5(fileRow);

	 
	 if (!this.attributesList.contains(rowHash))
	 {
		 File attributesCsv = new File(targetDirectory+"datastore_attributes.csv");
	 	 FileOutputStream attributesCsvFos = new FileOutputStream(attributesCsv,true);
	 	 BufferedWriter attributesBw = new BufferedWriter(new OutputStreamWriter(attributesCsvFos));
		
	 	 attributesBw.write(fileRow);
		 attributesBw.newLine();
		 this.attributesList.add(rowHash);
		 
		 attributesBw.close();
		 rowHash = "";
	 }
	
	 //Versorgungsverfahren
	 id = "Versorgungsart://"+this.versorgungsart+"/"+this.versorgungsverfahren;
	 coreName = this.versorgungsverfahren;
	 classType = "Versorgungsverfahren";
	
	 //fileRow = id+","+coreName+","+classType+","+  attributCsvVersorgungsverfahren+","+attributBeschreibung;
	 fileRow = id+","+coreName+","+classType+  attributCsvVersorgungsverfahren+","+attributBeschreibung;
	 
	 rowHash = MD5.getMd5(fileRow);

	 
	 if (!this.attributesList.contains(rowHash))
	 {
		 File attributesCsv = new File(targetDirectory+"datastore_attributes.csv");
	 	 FileOutputStream attributesCsvFos = new FileOutputStream(attributesCsv,true);
	 	 BufferedWriter attributesBw = new BufferedWriter(new OutputStreamWriter(attributesCsvFos));
		
	 	 attributesBw.write(fileRow);
		 attributesBw.newLine();
		 this.attributesList.add(rowHash);
		 
		 attributesBw.close();
		 rowHash = "";
	 }
	 
	
	 //Mapping
	 
	 int dynSheetsHashMapSize = dynSheetsHashMap.size();
	 
	 String mappingTyp = "";
	 int mappingCount = 0;
	 String mappingBeschreibung = "";
	 String [] mappingBeschreibungParts;
	 String mappingBeschreibungPart;
	 String mappingBeschreibungRaw;
	 
	 for (int curr_pos = 1; curr_pos <= dynSheetsHashMapSize; curr_pos ++)
	 {
		 mappingTyp = "";
		 mappingBeschreibung = "";
		 
		 try {
			 mappingTyp = dynSheetsHashMap.get(curr_pos).get("Mappingtyp");
			 mappingBeschreibungRaw = dynSheetsHashMap.get(curr_pos).get("Mappingbeschreibung");
			 
			 
			 mappingBeschreibungParts = mappingBeschreibungRaw.split("\n");
			 
			 
			 
			 if (mappingBeschreibungParts.length > 1)
			 {
				 for (int x = 0; x < mappingBeschreibungParts.length;x++)
				 {
					 mappingBeschreibungPart = mappingBeschreibungParts[x];
					 if (mappingBeschreibungPart == "" | mappingBeschreibungPart == null)
					 {
						 mappingBeschreibungPart = "<br>";
					 }
					 
					 mappingBeschreibung = mappingBeschreibung+"<p>"+mappingBeschreibungPart+"</p>"; 
				 }
				 
			 }
			 else
			 {
				 mappingBeschreibung = mappingBeschreibungRaw.replace("\n", " ").replace(";", "#");
			 }
			 
			
			 mappingBeschreibung = mappingBeschreibung.replace(",", " ");
		 }
		 catch(Exception e)
		 {
			 mappingTyp = "";
			 mappingBeschreibung ="";
		 }
		 
		 
		 
		 if ( mappingTyp != null) 
		 {
		 if (mappingTyp.equals("Eins_zu_Eins") | mappingTyp.equals("Konstanter_Wert")| mappingTyp.equals("Transformation"))
		 {
			 
			 mappingCount++;
			 
			 id = "Versorgungsart://"+this.versorgungsart+"/"+this.versorgungsverfahren+"/"+mappingTyp+"_"+mappingCount;
			 coreName = mappingTyp+"_"+mappingCount;
			 classType = "Mapping";
			
			 fileRow = id+","+coreName+","+classType+  attributCsvMapping+","+mappingBeschreibung;
			 
			 File attributesCsv = new File(targetDirectory+"datastore_attributes.csv");
		 	 FileOutputStream attributesCsvFos = new FileOutputStream(attributesCsv,true);
		 	 BufferedWriter attributesBw = new BufferedWriter(new OutputStreamWriter(attributesCsvFos));
			
		 	 attributesBw.write(fileRow);
			 attributesBw.newLine();
			 
			 attributesBw.close(); 
				
			 
			
			 
			
		 }
		 }
	 }
	 
}

}




