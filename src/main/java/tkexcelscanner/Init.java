package tkexcelscanner;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.HashMap;
import java.util.Map;

/**
* <p>
* Klasse fuer die Initialisierung von Konfigurationsparametern und Files  
* </p>

* @version 1.0

* @author integration-factory
*/
public class Init {

/** Name des config-files */
private File config = new File("config.csv");
/** hashMap fuer die Ablage der Konfigurationsparameter als Key/Value Pairs */
private Map<Integer, Map<String, String>> configHashMap = new HashMap<>();


/**
 * Klassenkonstruktor - bei der Instanziierung werden Files initialisiert und Parameter geladen

 */
public Init() throws IOException{
	
	this.loadConfigParamsToHashMap();
	this.initializeHubFiles();
	

}

/**
 * Initialisiert die zu erzeugenden HUB-Files bei jedem Applikationsstart 
 */
private void initializeHubFiles() throws IOException  {
		
		 String objectsCSVHeader = "class,identity,core.name,core.description";
	     String linksCSVHeader = "association,fromObjectIdentity,toObjectIdentity";	
	     String lineageCSVHeader = "Association,FromConnection,ToConnection,FromObject,ToObject";	
	     String targetDirectory = this.getParametervalue("zielverzeichnis");
	    	
	     File objectsCsv = new File(targetDirectory+"datastore_objects.csv");
	 	 FileOutputStream objectsCsvFos = new FileOutputStream(objectsCsv);
	 	 
	 	 File linksCsv = new File(targetDirectory+"datastore_links.csv");
	 	 FileOutputStream linksCsvFos = new FileOutputStream(linksCsv);
	 	 
	 	 File lineageCsv = new File(targetDirectory+"datastore_lineage.csv");
	 	 FileOutputStream lineageCsvFos = new FileOutputStream(lineageCsv);
	 	 
	 	 BufferedWriter objectsBw = new BufferedWriter(new OutputStreamWriter(objectsCsvFos));
	 	 BufferedWriter linksBw = new BufferedWriter(new OutputStreamWriter(linksCsvFos));
	 	 BufferedWriter lineageBw = new BufferedWriter(new OutputStreamWriter(lineageCsvFos));
	 	
		 objectsBw.write(objectsCSVHeader);
		 objectsBw.newLine();
		 objectsBw.close();
		 
		 linksBw.write(linksCSVHeader);
		 linksBw.newLine();
		 linksBw.close();
	 	 
		 lineageBw.write(lineageCSVHeader);
		 lineageBw.newLine();
		 lineageBw.close();	
}	

/**
 * Laed die in der Datei config.csv definierten Applikationsparameter bei jedem Applikationsstart in eine Config-HashMap
 */
private  void loadConfigParamsToHashMap() throws IOException  {
	
	String line = null;
	
	
    Map<String, String> hashMapInsert = new HashMap<>();
	
	FileReader fileReader = new FileReader(this.config);
	
    BufferedReader bufferedReader = new BufferedReader(fileReader);
    
    int paramCounter = 0;
    String firstChar = "";
    
    while ((line = bufferedReader.readLine()) != null) {
    	
    	if(line.length()>0) { 
    	firstChar = line.substring(0,1);
    	}
    	else
    	{
    	firstChar = "";	
    	}
    	
    	if (!firstChar.equals("/") && !firstChar.equals("*")  && !firstChar.equals("")) {
    		String[] splitLine = line.split(",");
    		hashMapInsert.put("Parametername",  splitLine[0]);
    		this.configHashMap.put(paramCounter, hashMapInsert);
    		hashMapInsert.put("Parametervalue",  splitLine[1]);
    		this.configHashMap.put(paramCounter, hashMapInsert);    	
    		hashMapInsert = new HashMap<>();
    		paramCounter++;
    	}
    }
    bufferedReader.close();
}


/**
 * Get-Methode zur Abfrage eines bestimmten Parameters
 * @param parametername Name des Parameters fuer den der Wert geliefert werden soll 
 * @return Wert des Parameters
 */
public String getParametervalue(String parametername) throws IOException  {
	
	 int configHashMapSize = this.configHashMap.size();
	 
	 String parametervalue = "#NA";
	 for (int curr_pos = 0; curr_pos < configHashMapSize; curr_pos ++)
	 {
		 if (parametername.equals(this.configHashMap.get(curr_pos).get("Parametername")))
				 {
			 	 parametervalue = this.configHashMap.get(curr_pos).get("Parametervalue");
				 break;
				 }

		   			 
	 }
	return parametervalue;
}




}
