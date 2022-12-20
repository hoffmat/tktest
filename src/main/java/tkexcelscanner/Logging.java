package tkexcelscanner;

import java.io.FileWriter;   // Import the FileWriter class
import java.io.IOException;  // Import the IOException class to handle errors
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

/**

* <p>
* Klasse fuer das Schreiben von Logging-Informationen 
*</p>
* <p>
* Initialisiert und schreibt Logging-Informationen in die Datei datastore_logging.csv
*</p>
*
* @version 1.0
* @author integration-factory
*/
public class Logging {
	
	/** Filewriter fuer das Log-File */
	private FileWriter logWriter;
	/** Header fuer das Log-File */
	private String loggingHeader = "Zeitpunkt,Excel_File,Reiter,Zeilennummer,Logtyp,Information";
	/** Status der angibt, ob das Log-File bereits initialisiert wurde */
	private int headerCreated = 0;
	/** Variable für Instanz der Klasse Init */
	private Init init;
	/** HUB-Zielverzeichnis */
	String targetDirectory;
	
	
/**
* Klassenkonstruktor - bei der Instanziierung wird das Logging-File initialisiert 
* @param init Instanz der Klasse Init
* @throws IOException 
*/	
public Logging(Init init) throws IOException {
      
	  this.init = init;
	  
	  this.targetDirectory = this.init.getParametervalue("zielverzeichnis");
	 	
	  this.logWriter = new FileWriter(this.targetDirectory+"datastore_logging.csv");
	  this.writeLog(this.loggingHeader);
	  this.headerCreated = 1;
		  
	 
  }
	  
	

/**
* Methode fuer das Schreiben von Logging-Informationen in die Logging-Datei 
* @param logMessage Meldung, die ins Log-File geschrieben werden soll
*/	
public void writeLog(String logMessage) {
		 DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
	     String zeitpunkt = dtf.format(LocalDateTime.now());
	     String cnvLogMessage;
	     if (this.headerCreated == 0) {
	    	 cnvLogMessage = this.loggingHeader;
	     }
	     else {
	    	 cnvLogMessage = zeitpunkt+","+logMessage;
	     }
	    
		
	    try {
	      this.logWriter.write(cnvLogMessage+"\n");
	   
	    } catch (IOException e) {
	      System.out.println("An error occurred.");
	      e.printStackTrace();
	    }
	  }

/**
* Schliessen des Log-Files
*/	
public void closeLogFile() throws IOException {
			
		  this.logWriter.close();
		    
		  }
}
