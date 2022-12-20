package tkexcelscanner;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**

* <p>
* Hilfsklasse fuer das Erstellen einer File-Liste  
* </p>

* @version 1.0

* @author integration-factory

*/

public class FileList {
	/** Verzeichnis der Excel-Files*/
	private String directory;
	/** File-Liste als ArrayList*/
	private List<Path> fileList = new ArrayList<>();



/**
 * Klassenkonstruktor - bei der Instanziierung wird das Dateiverzeichnis fuer die Excel-Files uebergeben
 * @param directory Quellverzeichnis für DataStore-Excels
 */	 
 public FileList(String directory){
        this.directory = directory;
 }
    
   

 /**
  * Erzeugt eine File-Liste fuer alle Excel-Files im angegebenen Verzeichnis und gibt sie zurueck
  * @return fileList
  */	   
 public List<Path> getFileList() throws IOException{
    	   
    	try (Stream<Path> stream = Files.walk(Paths.get(directory))) {
        // Do something with the stream.
        this.fileList = stream.map(Path::normalize)
          .filter(Files::isRegularFile)
          .filter(path -> path.getFileName().toString().endsWith(".xlsx"))
          .collect(Collectors.toList());
      }
      return this.fileList;
  }
}
  