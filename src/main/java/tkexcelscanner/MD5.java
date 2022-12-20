package tkexcelscanner;

import java.math.BigInteger;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;

/**

* <p>
* Hilfsklasse fuer die Berechnung von MD5-Hashkeys  
* </p>

* @version 1.0

* @author integration-factory

*/
public class MD5 {
	/**
	 * Erechnet einen MD5-Hashkey aus einem Input-String
	 * @param inputstring Wert fuer den der MD5 berechnet werden soll 
	 * @return hashtext 
	 */
	public static String getMd5(String inputstring)
	{
		try {

			// Static getInstance method is called with hashing MD5
			MessageDigest md = MessageDigest.getInstance("MD5");

			// digest() method is called to calculate message digest
			// of an input digest() return array of byte
			byte[] messageDigest = md.digest(inputstring.getBytes());

			// Convert byte array into signum representation
			BigInteger no = new BigInteger(1, messageDigest);

			// Convert message digest into hex value
			String hashtext = no.toString(16);
			while (hashtext.length() < 32) {
				hashtext = "0" + hashtext;
			}
			return hashtext;
		}

		// For specifying wrong message digest algorithms
		catch (NoSuchAlgorithmException e) {
			throw new RuntimeException(e);
		}
	}
}