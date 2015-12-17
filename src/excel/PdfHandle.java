package excel;

import java.io.IOException;

import org.pdfbox.exceptions.COSVisitorException;
import org.pdfbox.exceptions.CryptographyException;
import org.pdfbox.pdmodel.PDDocument;
import org.pdfbox.pdmodel.encryption.PDStandardEncryption;

/**
 * @author Administrator 给pdf加密
 * @description
 *2014-11-19  下午12:21:45
 */
public class PdfHandle {
    public static void pdfEncrypt(String filepath,String password){
        PDDocument pdf;
	try {
	    pdf = PDDocument.load(filepath);
	    //create the encryption options
	    PDStandardEncryption encryptionOptions =new PDStandardEncryption();
	    encryptionOptions.setCanPrint( true );
	    pdf.setEncryptionDictionary(encryptionOptions );
	    //encrypt the document
	    pdf.encrypt( "master", password );
	    System.out.println("isEncrypted : " + pdf.isEncrypted());
	    //save the encrypted documentto the file system
	    pdf.save(filepath);
	    pdf.close();
	} catch (IOException e) {
	    e.printStackTrace();
	} catch (CryptographyException e) {
	    e.printStackTrace();
	} catch (COSVisitorException e) {
	    e.printStackTrace();
	}
    }
    
}
