package excel;

import java.io.FileOutputStream;
import java.io.IOException;

import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;

/**
 * @author Administrator
 * @description
 *2014-11-19  上午11:13:30
 */
public class CutPdf {
    /** 
     * 截取pdf文档的第一页 
     * @param sourceFile 源文件 
     * @param targetFile 目标文件 
     * @param ranges   复制规则 
     */ 
  public static void copyPdf(String sourceFile ,String targetFile){ 
     PdfReader pdfReader;
     PdfStamper pdfStamper = null;
    try {
	pdfReader = new PdfReader(sourceFile);
	pdfStamper = new PdfStamper(pdfReader , new FileOutputStream(targetFile)); 
	pdfReader.selectPages("1"); 
    } catch (IOException e) {
	e.printStackTrace();
    } catch (DocumentException e) {
	e.printStackTrace();
    } 
    finally{
	try {
	    pdfStamper.close();
	} catch (DocumentException e) {
	    e.printStackTrace();
	} catch (IOException e) {
	    e.printStackTrace();
	} 
    }
    
  } 
  
}
