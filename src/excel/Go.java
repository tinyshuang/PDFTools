package excel;

/**
 * @author Administrator
 * @description
 *2015-12-17  下午10:23:34
 */
public class Go {
    public static void main(String[] args) {
	ToPdf.excel2pdf("d:\\test.xlsx","d:\\test.pdf");  
	CutPdf.copyPdf("d:\\test.pdf", "d:\\test01.pdf");
	PdfHandle.pdfEncrypt( "d:\\test01.pdf".replaceAll("\\\\", "//"), "123456");
    }
}
