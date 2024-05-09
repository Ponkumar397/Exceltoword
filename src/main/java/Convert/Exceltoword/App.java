package Convert.Exceltoword;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.*;



import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws IOException, InvalidFormatException, org.apache.poi.openxml4j.exceptions.InvalidFormatException
    {
    	 XWPFDocument doc = new XWPFDocument();

    	  // the body content
    	  XWPFParagraph paragraph = doc.createParagraph();
    	  XWPFRun run = paragraph.createRun();  
    	  run.setText("The Body...");

    	  // create header
    	  XWPFHeader header = doc.createHeader(HeaderFooterType.DEFAULT);

    	  // header's first paragraph
    	  paragraph = header.getParagraphArray(0);
    	  if (paragraph == null) paragraph = header.createParagraph();
    	  paragraph.setAlignment(ParagraphAlignment.CENTER);

    	  run = paragraph.createRun();
          String inputimg="C:\\Users\\PonkumarE\\Documents\\Test\\header.jpeg";
    	  FileInputStream in = new FileInputStream(inputimg);
    	  run.addPicture(in, Document.PICTURE_TYPE_JPEG, inputimg, Units.toEMU(100), Units.toEMU(50));
    	  in.close();  

    	  run.setText("HEADER"); 

    	  FileOutputStream out = new FileOutputStream("C:\\Users\\PonkumarE\\Documents\\Test\\output\\CreateWordHeaderWithImage.docx");
    	  doc.write(out);
    	  doc.close();
    	  out.close();
    }  
}

