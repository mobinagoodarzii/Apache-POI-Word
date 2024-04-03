package mobina.tech;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileOutputStream;


public class DocumentCreator {
    public static void createWordDocument(String documentPath) throws Exception {
        // Create a new document
        XWPFDocument doc = new XWPFDocument();

        XWPFParagraph paragraph = doc.createParagraph();
        XWPFRun run = paragraph.createRun();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        run.setText("ساخته شده توسط Apache POI");
        run.setFontFamily("IranNastaliq"); // Set font for Persian text
        run.getCTR().getRPr().addNewRtl().setVal(true); // Set RTL for Persian text

        // Add a line break
        paragraph.createRun().addBreak();


        // Save the document in DOCX format
        try (FileOutputStream out = new FileOutputStream(documentPath)) {
            doc.write(out);
            doc.close();
        }
    }
}