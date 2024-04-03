package mobina.tech;

import org.apache.commons.io.FileUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class DocumentEditor {
    public static void accessToTableCells(String documentPath) throws Exception {
        XWPFDocument doc = new XWPFDocument(new FileInputStream(documentPath));

        XWPFTable table = doc.getTableArray(0);

        // Access the first cell in the table
        XWPFTableCell cell = table.getRow(0).getCell(0);

        // Remove all content from the cell
        cell.removeParagraph(0);

        // Add new content to the cell
        XWPFParagraph paragraph = cell.addParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("edited cell");
        run.setColor("870418"); // Set color to blue

        // Save the document
        try (FileOutputStream out = new FileOutputStream(documentPath)) {
            doc.write(out);
        }
    }

    public static void addTableToDocument(String documentPath) throws Exception {
        XWPFDocument doc = new XWPFDocument(new FileInputStream(documentPath));
        Tables.formattedTable(doc);

        // Save the document
        try (FileOutputStream out = new FileOutputStream(documentPath)) {
            doc.write(out);
        }
    }

    public static void addImage(String documentPath) throws Exception {
        XWPFDocument doc = new XWPFDocument(new FileInputStream(documentPath));

        // Create a document builder to add content to the document
        XWPFParagraph paragraph = doc.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("Adding an image to the document:");

        // Add an image to the document
        byte[] imageBytes = FileUtils.readFileToByteArray(new File("E:\\pictures\\animations\\mango.jpg"));
        int format = XWPFDocument.PICTURE_TYPE_JPEG;
        int width = Units.toEMU(200); // Width of the image
        int height = Units.toEMU(200); // Height of the image
        doc.addPictureData(imageBytes, format);
        doc.createParagraph().createRun().addPicture(new ByteArrayInputStream(imageBytes), format, "mango.jpg", width, height);

        // Save the document
        try (FileOutputStream out = new FileOutputStream(documentPath)) {
            doc.write(out);
        }
    }
}


