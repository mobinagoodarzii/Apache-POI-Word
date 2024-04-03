package mobina.tech;

import org.apache.poi.xwpf.usermodel.*;


public class Tables {
    public static void formattedTable(XWPFDocument document) throws Exception {
        // Create a table
        XWPFTable table = document.createTable();
        table.setWidth("100%");

        // Create the first row
        XWPFTableRow tableRowOne = table.getRow(0);
        // Create a separate run for each cell
        XWPFRun runCell1 = tableRowOne.getCell(0).addParagraph().createRun();
        XWPFRun runCell2 = tableRowOne.addNewTableCell().addParagraph().createRun();
        XWPFRun runCell3 = tableRowOne.addNewTableCell().addParagraph().createRun();

        // Set text and properties for each cell
        runCell1.setText("شماره (Number)");
        runCell1.setFontFamily("IranNastaliq");
        runCell1.getCTR().getRPr().addNewRtl().setVal(true);
        runCell1.getParagraph().setAlignment(ParagraphAlignment.CENTER);
        tableRowOne.getCell(0).setColor("7e8b9e");

        runCell2.setText("نام(Name) , نام خانوادگی(FamilyName)");
        runCell2.setFontFamily("IranNastaliq");
        runCell2.getCTR().getRPr().addNewRtl().setVal(true);
        runCell2.getParagraph().setAlignment(ParagraphAlignment.CENTER);
        tableRowOne.getCell(1).setColor("7e8b9e");

        runCell3.setText("سن (Age)");
        runCell3.setFontFamily("IranNastaliq");
        runCell3.getCTR().getRPr().addNewRtl().setVal(true);
        runCell3.getParagraph().setAlignment(ParagraphAlignment.CENTER);
        tableRowOne.getCell(2).setColor("7e8b9e");

        // Create the second row
        XWPFTableRow tableRowTwo = table.createRow();
        XWPFRun runCell21 = tableRowTwo.getCell(0).addParagraph().createRun();
        XWPFRun runCell22 = tableRowTwo.getCell(1).addParagraph().createRun();
        XWPFRun runCell23 = tableRowTwo.getCell(2).addParagraph().createRun();
        runCell21.setText("1");
        runCell21.getParagraph().setAlignment(ParagraphAlignment.CENTER);

        runCell22.setText(" نام name مورد نظر یافت نشد!");
        runCell22.setFontFamily("IranNastaliq");
        runCell22.getCTR().getRPr().addNewRtl().setVal(true);
        runCell22.getParagraph().setAlignment(ParagraphAlignment.CENTER);

        runCell23.setText("--");
        runCell23.getParagraph().setAlignment(ParagraphAlignment.CENTER);


        // Create the third row
        XWPFTableRow tableRowThree = table.createRow();
        XWPFRun runCell31 = tableRowThree.getCell(0).addParagraph().createRun();
        XWPFRun runCell32 = tableRowThree.getCell(1).addParagraph().createRun();
        XWPFRun runCell33 = tableRowThree.getCell(2).addParagraph().createRun();
        runCell31.setText("2");
        runCell31.getParagraph().setAlignment(ParagraphAlignment.CENTER);

        runCell32.setText("مبینا Mobina گودرزی Goodarzi");
        runCell32.setFontFamily("IranNastaliq");
        runCell32.getCTR().getRPr().addNewRtl().setVal(true);
        runCell32.getParagraph().setAlignment(ParagraphAlignment.CENTER);

        runCell33.setText("23");
        runCell33.getParagraph().setAlignment(ParagraphAlignment.CENTER);
    }
}
