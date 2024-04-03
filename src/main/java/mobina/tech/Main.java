package mobina.tech;

public class Main {
    public static void main(String[] args) throws Exception {
        String documentPath = "D:\\sampleWord\\apache.docx";

        //to create word document
        if (ExistenceCheckerOfDocument.isDocumentExists(documentPath)) {
            System.out.println("The Document is already available");
        } else {
            DocumentCreator.createWordDocument(documentPath);
            System.out.println("Document created successfully!");
        }

        //to edit word document
        if (ExistenceCheckerOfDocument.isDocumentExists(documentPath)) {
//            DocumentEditor.accessToTableCells(documentPath);
            DocumentEditor.addTableToDocument(documentPath);
            DocumentEditor.addImage(documentPath);
            System.out.println("Document edited successfully!");
        } else {
            System.out.println("The Document is not available!");
        }
    }
}