import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.TrueFileFilter;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Collection;
import java.util.List;

/**
 * SmartReplacer replaces strings within documents in a folder and its subfolders
 * @author Vlado Repic
 * @author CYP Association
 * @version 29.01.2019
 */
public class SmartReplacer {

    /* Search Folders */
    // Directory Path
    public static String PATH = "C:/Users/Vlado Repic/Desktop/Demo";

    // File Search
    public Collection<File> searchFiles(final File directory, boolean filtera, boolean filterb) {
        return FileUtils.listFilesAndDirs(directory, TrueFileFilter.TRUE, TrueFileFilter.TRUE);
    }

    // Main method: Search for files with ending .doc or .docx
    public static void main(String[] args) throws Exception {

        Collection<File> documents = new SmartReplacer().searchFiles(new File(SmartReplacer.PATH), true, true);
        for (File document: documents) {
            if (!document.getName().contains(".")) {
                System.out.println("\n| " + document.getAbsolutePath() + "\n-----------------");
            }
            if (document.getName().toLowerCase().endsWith(".doc") || document.getName().toLowerCase().endsWith(".docx")) {
                String documentPath = document.getAbsolutePath().replace("\\", "/");
                replaceInFile(documentPath);
                System.out.println(" |- " + document.getName() + " - " + document.length());
            }
        }

    }

    // Search & replace within file
    private static void replaceInFile(String documentPath) {

        try{
            // Current document
            XWPFDocument docx = new XWPFDocument(new FileInputStream(documentPath));

            // Search for
            String target = "P:\\test";
            //TODO: Hyperlinks are not being replaced yet, this has to be done with XWPFHyerlinkRun
            String fileTarget = "file:///" + target.toLowerCase().replace("/", "\\");

            // Replace with
            String replacement = "H:\\test\\test2";
            //TODO: Hyperlinks are not being replaced yet, this has to be done with XWPFHyerlinkRun
            String fileReplacement = "file:///" + replacement.toLowerCase().replace("/", "\\");

            // Search & replace within paragraphs
            for (XWPFParagraph p : docx.getParagraphs()) {
                List<XWPFRun> runs = p.getRuns();
                if (runs != null) {
                    for (XWPFRun r : runs) {
                        replace(r, target, replacement);
                        replace(r, fileTarget, fileReplacement);
                        System.out.println(fileTarget + "..." + fileReplacement);
                    }
                }
            }
            // Search & replace within tables
            for (XWPFTable tbl : docx.getTables()) {
                for (XWPFTableRow row : tbl.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            for (XWPFRun r : p.getRuns()) {
                                replace(r, target, replacement);
                                replace(r, fileTarget, fileReplacement.replace("/", "\\"));
                            }
                        }
                    }
                }
            }

            // Save changed file as original file (overwrite)
            docx.write(new FileOutputStream(documentPath));

        }catch (Exception e){
            System.out.println(e);
        }

    }

    // Replace method
    private static void replace(XWPFRun r, String target, String replacement) {

        String text = r.getText(0);
        if (text != null && text.contains(target)) {
            text = text.replace(target, replacement);
            r.setText(text, 0);
        }
    }
}
