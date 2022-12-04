package demonicarrays.demo.ReportCreation;

import org.apache.poi.xwpf.usermodel.*;
import org.springframework.core.io.ClassPathResource;
import org.springframework.util.ResourceUtils;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

public class DocumentBuilder {
    private final Template template;
    private final File output;
    private final File input;


    private XWPFDocument doc = null;

    private XWPFDocument getDoc() {
        return doc;
    }

    private void setXWPFDocument(InputStream input){
        try {
            doc =  new XWPFDocument(input);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public DocumentBuilder(Template template, File output,File input) throws FileNotFoundException {
        this.template = template;
        this.output = output;
        this.input = input;
        setXWPFDocument(new FileInputStream(this.input));


    }
    private void stringSearch(XWPFDocument doc){
        for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    for (String oldReplacement : template.getReplaceableNames()) {
                        if (text != null && text.contains(oldReplacement)) {
                            text = text.replace(oldReplacement, template.getKeysAndValues().get(oldReplacement));
                            r.setText(text, 0);
                        }

                    }

                }
            }
        }
    }
    private void tableSearch(XWPFDocument doc){
        for (XWPFTable tbl : doc.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            for (String oldReplacement : template.getReplaceableNames()) {
                                if (text != null && text.contains(oldReplacement)) {
                                    text = text.replace(oldReplacement, template.getKeysAndValues().get(oldReplacement));
                                    r.setText(text, 0);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    public void buildDoc(){
        XWPFDocument doc = getDoc();
        stringSearch(doc);
        tableSearch(doc);
    }
    public void saveDoc(){
        try (FileOutputStream out = new FileOutputStream(output)) {
            doc.write(out);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    public static void clearDoc(File outFile){
        try (FileOutputStream out = new FileOutputStream(outFile)) {
            out.write(0);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    static void deleteWM(String path, String WM) {
        try (XWPFDocument doc = new XWPFDocument(new ClassPathResource(path).getInputStream());
                //Files.newInputStream(Paths.get(path)));
             FileOutputStream out = new FileOutputStream(path)) {
            for (XWPFParagraph paragraph : doc.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                if (runs != null) {
                    for (XWPFRun r : runs) {
                        String text = r.getText(0);
                        if (text != null && text.contains(WM)) {
                            text = text.replace(WM, "");
                            r.setText(text, 0);
                        }
                    }
                }
            }
            doc.write(out);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
