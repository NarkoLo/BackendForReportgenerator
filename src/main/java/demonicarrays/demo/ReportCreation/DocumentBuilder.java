package demonicarrays.demo.ReportCreation;

import lombok.SneakyThrows;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.core.io.ClassPathResource;
import org.springframework.util.ResourceUtils;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

public class DocumentBuilder {
    private final Template template;
    private final Path output;
    private final Path input;


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

    @SneakyThrows
    public DocumentBuilder(Template template, Path output, Path input){
        this.template = template;
        this.output = output;
        this.input = input;
        setXWPFDocument(Files.newInputStream(this.input));


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
        try (OutputStream out = Files.newOutputStream(output)) {
            doc.write(out);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    public static void clearDoc(Path outFile){
        try (OutputStream out = Files.newOutputStream(outFile)) {
            out.write(0);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    static void deleteWM(Path path, String WM) {
        try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(path)); OutputStream out = Files.newOutputStream(path))
                //Files.newInputStream(Paths.get(path)));
             {
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
