package demonicarrays.demo.ReportCreation;


import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import demonicarrays.demo.UniversityProject.Entity.Report;
import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

public class FileReplacerAndMerger {
    //change path
    private final  static Path pathToTitleList = TempFileUtility.createTempFile("wordSource/TitleLists.docx");
    private final static Path inputFile = TempFileUtility.createTempFile("wordSource/mainWordFile.docx");
    private final static Path outputFile = TempFileUtility.createTempFile("wordSource/fileForTesting.docx");

    public static void fileReplacerAndMerger(Report report) throws IOException {

        List<String> students = ExcelParsing.pushToArrayList(ResultPusher.getCorrectPath(report.getFileChooser()));

        ArrayList<String> replaceableNames = new ArrayList<>(){{
            add("studentFullName");
            add("studentFN");
            add("instituteName");
            add("departmentName");
            add("practiceName");
            add("orderDate");
            add("orderName");
            add("sessionDate");
            add("supervisorFN");
            add("currentYear");
            add("courseNum");
            add("groupName");
            add("practicePlaceAndTime");
            add("position");
            add("currentDate");
            add("headOfDFN");
            add("directionName");
            add("profileName");
        }};

        DocumentBuilder.clearDoc(pathToTitleList);
        for (String studentName: students) {
            Template template = new Template(replaceableNames, studentName);
            for (String replaceableName : replaceableNames) {
                switch (replaceableName) {
                    case "instituteName" -> template.setField("instituteName", report.getInstituteName());//objUpdateWord.updateDocument(inputPath, outputPath, "${instituteName}", (String) instituteName.getValue());
                    case "departmentName" -> template.setField("departmentName",report.getDepartmentName());//objUpdateWord.updateDocument(inputPath, outputPath, "${departmentName}", (String) departmentName.getValue());
                    case "practiceName" -> template.setField("practiceName",report.getPracticeName());//objUpdateWord.updateDocument(inputPath, outputPath, "${practiceName}", (String) practiceName.getValue());
                    case "orderDate" -> template.setField("orderDate",report.getOrderDate());//objUpdateWord.updateDocument(inputPath, outputPath, "${orderDate}", String.valueOf(orderDate.getValue()));
                    case "orderName" -> template.setField("orderName", report.getOrderName());//objUpdateWord.updateDocument(inputPath, outputPath, "${orderName}", orderName.getText());
                    case "sessionDate" -> template.setField("sessionDate",report.getSessionDate());//objUpdateWord.updateDocument(inputPath, outputPath, "${sessionDate}", String.valueOf(sessionDate.getValue()));
                    case "supervisorFN" -> template.setField("supervisorFN",report.getSupervisorFN());//objUpdateWord.updateDocument(inputPath, outputPath, "${supervisorFN}", supervisorFN.getText());
                    case "currentYear" -> template.setField("currentYear",report.getCurrentYear());//objUpdateWord.updateDocument(inputPath, outputPath, "${currentYear}", String.valueOf(currentYear));
                    case "courseNum" -> template.setField("courseNum",report.getCourseNum());//objUpdateWord.updateDocument(inputPath, outputPath, "${courseNum}", (String) courseNum.getValue());
                    case "groupName" -> template.setField("groupName", report.getGroupName());//objUpdateWord.updateDocument(inputPath, outputPath, "${groupName}", groupName.getText());
                    case "practicePlaceAndTime" -> template.setField("practicePlaceAndTime",report.getPracticePlaceAndTime());//objUpdateWord.updateDocument(inputPath, outputPath, "${practicePlaceAndTime}", practicePlaceAndTime.getText());
                    case "position" -> template.setField("position",report.getPosition());//objUpdateWord.updateDocument(inputPath, outputPath, "${position}", position.getText());
                    case "currentDate" -> template.setField("currentDate",report.getCurrentDate());//objUpdateWord.updateDocument(inputPath, outputPath, "${currentDate}", String.valueOf(currentDate.getValue()));
                    case "headOfDFN" -> template.setField("headOfDFN",report.getHeadOfDFN()); //objUpdateWord.updateDocument(inputPath, outputPath, "${headOfDFN}", headOfDFN.getText());
                    case "directionName" -> template.setField("directionName", report.getDirectionName());//objUpdateWord.updateDocument(inputPath, outputPath, "${directionName}", (String) directionName.getValue());
                    case "profileName" -> template.setField("profileName", report.getProfileName());//objUpdateWord.updateDocument(inputPath, outputPath, "${profileName}", profileName.getText());
                    default -> {
                    }
                }
            }
            template.createMap();

            DocumentBuilder documentBuilder = new DocumentBuilder(template, outputFile, inputFile);
            documentBuilder.buildDoc();
            documentBuilder.saveDoc();
            Document document = new Document(pathToTitleList.toString());
            document.insertTextFromFile(outputFile.toString(), FileFormat.Docx_2013);
            document.saveToFile(pathToTitleList.toString(), FileFormat.Docx_2013);

            DocumentBuilder.clearDoc(outputFile);
        }
        DocumentBuilder.deleteWM(pathToTitleList,"Evaluation Warning: The document was created with Spire.Doc for JAVA.");
        ResultPusher.pushFile(pathToTitleList, report.getGroupName() + ".docx");
        DocumentBuilder.clearDoc(pathToTitleList);
    }
}
