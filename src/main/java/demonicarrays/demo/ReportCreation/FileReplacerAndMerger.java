package demonicarrays.demo.ReportCreation;


import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import demonicarrays.demo.UniversityProject.Entity.Report;
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

            add("supervisorPosition");
            add("supervisorTitle");
            add("supervisorDegree");
            add("headOfDTitle");
            add("headOfDDegree");
            add("supervisorCompanyFN");
            add("supervisorCompanyPosition");
        }};

        DocumentBuilder.clearDoc(pathToTitleList);
        assert students != null;
        for (String studentName: students) {
            Template template = new Template(replaceableNames, studentName);
            for (String replaceableName : replaceableNames) {
                switch (replaceableName) {
                    case "instituteName" -> template.setField("instituteName", report.getInstituteName());
                    case "departmentName" -> template.setField("departmentName",report.getDepartmentName());
                    case "practiceName" -> template.setField("practiceName",report.getPracticeName());
                    case "orderDate" -> template.setField("orderDate",report.getOrderDate());
                    case "orderName" -> template.setField("orderName", report.getOrderName());
                    case "sessionDate" -> template.setField("sessionDate",report.getSessionDate());
                    case "supervisorFN" -> template.setField("supervisorFN",report.getSupervisorFN());
                    case "currentYear" -> template.setField("currentYear",report.getCurrentYear());
                    case "courseNum" -> template.setField("courseNum",report.getCourseNum());
                    case "groupName" -> template.setField("groupName", report.getGroupName());
                    case "practicePlaceAndTime" -> template.setField("practicePlaceAndTime",report.getPracticePlaceAndTime());
                    case "position" -> template.setField("position",report.getPosition());
                    case "currentDate" -> template.setField("currentDate",report.getCurrentDate());
                    case "headOfDFN" -> template.setField("headOfDFN",report.getHeadOfDFN());
                    case "directionName" -> template.setField("directionName", report.getDirectionName());
                    case "profileName" -> template.setField("profileName", report.getProfileName());
                    case "supervisorPosition" -> template.setField("supervisorPosition", report.getSupervisorPosition());
                    case "supervisorTitle" -> template.setField("supervisorTitle", report.getSupervisorTitle());
                    case "supervisorDegree" -> template.setField("supervisorDegree", report.getSupervisorDegree());
                    case "headOfDTitle" -> template.setField("headOfDTitle", report.getHeadOfDTitle());
                    case "headOfDDegree" -> template.setField("headOfDDegree", report.getHeadOfDDegree());
                    case "supervisorCompanyFN" -> template.setField("supervisorCompanyFN", report.getSupervisorCompanyFN());
                    case "supervisorCompanyPosition" -> template.setField("supervisorCompanyPosition", report.getSupervisorCompanyPosition());
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
        System.out.println(report.getSupervisorTitle());
    }
}
