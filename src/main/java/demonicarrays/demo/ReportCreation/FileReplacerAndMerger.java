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

        String orderDateYear = report.getOrderDate().substring(0,4);
        String orderDateMonth = report.getOrderDate().substring(5, 7);
        String orderDateDay = report.getOrderDate().substring(8, 10);
        String currentDateYear = report.getCurrentDate().substring(0,4);
        String currentDateMonth = report.getCurrentDate().substring(5,7);
        String currentDateDay = report.getCurrentDate().substring(8,10);

        List<String> students = ExcelParsing.pushToArrayList(ResultPusher.getCorrectPath(report.getFileChooser()));

        String stringOrderDateMonth = "";
        String stringCurrentDateMonth = "";

        switch (orderDateMonth){
            case "09" -> stringOrderDateMonth = "сентября";
            case "02" -> stringOrderDateMonth = "февраля";
        }

        switch (currentDateMonth){
            case "09" -> stringCurrentDateMonth = "сентября";
            case "02" -> stringCurrentDateMonth = "февраля";
        }

        if(stringOrderDateMonth.equals("сентября")){
            report.setSessionDate("декабря " +orderDateYear);
        }else{
            report.setSessionDate("мая " + orderDateYear);
        }


        report.setCurrentYear(orderDateYear);

        String convertedOrderDate = "«" + orderDateDay + "» " + stringOrderDateMonth + " " + orderDateYear + " г.";
        String convertedCurrentDate = "«" + currentDateDay + "» " + stringCurrentDateMonth + " " + currentDateYear + " г.";

        String brandNewSuperVisor;
        if(report.getSupervisorDegree() != "" && report.getSupervisorTitle() != ""){
            brandNewSuperVisor = report.getSupervisorFN() + ", " + report.getSupervisorDegree()
                    + ", " + report.getSupervisorTitle();
        }else{
            if(report.getSupervisorDegree() != ""){
                brandNewSuperVisor = report.getSupervisorFN() + ", " + report.getSupervisorDegree();
            } else if (report.getSupervisorTitle() != "") {
                brandNewSuperVisor = report.getSupervisorFN() + ", " + report.getSupervisorTitle();
            }else{
                brandNewSuperVisor = report.getSupervisorFN();
            }
        }

        String brandNewHeadOfDFN;
        if(report.getHeadOfDDegree() != "" && report.getHeadOfDTitle() != ""){
            brandNewHeadOfDFN = report.getHeadOfDFN() + ", " + report.getHeadOfDDegree()
                    + ", " + report.getHeadOfDTitle();
        }else{
            if(report.getHeadOfDDegree() != ""){
                brandNewHeadOfDFN = report.getHeadOfDFN() + ", " + report.getHeadOfDDegree();
            } else if (report.getHeadOfDTitle() != "") {
                brandNewHeadOfDFN = report.getHeadOfDFN() + ", " + report.getHeadOfDTitle();
            }else{
                brandNewHeadOfDFN = report.getHeadOfDFN();
            }
        }


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
            add("brandNewSuperVisor");
            add("brandNewHeadOfDFN");
//            if(report.getSupervisorTitle() != null) add("supervisorTitle");
//            if(report.getSupervisorDegree() != null) add("supervisorDegree");
//            if(report.getHeadOfDTitle() != null) add("headOfDTitle");
//            if(report.getHeadOfDDegree() != null) add("headOfDDegree");

            //add this line of codes if we'll have report for company
//            add("supervisorCompanyFN");
//            add("supervisorCompanyPosition");
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
                    case "orderDate" -> template.setField("orderDate", convertedOrderDate);
                    case "orderName" -> template.setField("orderName", report.getOrderName());
                    case "sessionDate" -> template.setField("sessionDate",report.getSessionDate());
                    case "supervisorFN" -> template.setField("supervisorFN",report.getSupervisorFN());
                    case "currentYear" -> template.setField("currentYear",report.getCurrentYear());
                    case "courseNum" -> template.setField("courseNum",report.getCourseNum());
                    case "groupName" -> template.setField("groupName", report.getGroupName());
                    case "practicePlaceAndTime" -> template.setField("practicePlaceAndTime",report.getPracticePlaceAndTime());
                    case "position" -> template.setField("position",report.getPosition());
                    case "currentDate" -> template.setField("currentDate", convertedCurrentDate);
                    case "headOfDFN" -> template.setField("headOfDFN",report.getHeadOfDFN());
                    case "directionName" -> template.setField("directionName", report.getDirectionName());
                    case "profileName" -> template.setField("profileName", report.getProfileName());
                    case "supervisorPosition" -> template.setField("supervisorPosition", report.getSupervisorPosition());
                    case "brandNewSuperVisor" -> template.setField("brandNewSuperVisor", brandNewSuperVisor);
                    case "brandNewHeadOfDFN" -> template.setField("brandNewHeadOfDFN", brandNewHeadOfDFN);
//                    case "supervisorTitle" -> template.setField("supervisorTitle", report.getSupervisorTitle());
//                    case "supervisorDegree" -> template.setField("supervisorDegree", report.getSupervisorDegree());
//                    case "headOfDTitle" -> template.setField("headOfDTitle", report.getHeadOfDTitle());
//                    case "headOfDDegree" -> template.setField("headOfDDegree", report.getHeadOfDDegree());
                    //add this line of codes if we'll have report for company
//                    case "supervisorCompanyFN" -> template.setField("supervisorCompanyFN", report.getSupervisorCompanyFN());
//                    case "supervisorCompanyPosition" -> template.setField("supervisorCompanyPosition", report.getSupervisorCompanyPosition());
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
