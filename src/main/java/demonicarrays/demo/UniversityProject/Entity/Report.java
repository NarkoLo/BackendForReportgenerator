package demonicarrays.demo.UniversityProject.Entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Calendar;
import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class Report {

    private int ver;
    private String instituteName;
    private String departmentName;
    private String practiceName;
    private String orderDate;
    private String orderName;
    private String currentYear = "";
    private String sessionDate = "";
    private String supervisorFN;
    private String courseNum;
    private String groupName;
    private String practicePlaceAndTime;

    private String position;

    private String currentDate;
    private String headOfDFN;
    private String directionName;
    private String profileName;
    private String fileChooser;

    private String supervisorTitle;
    private String supervisorDegree;
    private String supervisorPosition;
    private String headOfDTitle;
    private String headOfDDegree;
    private String supervisorCompanyFN;
    private String supervisorCompanyPosition;



}
