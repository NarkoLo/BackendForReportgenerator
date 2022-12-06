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
    Date currentDateInstance = new Date();
    Calendar cal = Calendar.getInstance();

    int currentYearInstance;
    int currentMonthInstance = cal.get(Calendar.MONTH);

    private int ver;
    private String instituteName;
    private String departmentName;
    private String practiceName;
    private String orderDate;
    private String orderName;
    private String currentYear = String.valueOf(currentYearInstance = cal.get(Calendar.YEAR));
    private String sessionDate = (currentMonthInstance > 6 ? "декабря " + currentYear : "июня " + currentYear);
    private String supervisorFN;
    private String courseNum;
    private String groupName;
    private String practicePlaceAndTime;

    //this shit is not changing at all
    private String position;

    private String currentDate;
    //this shit is not changing in last page of report of every student
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
