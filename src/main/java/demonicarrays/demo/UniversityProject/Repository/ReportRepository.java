package demonicarrays.demo.UniversityProject.Repository;


import demonicarrays.demo.ReportCreation.FileReplacerAndMerger;
import demonicarrays.demo.UniversityProject.Entity.Report;
import org.springframework.stereotype.Repository;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Repository
public class ReportRepository {
    private List<Report> list = new ArrayList<Report>();


    public void createReport(){
        list = List.of();
    }

    public List<Report> getAllReports(){
        return list;
    }
    public Report save(Report newReport) throws IOException {
        list.add(newReport);
        System.out.println(newReport.getCurrentYear());
        FileReplacerAndMerger.fileReplacerAndMerger(newReport);

        return newReport;
    }
}
