package demonicarrays.demo.UniversityProject.Service;

import demonicarrays.demo.UniversityProject.Entity.Report;
import demonicarrays.demo.UniversityProject.Repository.ReportRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.List;

@Service
public class ReportService {

    @Autowired
    private ReportRepository repository;

    public Report saveReport(Report report) throws IOException {
        return repository.save(report);
    }

    public List<Report> getReport(){
        return repository.getAllReports();
    }
}
