package demonicarrays.demo.UniversityProject.Controler;

import demonicarrays.demo.UniversityProject.Entity.Report;
import demonicarrays.demo.UniversityProject.Service.ReportService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.util.List;

@CrossOrigin
@RestController
@RequestMapping("/api/test")
public class ReportController {

    @Autowired
    private ReportService service;

    @PostMapping
    public Report addReport(@RequestBody Report report) throws IOException {
        return service.saveReport(report);
    }

    @GetMapping
    public List<Report> findAllReports(){
        return service.getReport();
    }

}
