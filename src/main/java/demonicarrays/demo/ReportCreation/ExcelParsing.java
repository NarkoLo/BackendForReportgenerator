package demonicarrays.demo.ReportCreation;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelParsing {
    public static List<String> pushToArrayList(String path) throws IOException {
        //create new workbook
        try(FileInputStream file = new FileInputStream(path)){
            Workbook workbook = new XSSFWorkbook(file);

            Sheet sheet =  workbook.getSheetAt(0);

            //initialize arrayList with list of Students
            List<String> students = new ArrayList<>();

            //parsing excel table
            for(Row row : sheet){
                for(Cell cell : row){
                    if (cell.getCellType() == CellType.STRING ){
                        //add students to arrayList
                        students.add(cell.getRichStringCellValue().getString());
                    }
                }
            }
            return students;
        }
        catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return null;
    }
}
