package demonicarrays.demo.ReportCreation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;

public class ResultPusher {
    static void pushFile(String path, String fileName){
        Path file = Paths.get(path);
        String pathToDesktop = makeCorrectPath(path);
        try {
            Files.copy(file, Paths.get(pathToDesktop + "\\" + fileName),REPLACE_EXISTING);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }


    }
    static String getCorrectPath(String pathToExcelFile){
        String subString = "C:\\fakepath";

        return pathToExcelFile.replace(subString, makeCorrectPath(""));
    }
    private static String makeCorrectPath(String path){

        path = System.getProperty(("user.home"))  + "\\OneDrive\\Desktop";
        if(!isCorrectPath(path)){
            path = path.replace("Desktop", "Рабочий Стол");
            if(!isCorrectPath(path))
                path = path.replace("\\OneDrive","");
            if(!isCorrectPath(path))
                path = path.replace("Рабочий Стол","Desktop");
        }
        return path;
    }
    static private boolean isCorrectPath(String pathToDesktop){
        return Files.exists(Paths.get(pathToDesktop));
    }
}
