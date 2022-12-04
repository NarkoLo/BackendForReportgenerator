package demonicarrays.demo.ReportCreation;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class StringConverter {
    public static List<String> cutStud(List<String> students){
        List<String> newStudentsFormat = new ArrayList<>();
        for (String newFIOFormat:students) {
            StringBuilder latter = new StringBuilder();
            Matcher m = Pattern.compile("\\s\\W").matcher(newFIOFormat);
            while (m.find()) {
                latter.append(m.group() +".");
            }
            newStudentsFormat.add(newFIOFormat.replaceAll("\\s\\W+", String.valueOf(latter)));
        }
        return newStudentsFormat;
    }
    public static String cutStud(String student) {
        StringBuilder latter = new StringBuilder();
        Matcher m = Pattern.compile("\\s\\W").matcher(student);
        while (m.find()) {
            latter.append(m.group() + ".");
        }
        return student.replaceAll("\\s\\W+", String.valueOf(latter));
    }
}
