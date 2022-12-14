package demonicarrays.demo.ReportCreation;

import org.apache.poi.util.IOUtils;
import org.springframework.core.io.ClassPathResource;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;


class TempFileUtility {
    private static final Path tempDir = Paths.get(System.getProperty("java.io.tmpdir") + "/UniversityProjectTemplates");
    static Path  createTempFile(String path){
        Path filePath;
        if(!Files.exists(tempDir)) {
            try {
                Files.createDirectories(tempDir);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
        String fileName = path.replaceAll("wordSource/","");
        if(!Files.exists(tempDir.resolve(fileName))) {
            try(InputStream inputStream = new ClassPathResource(path).getInputStream())
            {
                filePath = Files.createFile(tempDir.resolve(fileName));
                IOUtils.copy(inputStream,Files.newOutputStream(tempDir.resolve(fileName)));
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
        else filePath = tempDir.resolve(fileName);
        return filePath;
    }
}
