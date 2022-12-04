package demonicarrays.demo.ReportCreation;

import org.apache.xmlbeans.impl.common.ResolverUtil;
import org.springframework.core.io.ClassPathResource;
import org.springframework.util.ResourceUtils;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

class TempFileUtility {
    static File  createTempFile(String path){
        Path tempDir = Paths.get(System.getProperty("user.dir") + "\\temp");
        try{
            if(!Files.exists(tempDir))
                Files.createDirectories(tempDir);
            File temp = new File(tempDir, path.replaceAll("",""))
            OutputStream outputStream = new FileOutputStream(temp)

        } catch (IOException e) {
                throw new RuntimeException(e);
        }


        try {
            temp = new File(tempDir, new ClassPathResource(path).getInputStream());
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return temp;
    }
}
