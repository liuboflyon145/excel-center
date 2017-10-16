package excel.current;

import excel.utils.Constants;
import excel.utils.Utils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.concurrent.Callable;

/**
 * Created by liubo on 2017/8/14.
 */
public class UploadHandler implements Callable<String> {

    private final Logger log = LoggerFactory.getLogger(UploadHandler.class);
    private MultipartFile file;

    private final String dest = "temp/upload";

    public UploadHandler(MultipartFile file) {
        this.file = file;
    }

    @Override
    public String call() throws Exception {
        long start = System.currentTimeMillis();

        log.info("读取任务开始");
        String dateStr = Utils.getDate(Utils.DateFormatter.DAY_FORMATTER);
        String timeStr = Utils.getDate(Utils.DateFormatter.TIME_FORMATTER);

        String fileName = dest + Constants.SYSTEM_SEPARATOR + dateStr + Constants.SYSTEM_SEPARATOR + timeStr + file.getOriginalFilename();
        Utils.checkFile(fileName);
        File uploadFile = new File(fileName);
        byte[] bytes = file.getBytes();
        try (BufferedOutputStream stream = new BufferedOutputStream(new FileOutputStream(uploadFile))) {
            stream.write(bytes);
            stream.close();
            long end = System.currentTimeMillis();
            log.info("读取任务结束,任务处理耗时：{} 毫秒",(end-start));
        }
        try(BufferedInputStream input = new BufferedInputStream(new ByteArrayInputStream(bytes))){

        }
        return fileName;
    }
}
