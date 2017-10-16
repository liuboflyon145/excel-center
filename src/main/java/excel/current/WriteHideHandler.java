package excel.current;

import excel.center.ExcelHide;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.concurrent.Callable;

/**
 * Created by liubo on 2017/8/14.
 */
public class WriteHideHandler implements Callable<String> {
    private final Logger log = LoggerFactory.getLogger(WriteHideHandler.class);
    private String exportFile;
    private List data;

    public WriteHideHandler(String exportFile, List data) {
        this.exportFile = exportFile;
        this.data = data;
    }

    @Override
    public String call() throws Exception {
        log.info("生成隐藏excel结果文件开始");
        long start = System.currentTimeMillis();
        String res = ExcelHide.writeHideExcel(data, exportFile);
        long end = System.currentTimeMillis()-start;
        log.info("生成隐藏excel结果文件结束，总耗时：{} 毫秒",end);
        return res;
    }
}
