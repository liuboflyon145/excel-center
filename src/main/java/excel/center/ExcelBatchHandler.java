package excel.center;

import jxl.HeaderFooter;
import jxl.SheetSettings;
import jxl.Workbook;
import jxl.format.*;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.ResultSet;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

/**
 * Created by liubo on 2017/8/14.
 */
public class ExcelBatchHandler {
    private static  ExecutorService service = Executors.newFixedThreadPool(5);

    public static Future<String> workerCenter(List<String> fileName){
//        service.submit()
        return null;
    }


}
