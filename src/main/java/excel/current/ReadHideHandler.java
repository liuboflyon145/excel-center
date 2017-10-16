package excel.current;

import excel.utils.Constants;
import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Callable;

/**
 * Created by liubo on 2017/8/14.
 */
public class ReadHideHandler implements Callable<List> {
    private final Logger log = LoggerFactory.getLogger(ReadHideHandler.class);
    private String filePath ;

    public ReadHideHandler(String filePath) {
        this.filePath = filePath;
    }

    @Override
    public List call() throws Exception {
        List result = new ArrayList();
        log.info("读取隐藏内容excel源文件开始");
        long start = System.currentTimeMillis();
        try {
            File file = new File(filePath);
//            BufferedInputStream input = new BufferedInputStream(new ByteArrayInputStream(filePath));
            Workbook book = Workbook.getWorkbook(file);
            //  获得第一个工作表对象
            Sheet sheet = book.getSheet(0);

            int rows = sheet.getRows();
            for (int i = 0; i < rows; i++) {
                List cellList = new ArrayList();
                Cell[] cells = sheet.getRow(i);
                for (int j = 0; j < cells.length; j++) {
                    Cell cell = cells[j];
                    CellType type = cell.getType();
                    if (type == CellType.LABEL) {
                        String val = cell.getContents();
                        if (val.startsWith("卧罐")) {
                            val = val.replace("卧罐", "");
                        }
                        if ("单 位:".equals(val)) {
                            val = "客户名称:";
                        }
                        if (val.startsWith("标定单位:")) {
                            val = Constants.BLANK;
                        }
                        if ("液  高".equals(val)) {
                            val = Constants.BLANK;
                        }
                        cellList.add(j, val);
                        continue;
                    }
                    if (type == CellType.NUMBER) {
                        cellList.add(j, Integer.parseInt(cell.getContents()));
                        continue;
                    }
                    cellList.add(j, cell.getContents());

                }
                result.add(cellList);
//                添加证书编号
                List tmp = (List) result.get(result.size() - 1);
                if (tmp.size() > 0 && tmp.get(0).toString().startsWith("罐")) {
                    List noList = new ArrayList();
                    noList.add("证书编号:");
                    result.add(noList);
                }
            }
            book.close();
            System.out.println("替换结果：" + result.toString());
            int nums = result.size();
//            删除多余信息
            for (int i = result.size() - 1; i >= 0; i--) {
                List lst = (List) result.get(i);
                if (i > nums - 11)
                    result.remove(i);
                if (lst.size() != 0 && "标定:".equals(lst.get(0))) {
                    for (int j = 0; j < lst.size(); j++) {
                        if (j != 6)
                            lst.set(j, Constants.BLANK);
                    }
                    System.out.println("标定:" + i + "      " + lst.toString());
                }
            }
            System.out.println("删除结果：" + result.toString());
        } catch (Exception e) {
            System.out.println(e);
        }finally {
            long end = System.currentTimeMillis();
            log.info("读取隐藏内容源文件结束，总耗时：{} 毫秒",(end-start));
        }
        return result;
    }
}
