package excel.center;


import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.write.*;

import jxl.write.Number;

import java.io.File;
import java.util.*;


/**
 * Created by liubo on 2017/8/6.
 */
public class ExcelHide {
    public static final String BLANK = "";

    /**
     * 处理过的数据生成最终excel
     *
     * @param sourceData
     */
    public static String writeHideExcel(List sourceData, String fileName) {
        try {

            File writeFile = new File(fileName);
            //  打开文件
            WritableWorkbook book = Workbook.createWorkbook(writeFile);
            //  生成名为“第一页”的工作表，参数0表示这是第一页
            WritableSheet sheet = book.createSheet("sheet1", 0);
            //  在Label对象的构造子中指名单元格位置是第一列第一行(0,0)
//            设置表头字体
            WritableFont font = new WritableFont(WritableFont.TIMES, 16, WritableFont.BOLD);
            WritableCellFormat format = new WritableCellFormat(font);
            format.setAlignment(Alignment.getAlignment(2));

            WritableFont bodyFont = new WritableFont(WritableFont.TIMES, 13, WritableFont.NO_BOLD);
            WritableCellFormat bodyFormat = new WritableCellFormat(bodyFont);
            bodyFormat.setAlignment(Alignment.getAlignment(1));
            out:
            for (int i = 0; i < sourceData.size(); i++) {
                List row = (List) sourceData.get(i);
                inner:
                for (int j = 0; j < row.size(); j++) {
                    Object obj = row.get(j);
                    if (obj instanceof String) {
                        if (((String) obj).startsWith("容积表")) {
                            sheet.mergeCells(0, i, 10, i);
                            Label label = new Label(j, i, obj.toString(), format);
                            sheet.addCell(label);
                            continue out;
                        }
                        if (((String) obj).startsWith("客户名称")) {
                            sheet.mergeCells(1, i, 5, i);
                            sheet.mergeCells(6, i, 10, i);
                            Label label1 = new Label(j, i, obj.toString(), bodyFormat);
                            Label label2 = new Label(1, i, row.get(1).toString(), bodyFormat);
                            sheet.addCell(label1);
                            sheet.addCell(label2);
                            continue out;
                        }
                        if (((String) obj).startsWith("单位")) {
                            sheet.mergeCells(10, i, 11, i);
                            Label label1 = new Label(j, i, obj.toString(), bodyFormat);
                            sheet.addCell(label1);
                            continue out;
                        }
                        if (((String) obj).startsWith("---")) {
                            sheet.mergeCells(0, i, 10, i);
                            Label label1 = new Label(j, i, "-------------------------------------------------------------------------------------------------------------------------------------", bodyFormat);
                            sheet.addCell(label1);
                            continue out;
                        }
                        Label label = new Label(j, i, obj.toString(), bodyFormat);
                        sheet.addCell(label);
                        continue;
                    }
                    if (obj instanceof Integer) {
                        Number number = new Number(j, i, ((Integer) obj).longValue(), bodyFormat);
                        sheet.addCell(number);
                        continue;
                    }
                }
            }

            //  写入数据并关闭文件
            book.write();
            book.close();

        } catch (Exception e) {
            e.getStackTrace();
            System.out.println(e);
        }
        return fileName;
    }

    /**
     * excel文件数据读取，传入文件绝对路径
     *
     * @param filePath
     * @return
     */
    public static List readHideExcel(String filePath) {
        List result = new ArrayList();
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
                            val = BLANK;
                        }
                        if ("液  高".equals(val)) {
                            val = BLANK;
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
                            lst.set(j, BLANK);
                    }
                    System.out.println("标定:" + i + "      " + lst.toString());
                }
            }
            System.out.println("删除结果：" + result.toString());
        } catch (Exception e) {
            System.out.println(e);
        }
        return result;
    }

}
