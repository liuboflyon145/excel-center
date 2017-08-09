package excel.center;


import excel.utils.Utils;
import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.read.biff.BiffException;
import jxl.write.*;

import jxl.write.Number;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;


/**
 * Created by liubo on 2017/8/6.
 */
public class ExcelTools {
    public static final String BLANK = "";

    private static String customerName = "";
    private static String numbers = "";
    private static String expireDate = "";
    private static List commentList = new ArrayList();


    public static void main(String args[]) {
        System.out.println(args[0]);
        System.out.println(args[1]);
//        String excelFile = "/Users/liubo/workspace/nodepros/excel/text.xls";
//        List result = readHideExcel(excelFile);
//        writeHideExcel(result);
//
//        String dataFile = "/Users/liubo/workspace/nodepros/excel/data.xls";
//        List convert = readSourceExcel(dataFile);
//        writeConvertExcel(convert);
    }

    /**
     * 生成转置后数据excel
     *
     * @param sourceData
     */
    public static void writeConvertExcel(List sourceData,String fileName) {
//        生成随机数
        String dateStr  = Utils.getDate();
        File writeFile = new File(fileName);

//        打开生成excel文件
        try {
            WritableWorkbook book = Workbook.createWorkbook(writeFile);
            //  生成名为“第一页”的工作表，参数0表示这是第一页
            WritableSheet sheet = book.createSheet("sheet1", 0);
            //  在Label对象的构造子中指名单元格位置是第一列第一行(0,0)
//            设置表头字体
            WritableFont headerFont = new WritableFont(WritableFont.TIMES, 16, WritableFont.BOLD);
            WritableCellFormat headerFormat = new WritableCellFormat(headerFont);
            headerFormat.setAlignment(Alignment.getAlignment(2));
//设置标题字体
            WritableFont bodyFont = new WritableFont(WritableFont.TIMES, 13, WritableFont.NO_BOLD);
            WritableCellFormat bodyFormat = new WritableCellFormat(bodyFont);
            bodyFormat.setAlignment(Alignment.getAlignment(2));

            stepTableHeader(0, sheet, headerFormat, bodyFormat);
//            设置表头内容
            int offset = 1;
            for (int i = 0; i < 3; i++) {
                sheet.addCell(new Label(i + offset, 3, "实高", bodyFormat));
                sheet.addCell(new Label(i + 1 + offset, 3, "容量", bodyFormat));
                sheet.addCell(new Label(i + offset, 4, "(cm)", bodyFormat));
                sheet.addCell(new Label(i + 1 + offset, 4, "(kL)", bodyFormat));
                offset++;
            }

            //第一步 数据分片
            stepGroupOne(sourceData, sheet, headerFormat, bodyFormat);

            book.write();
            book.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }


    /**
     * 数据分片
     *
     * @param sourceData
     * @param sheet
     * @param headerFormat
     * @param bodyFormat
     */
    private static void stepGroupOne(List sourceData, WritableSheet sheet, WritableCellFormat headerFormat, WritableCellFormat bodyFormat) throws WriteException {
//            表体数据填充
        double step1 = 12;
        double stepTotal = Math.ceil(sourceData.size() / step1);
        int stepOffSet1 = 0;


        for (int i = 0; i < stepTotal; i++) {
            int from = stepOffSet1;
            int to = from + 12;
            if (to > sourceData.size()) {
                to = sourceData.size();
            }
            List sub = sourceData.subList(from, to);
            stepGroupTwo(sub, sheet, sheet.getRows(), bodyFormat);

            stepOffSet1 += 12;
            System.out.println("row " + sheet.getRows());
            stepTailInfo(sheet.getRows(), sheet, headerFormat, bodyFormat, i + 1);
            if (i < stepTotal - 1) {
                stepTableHeader(sheet.getRows() + 1, sheet, headerFormat, bodyFormat);
            }
            if(i==stepTotal-1){
                stepTableTailComment(sheet.getRows(),sheet);
            }
        }
    }

    /**
     * 添加尾页说明信息
     * @param rows
     * @param sheet
     * @throws WriteException
     */
    private static void stepTableTailComment(int rows, WritableSheet sheet) throws WriteException {
        sheet.addCell(new Label(1,rows+1,commentList.get(0).toString()));
        sheet.mergeCells(2,rows+1,6,rows+1);
        sheet.addCell(new Label(2,rows+1,commentList.get(1).toString()));
        sheet.mergeCells(2,rows+2,6,rows+2);
        sheet.addCell(new Label(2,rows+2,commentList.get(2).toString()));
    }

    /**
     * 设置页尾分页信息
     *
     * @param row
     * @param sheet
     * @param headerFormat
     * @param bodyFormat
     * @param page
     */
    private static void stepTailInfo(int row, WritableSheet sheet, WritableCellFormat headerFormat, WritableCellFormat bodyFormat, int page) throws WriteException {
        sheet.addCell(new Label(1, row + 1, "有效日期:", bodyFormat));
        sheet.mergeCells(2, row + 1, 5, row + 1);
        sheet.addCell(new Label(2, row + 1, "2017年7月27日至2021年7月26日", bodyFormat));
        sheet.addCell(new Label(6, row + 1, String.format("第%d页", page), bodyFormat));
    }

    /**
     * 设置分片表头信息
     *
     * @param rowNum
     * @param sheet
     * @param headerFormat
     * @param bodyFormat
     * @throws WriteException
     */
    private static void stepTableHeader(int rowNum, WritableSheet sheet, WritableCellFormat headerFormat, WritableCellFormat bodyFormat) throws WriteException {
        //            设置表头第一行内容
        sheet.mergeCells(1, rowNum, 6, rowNum);
        sheet.addCell(new Label(1, rowNum, "卧式金属罐容量表", headerFormat));

//            设置第二行内容
        sheet.addCell(new Label(1, rowNum + 1, "客户名称:", bodyFormat));
        sheet.mergeCells(2, rowNum + 1, 6, rowNum + 1);
        sheet.addCell(new Label(2, rowNum + 1, "中国石油四川成都龙泉驿区北干道站", bodyFormat));

//            设置第三行内容
        sheet.addCell(new Label(1, rowNum + 2, "罐号:", bodyFormat));
        sheet.mergeCells(2, rowNum + 2, 3, rowNum + 2);
        sheet.addCell(new Label(2, rowNum + 2, "1"));
    }

    /**
     * 数据分片填充
     *
     * @param sub
     * @param sheet
     * @param startRow
     * @param bodyFormat
     * @throws WriteException
     */
    private static void stepGroupTwo(List sub, WritableSheet sheet, int startRow, WritableCellFormat bodyFormat) throws WriteException {
        int colNum = 1;
        int subOffSet = 0;
        double step = 4;
        double total = Math.ceil(sub.size() / step);

        for (int i = 0; i < total; i++) {
            int from = subOffSet;
            int to = from + 4;
            if (to > sub.size()) {
                to = sub.size();
            }
            List sub1 = sub.subList(from, to);
            recursiveFill(sub1, sheet, startRow, colNum, bodyFormat);
            subOffSet += 4;
            colNum += 2;


        }
    }

    /**
     * 行列转置数据填充
     *
     * @param sourceData
     * @param sheet
     * @param rowNum
     * @param col
     * @param bodyFormat
     * @throws WriteException
     */
    private static void recursiveFill(List sourceData, WritableSheet sheet, int rowNum, int col, WritableCellFormat bodyFormat) throws WriteException {
        for (int i = 0; i < sourceData.size(); i++) {
            List row1 = (List) sourceData.get(i);
//                行数据转置为列数据
            for (int j = 0; j < row1.size(); j++) {
                if (j % 2 == 0) {
                    sheet.addCell(new Label(col, rowNum, row1.get(j).toString(), bodyFormat));
                } else {
                    sheet.addCell(new Label(col + 1, rowNum, row1.get(j).toString(), bodyFormat));
                    rowNum++;
                }
            }
            rowNum++;
        }

    }

    /**
     * 读取转置的原始数据
     *
     * @param dataFile
     * @return
     */
    public static List readSourceExcel(String dataFile) {
        File file = new File(dataFile);

        List convertList = new ArrayList();
        try {
//            读取excel
            Workbook book = Workbook.getWorkbook(file);
//            获取页签
            Sheet sheet = book.getSheet(0);
//            获取总行数
            int rowTotal = sheet.getRows();
            Cell[] c1 = sheet.getRow(rowTotal-11);
            for (int i=0;i<c1.length;i++){
                String val = c1[i].getContents();
                if (val!= null&&val.length()>0) {
                    commentList.add(val);
                }
            }
            Cell[] c2 = sheet.getRow(rowTotal-10);
            for (int i=0;i<c2.length;i++){
                String val = c2[i].getContents();
                if (val!= null&&val.length()>0) {
                    commentList.add(val);
                }
            }
            System.out.println(commentList.toString());
            List result = new ArrayList();
            for (int i = 0; i < rowTotal; i++) {
                if (i <= 9 || i >= rowTotal - 14) {
                    continue;
                }
                List row = new ArrayList();
//                获取一行数据
                Cell[] cols = sheet.getRow(i);
                for (int j = 0; j < cols.length; j++) {
                    Cell cell = cols[j];
                    CellType type = cell.getType();
                    if (type == CellType.LABEL) {
                        String val = cell.getContents();
                        if (val.startsWith("卧罐")) {
                            numbers = cols[j + 1].getContents();
                        }
                        if ("单 位:".equals(val)) {
                            customerName = cols[j + 1].getContents();
                        }
                        if (val.startsWith("有效期:")) {
                            expireDate = val;
                        }
                    }
                    if (type == CellType.NUMBER) {
                        row.add(j, Integer.parseInt(cell.getContents()));
                    }
                }
                result.add(row);
            }
            book.close();
            System.out.println("原始数据结果：" + result.toString());

            for (int i = 0; i < result.size(); i++) {
//                获取元数据的一行
                List row = (List) result.get(i);
                if (row.size() == 0)
                    continue;
                Integer index = (Integer) row.get(0);
//                新生成一个数据行
                List cntList = new ArrayList();
                for (int j = 0; j < row.size() - 1; j++) {
                    cntList.add(index++);
                    Integer rs = (Integer) row.get(j + 1);
                    Double r1 = Double.valueOf(rs.doubleValue() / 1000);
                    cntList.add(r1);
                }
                convertList.add(cntList);
            }
            System.out.println("转置后的数据：" + convertList.toString());
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return convertList;
    }

    /**
     * 处理过的数据生成最终excel
     *
     * @param sourceData
     */
    public static void writeHideExcel(List sourceData,String fileName) {
        try {
            String dateStr = Utils.getDate();
//            String.format("/Users/liubo/workspace/nodepros/excel/test%d.xls", dateStr)
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
                            Label label1 = new Label(j, i, obj.toString(), bodyFormat);
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
