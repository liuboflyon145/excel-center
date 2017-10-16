package excel.center;


import jxl.read.biff.BiffException;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class PoiExcelConvert {
    private static List commentList = new ArrayList();
    private static String customerName = "";
    private static String numbers = "";
    private static String expireDate = "";

//    public static void main(String[] args) {
//        try {
//            List data = readSourceExcel("/Users/liubo/workspace/javapros/excel-center/temp/upload/20170918/092529/中国石油四川成都郫县华浦站4号罐.xls");
//            writeConvertExcel(data, "");
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }

    public static void writeConvertExcel(List sourceData, String fileName) throws IOException {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("sheet1");
        dataSplitToPage(sourceData, sheet,wb);
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            wb.write(fileOut);
            fileOut.close();
        }
    }

    private static void dataSplitToPage(List sourceData, Sheet sheet, Workbook wb) {
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
//            数据12组分组
            List sub = sourceData.subList(from, to);

            System.out.println("rows " + sheet.getLastRowNum());
            int start = sheet.getLastRowNum();
            if (start > 0) {
                start += 7;
            }

//            2 数据渲染
            renderPageTemplate(sub, sheet, i,wb);
            if (i == stepTotal - 1) {
                setFinishedFlag(sheet, sub.size(), i,wb);
            }
            stepOffSet1 += 12;
//            设置表尾
            stepTailInfo(sheet.getLastRowNum() + 1, sheet, i + 1,wb);
//
            if (i == stepTotal - 1) {
                stepTableTailComment(sheet.getLastRowNum() + 1, sheet,wb);
            }
        }
    }


    private static void setFinishedFlag(Sheet sheet, int size, int page, Workbook wb) {
        int start = page * 56 + 6;
        int col = 0;
        int rowNum = 0;


        if (size == 1 || size == 5 || size == 9) {
            rowNum = start + 10;
        } else if (size == 2 || size == 6 || size == 10) {
            rowNum = start + 20 + 1;
        } else if (size == 3 || size == 7 || size == 11) {
            rowNum = start + 30 + 2;
        } else if (size == 4 || size == 8) {
            rowNum = start;
        } else {
            rowNum = start + 43;
        }
        if (size >= 1 && size < 4) {
            col = 1;
        } else if (size >= 4 && size < 8) {
            col = 3;
        } else if (size >= 8 && size < 12) {
            col = 5;
        } else {
            col = 1;
        }

        CellStyle style = wb.createCellStyle();
        style.setBorderLeft(BorderStyle.THIN);
        Row rowLine = sheet.getRow(rowNum);
        Cell cell = rowLine.createCell(col);
        cell.setCellValue("罐表结束");
        cell.setCellStyle(style);
    }

    /**
     * 添加尾页说明信息
     *  @param rows
     * @param sheet
     * @param wb
     */
    private static void stepTableTailComment(int rows, Sheet sheet, Workbook wb) {
        CellStyle right = wb.createCellStyle();
        right.setAlignment(HorizontalAlignment.RIGHT);

        CellStyle left = wb.createCellStyle();
        left.setAlignment(HorizontalAlignment.LEFT);

        Row row = sheet.createRow(rows);
        Cell cell = row.createCell(1);
        cell.setCellValue(commentList.get(0).toString());
        cell.setCellStyle(right);

        sheet.addMergedRegion(new CellRangeAddress(rows, rows, 2, 6));
        cell = row.createCell(2);
        cell.setCellValue(commentList.get(1).toString());
        cell.setCellStyle(left);

        row = sheet.createRow(rows + 1);
        sheet.addMergedRegion(new CellRangeAddress(rows + 1, rows + 1, 2, 6));
        cell = row.createCell(2);
        cell.setCellValue(commentList.get(2).toString());
        cell.setCellStyle(left);
    }

    /**
     * 设置分区结尾信息
     *  @param rowNum
     * @param sheet
     * @param page
     * @param wb
     */
    private static void stepTailInfo(int rowNum, Sheet sheet, int page, Workbook wb) {

        CellStyle rightStyle = wb.createCellStyle();
        rightStyle.setAlignment(HorizontalAlignment.RIGHT);

        CellStyle leftStyle = wb.createCellStyle();
        leftStyle.setAlignment(HorizontalAlignment.LEFT);

//        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 1, 2));
        Row row = sheet.createRow(rowNum);
        Cell cell = row.createCell(1);
        cell.setCellValue("有效日期：");
        cell.setCellStyle(rightStyle);

        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 2, 5));
        cell = row.createCell(2);
        cell.setCellStyle(leftStyle);
        cell.setCellValue(expireDate);

        cell = row.createCell(6);
        cell.setCellValue(String.format("第%d页", page));

    }

    private static void renderPageTemplate(List sub, Sheet sheet, int page, Workbook wb) {
        System.out.println(page);
        //            1 设置表头
        int start = page * 56;
        renderTempalteHeader(start, sheet,wb);

        int startRow = start + 6;
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
//            第三步：每4条数据片生成一列数据渲染到列
            renderPageData(sub1, sheet, startRow, colNum,wb);

            subOffSet += 4;
            colNum += 2;
        }
    }


    private static void renderPageData(List sourceData, Sheet sheet, int rowNum, int col, Workbook wb) {

        CellStyle leftStyle = wb.createCellStyle();
        leftStyle.setBorderLeft(BorderStyle.THIN);
        leftStyle.setAlignment(HorizontalAlignment.CENTER);

        CellStyle rightStyle = wb.createCellStyle();
        rightStyle.setBorderRight(BorderStyle.THIN);
        rightStyle.setAlignment(HorizontalAlignment.CENTER);

        for (int i = 0; i < sourceData.size(); i++) {
            List data = (List) sourceData.get(i);
            for (int j = 0; j < data.size() / 2; j++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) {
                    row = sheet.createRow(rowNum);
                }
                Cell cell = row.createCell(col);

                //分片最后一条数据增加底部样式
                if (i == sourceData.size() - 1 && j == (data.size() / 2) - 1) {
                    CellStyle leftbStyle = wb.createCellStyle();
                    leftbStyle.setBorderLeft(BorderStyle.THIN);
                    leftbStyle.setBorderBottom(BorderStyle.THIN);
                    leftbStyle.setAlignment(HorizontalAlignment.CENTER);

                    cell.setCellValue(data.get(j * 2).toString());
                    cell.setCellStyle(leftbStyle);

                    CellStyle rightbStyle = wb.createCellStyle();
                    rightbStyle.setBorderRight(BorderStyle.THIN);
                    rightbStyle.setBorderBottom(BorderStyle.THIN);
                    rightbStyle.setAlignment(HorizontalAlignment.CENTER);

                    cell = row.createCell(col + 1);
                    cell.setCellValue(data.get(j * 2 + 1).toString());
                    cell.setCellStyle(rightbStyle);
                } else {

                    cell.setCellValue(data.get(j * 2).toString());
                    cell.setCellStyle(leftStyle);

                    cell = row.createCell(col + 1);
                    cell.setCellValue(data.get(j * 2 + 1).toString());
                    cell.setCellStyle(rightStyle);
                }
                rowNum++;
            }

            //分片的最后一条数据结尾不增加空行
            if (i < sourceData.size() - 1) {
                Row row = sheet.getRow(rowNum);
                if (row == null) {
                    row = sheet.createRow(rowNum);
                }
                Cell cell = row.createCell(col);
                cell.setCellValue("");
                cell.setCellStyle(leftStyle);

                cell = row.createCell(col + 1);
                cell.setCellValue("");
                cell.setCellStyle(rightStyle);
                rowNum++;
            }
        }
    }

    private static void renderTempalteHeader(int start, Sheet sheet, Workbook wb) {

        CellStyle rightStyle = wb.createCellStyle();
        rightStyle.setAlignment(HorizontalAlignment.RIGHT);

        CellStyle leftStyle = wb.createCellStyle();
        leftStyle.setAlignment(HorizontalAlignment.LEFT);


        HSSFFont font = (HSSFFont) wb.createFont();
        font.setBold(true);

        CellStyle centerStyle = wb.createCellStyle();
        centerStyle.setAlignment(HorizontalAlignment.CENTER);
        centerStyle.setFont(font);


        Row row = sheet.createRow(start);
        sheet.addMergedRegion(new CellRangeAddress(start, start, 1, 6));

        Cell cell = row.createCell(1);
        cell.setCellValue("容积表");
        cell.setCellStyle(centerStyle);

        start++;
        sheet.addMergedRegion(new CellRangeAddress(start, start, 1, 2));
        sheet.addMergedRegion(new CellRangeAddress(start, start, 3, 6));
        row = sheet.createRow(start);

        cell = row.createCell(1);
        cell.setCellValue("客户名称：");
        cell.setCellStyle(rightStyle);

        cell = row.createCell(3);
        cell.setCellValue(customerName);
        cell.setCellStyle(leftStyle);

        start++;
        sheet.addMergedRegion(new CellRangeAddress(start, start, 1, 2));
        sheet.addMergedRegion(new CellRangeAddress(start, start, 3, 6));
        row = sheet.createRow(start);

        cell = row.createCell(1);
        cell.setCellValue("罐号：");
        cell.setCellStyle(rightStyle);

        cell = row.createCell(3);
        cell.setCellValue(numbers);
        cell.setCellStyle(leftStyle);

        start++;
        sheet.addMergedRegion(new CellRangeAddress(start, start, 1, 2));
        sheet.addMergedRegion(new CellRangeAddress(start, start, 3, 6));
        row = sheet.createRow(start);
        cell = row.createCell(1);
        cell.setCellValue("证书编号：");
        cell.setCellStyle(rightStyle);


        CellStyle leftBottom = wb.createCellStyle();
        leftBottom.setBorderBottom(BorderStyle.THIN);
        leftBottom.setBorderLeft(BorderStyle.THIN);
        leftBottom.setAlignment(HorizontalAlignment.CENTER);

        CellStyle rightBottom = wb.createCellStyle();
        rightBottom.setBorderBottom(BorderStyle.THIN);
        rightBottom.setBorderRight(BorderStyle.THIN);
        rightBottom.setAlignment(HorizontalAlignment.CENTER);

        CellStyle leftTop = wb.createCellStyle();
        leftTop.setBorderTop(BorderStyle.THIN);
        leftTop.setBorderLeft(BorderStyle.THIN);
        leftTop.setAlignment(HorizontalAlignment.CENTER);

        CellStyle rightTop = wb.createCellStyle();
        rightTop.setBorderTop(BorderStyle.THIN);
        rightTop.setBorderRight(BorderStyle.THIN);
        rightTop.setAlignment(HorizontalAlignment.CENTER);

        CellStyle left = wb.createCellStyle();
        left.setBorderLeft(BorderStyle.THIN);

        CellStyle right = wb.createCellStyle();
        right.setBorderRight(BorderStyle.THIN);

        start++;
        row = sheet.createRow(start);
        cell = row.createCell(1);
        cell.setCellValue("实高");
        cell.setCellStyle(leftTop);

        cell = row.createCell(3);
        cell.setCellValue("实高");
        cell.setCellStyle(leftTop);

        cell = row.createCell(5);
        cell.setCellValue("实高");
        cell.setCellStyle(leftTop);


        cell = row.createCell(2);
        cell.setCellValue("容量");
        cell.setCellStyle(rightTop);

        cell = row.createCell(4);
        cell.setCellValue("容量");
        cell.setCellStyle(rightTop);

        cell = row.createCell(6);
        cell.setCellValue("容量");
        cell.setCellStyle(rightTop);


        start++;
        row = sheet.createRow(start);
        cell = row.createCell(1);
        cell.setCellValue("(cm)");
        cell.setCellStyle(leftBottom);

        cell = row.createCell(3);
        cell.setCellValue("(cm)");
        cell.setCellStyle(leftBottom);

        cell = row.createCell(5);
        cell.setCellValue("(cm)");
        cell.setCellStyle(leftBottom);


        cell = row.createCell(2);
        cell.setCellValue("(kL)");
        cell.setCellStyle(rightBottom);

        cell = row.createCell(4);
        cell.setCellValue("(kL)");
        cell.setCellStyle(rightBottom);

        cell = row.createCell(6);
        cell.setCellValue("(kL)");
        cell.setCellStyle(rightBottom);

        start++;
        for (int i = 0; i < 43; i++) {
            if (i == 42) {
                row = sheet.createRow(start + i);
                cell = row.createCell(1);
                cell.setCellStyle(leftBottom);
                cell = row.createCell(3);
                cell.setCellStyle(leftBottom);
                cell = row.createCell(5);
                cell.setCellStyle(leftBottom);

                cell = row.createCell(2);
                cell.setCellStyle(rightBottom);
                cell = row.createCell(4);
                cell.setCellStyle(rightBottom);
                cell = row.createCell(6);
                cell.setCellStyle(rightBottom);
            } else {
                row = sheet.createRow(start + i);
                cell = row.createCell(1);
                cell.setCellStyle(left);
                cell = row.createCell(3);
                cell.setCellStyle(left);
                cell = row.createCell(5);
                cell.setCellStyle(left);

                cell = row.createCell(2);
                cell.setCellStyle(right);
                cell = row.createCell(4);
                cell.setCellStyle(right);
                cell = row.createCell(6);
                cell.setCellStyle(right);
            }
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
            jxl.Workbook book = jxl.Workbook.getWorkbook(file);
//            获取页签
            jxl.Sheet sheet = book.getSheet(0);
//            获取总行数
            int rowTotal = sheet.getRows();
            jxl.Cell[] c1 = sheet.getRow(rowTotal - 11);
            for (int i = 0; i < c1.length; i++) {
                String val = c1[i].getContents();
                if (val != null && val.length() > 0) {
                    commentList.add(val);
                }
            }
            jxl.Cell[] c2 = sheet.getRow(rowTotal - 10);
            for (int i = 0; i < c2.length; i++) {
                String val = c2[i].getContents();
                if (val != null && val.length() > 0) {
                    commentList.add(val);
                }
            }
            System.out.println("说明："+commentList.toString());
            List result = new ArrayList();
            for (int i = 0; i < rowTotal; i++) {
                //                获取一行数据
                jxl.Cell[] cols = sheet.getRow(i);
                if (i <= 9 || i >= rowTotal - 14) {
                    for (int j = 0; j < cols.length; j++) {
                        jxl.Cell cell = cols[j];
                        jxl.CellType type = cell.getType();
                        if (type == jxl.CellType.LABEL) {
                            String val = cell.getContents();
                            if (val.startsWith("罐 号:")) {
                                numbers = cols[j + 1].getContents();
                                System.out.println("罐号："+numbers);
                                break;
                            }
                            if ("单 位:".equals(val)) {
                                customerName = cols[j + 1].getContents();
                                System.out.println("单位："+customerName);
                                break;
                            }
                            if (val.trim().startsWith("有效期") || val.trim().startsWith("有效日期")) {
                                expireDate = val.trim().substring(val.trim().indexOf(":")+1);
                                System.out.println("有效期："+val);
                                break;
                            }
                        }
                    }
                } else {
                    List row = new ArrayList();
                    for (int j = 0; j < cols.length; j++) {
                        jxl.Cell cell = cols[j];
                        jxl.CellType type = cell.getType();
                        if (type == jxl.CellType.NUMBER) {
                            row.add(j, Integer.parseInt(cell.getContents()));
                        }
                    }
                    result.add(row);
                }
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
}


