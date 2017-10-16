package excel.center;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class PoiExcelHide {
//    public static void main(String[] args) {
//        List data = ExcelHide.readHideExcel("/Users/liubo/workspace/javapros/excel-center/temp/upload/20171014/212633/中国石油四川成都龙泉驿区北干道站1号罐毫米.xls");
//        try {
//            writeHideExcel(data, "/Users/liubo/Downloads/t.xls");
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }

    public static void writeHideExcel(List sourceData, String fileName) throws IOException {
        Workbook wb = new HSSFWorkbook();

        Sheet sheet = wb.createSheet("sheet1");
        sheet.setDefaultColumnWidth(6);

        CellStyle style = wb.createCellStyle();
        HSSFFont font = (HSSFFont) wb.createFont();
        font.setBold(true);
        font.setFontName("等线");
        font.setFontHeightInPoints((short) 10);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setFont(font);





        out:
        for (int i = 0; i < sourceData.size(); i++) {
            List row = (List) sourceData.get(i);
            Row sheetRow = sheet.createRow(i);
            inner:
            for (int j = 0; j < row.size(); j++) {
                Cell cell = sheetRow.createCell(j);
                Object obj = row.get(j);
                if (obj instanceof String) {
                    if (((String) obj).startsWith("容积表")) {
                        sheet.addMergedRegion(new CellRangeAddress(i, i, 0, 10));
                        cell.setCellValue(obj.toString());

                        CellStyle style1 = wb.createCellStyle();
                        HSSFFont font1 = (HSSFFont) wb.createFont();
                        font1.setBold(true);
                        font1.setFontName("黑体");
                        font1.setFontHeightInPoints((short) 20);
                        style1.setAlignment(HorizontalAlignment.CENTER);
                        style1.setFont(font1);
                        cell.setCellStyle(style1);
                        continue out;
                    }
                    if (((String) obj).startsWith("客户名称")) {
                        sheet.addMergedRegion(new CellRangeAddress(i, i, 0, 1));
                        sheet.addMergedRegion(new CellRangeAddress(i, i, 2, 10));
                        cell.setCellValue(obj.toString());
                        cell.setCellStyle(style);

                        Cell cell1 = sheetRow.createCell(j + 2);
                        cell1.setCellValue(row.get(1).toString());
                        cell1.setCellStyle(style);
                        continue out;
                    }
//                    if (((String) obj).startsWith("客户名称")){
//                        sheet.addMergedRegion(new CellRangeAddress(i,i,1,2));
//                        cell.setCellValue(obj.toString());
//                        cell.setCellStyle(style);
//                    }
                    if (((String) obj).startsWith("证书编号")){
                        sheet.addMergedRegion(new CellRangeAddress(i,i,0,1));
                        sheet.addMergedRegion(new CellRangeAddress(i, i, 2, 10));
                        cell.setCellValue(obj.toString());
                        cell.setCellStyle(style);
                        continue out;
                    }
                    if (((String) obj).startsWith("单位")) {
                        sheet.addMergedRegion(new CellRangeAddress(i, i, 10, 11));
                        cell.setCellValue(obj.toString());

                        cell.setCellStyle(style);

                        continue out;
                    }
                    if (((String) obj).startsWith("罐")){
                        sheet.addMergedRegion(new CellRangeAddress(i,i,0,1));
                        sheet.addMergedRegion(new CellRangeAddress(i, i, 2, 10));
                        cell.setCellValue("罐       号:");
                        cell.setCellStyle(style);
                        Cell cell1 = sheetRow.createCell(j+2);
                        cell1.setCellValue(row.get(1).toString());
                        cell1.setCellStyle(style);
                        continue out;
                    }
                    if (((String) obj).startsWith("---")) {
                        sheet.addMergedRegion(new CellRangeAddress(i, i, 0, 10));
                        cell.setCellValue("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                        cell.setCellStyle(style);
                        continue out;
                    }
                    if (((String) obj).trim().startsWith("有效期")) {
                        sheet.addMergedRegion(new CellRangeAddress(i, i, j - 1, j + 3));
                        Cell cell1 = sheetRow.createCell(j - 1);
                        cell1.setCellValue(obj.toString());
                        cell1.setCellStyle(style);
                        continue out;
                    }
                    cell.setCellValue(obj.toString());
                    cell.setCellStyle(style);
                    continue;
                }
                if (obj instanceof Integer) {
                    cell.setCellValue(((Integer) obj).longValue());
                    cell.setCellStyle(style);
                    continue;
                }
            }
        }
        //  写入数据并关闭文件
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            wb.write(fileOut);
            fileOut.close();
            wb.close();
        }
    }
}
