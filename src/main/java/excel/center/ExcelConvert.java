package excel.center;


import jxl.*;
import jxl.format.*;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.VerticalAlignment;
import jxl.read.biff.BiffException;
import jxl.write.*;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by liubo on 2017/8/11.
 */
public class ExcelConvert {

    private static List commentList = new ArrayList();
    private static String customerName = "";
    private static String numbers = "";
    private static String expireDate = "";
    private int last = 0;

    /**
     * 生成转置后数据excel
     *
     * @param sourceData
     */
    public static String writeConvertExcel(List sourceData, String fileName) {
        File writeFile = new File(fileName);
//        打开生成excel文件
        try {
            WritableWorkbook book = Workbook.createWorkbook(writeFile);
            //  生成名为“第一页”的工作表，参数0表示这是第一页
            WritableSheet sheet = book.createSheet("sheet1", 0);

            //            第二步：数据12组分组填充
            stepGroupOne(sourceData, sheet);

            book.write();
            book.close();
            return fileName;
        } catch (IOException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }
        return null;
    }


    /**
     * 设置分片表头信息
     *
     * @param sheet
     * @throws WriteException
     */
    private static void stepTableHeader(int rowNum, WritableSheet sheet) throws WriteException {
//        int rowNum = sheet.getRows();
        //            设置表头字体
        WritableFont headerFont = new WritableFont(WritableFont.TIMES, 16, WritableFont.BOLD);
        WritableCellFormat headerFormat = new WritableCellFormat(headerFont);
        headerFormat.setAlignment(Alignment.getAlignment(2));
//设置标题字体
        WritableFont bodyFont = new WritableFont(WritableFont.TIMES, 12, WritableFont.NO_BOLD);
        WritableCellFormat bodyFormat = new WritableCellFormat(bodyFont);
        bodyFormat.setAlignment(Alignment.getAlignment(2));


        WritableFont font = new WritableFont(WritableFont.TIMES, 12);
        WritableCellFormat bottom = new WritableCellFormat(font);
        bottom.setAlignment(Alignment.getAlignment(2));
        bottom.setBorder(Border.BOTTOM, BorderLineStyle.THIN);

        //            设置表头第一行内容
        sheet.mergeCells(1, rowNum, 6, rowNum);
        sheet.addCell(new Label(1, rowNum, "容积表", headerFormat));
//        sheet.addCell(new Label(1, rowNum, "容积表"));
        rowNum++;
//            设置第二行内容
        sheet.mergeCells(1, rowNum, 2, rowNum);
        sheet.addCell(new Label(1, rowNum, "客户名称:", bodyFormat));
//        sheet.addCell(new Label(1, rowNum, "客户名称:"));
        sheet.mergeCells(3, rowNum, 6, rowNum);
        sheet.addCell(new Label(3, rowNum, customerName, bodyFormat));
//        sheet.addCell(new Label(3, rowNum, customerName));
        rowNum++;
//            设置第三行内容
        sheet.mergeCells(1, rowNum, 2, rowNum);
        sheet.addCell(new Label(1, rowNum, "罐号:", bodyFormat));
//        sheet.addCell(new Label(1, rowNum, "罐号:"));
        sheet.mergeCells(3, rowNum, 6, rowNum);
        sheet.addCell(new Label(3, rowNum, "1"));
        rowNum++;
//        设置第四行内容
        sheet.mergeCells(1, rowNum, 2, rowNum);
        sheet.addCell(new Label(1, rowNum, "证书编号：", bottom));
//        sheet.addCell(new Label(1, rowNum, "证书编号："));
        sheet.mergeCells(3, rowNum, 6, rowNum);
        sheet.addCell(new Label(3, rowNum, "", bottom));
//        sheet.addCell(new Label(3, rowNum, ""));
        rowNum++;

        //            设置表头内容

        WritableCellFormat leftTop = new WritableCellFormat(bodyFont);
        leftTop.setAlignment(Alignment.getAlignment(2));
        leftTop.setBorder(Border.LEFT, BorderLineStyle.THIN);
        leftTop.setBorder(Border.RIGHT, BorderLineStyle.THIN);
        leftTop.setBorder(Border.TOP, BorderLineStyle.THIN);

        WritableCellFormat rightTop = new WritableCellFormat(bodyFont);
        rightTop.setAlignment(Alignment.getAlignment(2));
        rightTop.setBorder(Border.RIGHT, BorderLineStyle.THIN);
        rightTop.setBorder(Border.TOP, BorderLineStyle.THIN);

        WritableCellFormat leftBottom = new WritableCellFormat(bodyFont);
        leftBottom.setAlignment(Alignment.getAlignment(2));
        leftBottom.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
        leftBottom.setBorder(Border.RIGHT, BorderLineStyle.THIN);
        leftBottom.setBorder(Border.LEFT, BorderLineStyle.THIN);

        WritableCellFormat rightBottom = new WritableCellFormat(bodyFont);
        rightBottom.setAlignment(Alignment.getAlignment(2));
        rightBottom.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
        rightBottom.setBorder(Border.RIGHT, BorderLineStyle.THIN);
//
        int offset = 1;
//
        for (int i = 0; i < 3; i++) {
//                数据渲染
            sheet.addCell(new Label(i + offset, rowNum, "实高",leftTop));//leftTop
            sheet.addCell(new Label(i + 1 + offset, rowNum, "容量",rightTop));//, rightTop
            sheet.addCell(new Label(i + offset, rowNum + 1, "(cm)",leftBottom));//, leftBottom
            sheet.addCell(new Label(i + 1 + offset, rowNum + 1, "(kL)",rightBottom));//, rightBottom

            offset++;
        }
    }


    /**
     * 数据分片
     *
     * @param sourceData
     * @param sheet
     */
    private static void stepGroupOne(List sourceData, WritableSheet sheet) throws WriteException {
//            表体数据填充
        double step1 = 12;
        double stepTotal = Math.ceil(sourceData.size() / step1);
        int stepOffSet1 = 0;

        SheetSettings settings = sheet.getSettings();
//        sheet = copySheetSettingToSheet(sheet,settings);
        for (int i = 0; i < stepTotal; i++) {
            int from = stepOffSet1;
            int to = from + 12;
            if (to > sourceData.size()) {
                to = sourceData.size();
            }
//            数据12组分组
            List sub = sourceData.subList(from, to);

            System.out.println("rows " + sheet.getRows());
            int start = sheet.getRows();
            if (start > 0) {
                start += 6;
            }
//            1 设置表头
            stepTableHeader(start, sheet);
//            2 数据渲染
            stepGroupTwo(sub, sheet, sheet.getRows());

            if (i == stepTotal - 1) {
                setFinishedContent(sheet, sub.size());
            }
            stepOffSet1 += 12;
//            设置表尾
            stepTailInfo(sheet.getRows(), sheet, i + 1);

            if (i == stepTotal - 1) {
                stepTableTailComment(sheet.getRows(), sheet);
            }
//            设置打印格式
            int end = sheet.getRows();
            end += 6;
//            sheet.getSettings().setShowGridLines(false);
//            sheet.getSettings().setPrintHeaders(false);
            System.out.println("start row " + start + " end row " + end);
//            settings.setPrintArea(0, start, 8, end);
//            settings.setAutomaticFormulaCalculation(true);
//            settings.setDefaultRowHeight(15);
            settings.setFitHeight(sheet.getSettings().getFitHeight());
            settings.setFitWidth(sheet.getSettings().getFitWidth());
            settings.setFitToPages(true);
//            settings.setRecalculateFormulasBeforeSave(true);

            settings.setPrintTitlesCol(0, 8);
            settings.setPrintTitlesRow(start, end);
//            settings.setPrintArea(0, start, 7, end);
            settings.setPageOrder(PageOrder.RIGHT_THEN_DOWN);
           settings.setAutomaticFormulaCalculation(true);
        }
//        settings.setPrintArea(0,0,7,sheet.getRows());
    }

    private static void setFinishedContent(WritableSheet sheet, int size) throws WriteException {

        WritableCellFormat leftBottom = new WritableCellFormat();
        leftBottom.setBorder(Border.LEFT, BorderLineStyle.THIN);
        leftBottom.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
        leftBottom.setAlignment(Alignment.getAlignment(2));

        WritableCellFormat rightBottom = new WritableCellFormat();
        rightBottom.setBorder(Border.RIGHT, BorderLineStyle.THIN);
        rightBottom.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
        rightBottom.setAlignment(Alignment.getAlignment(2));
        int row = sheet.getRows();
        int col = 0;
        int rowNum = 0;
        if (size == 1 || size == 5 || size == 9) {
            rowNum = row - 3 * 10 - 2;
        } else if (size == 2 || size == 6 || size == 10) {
            rowNum = row - 2 * 10 - 1;
        } else if (size == 3 || size == 7 || size == 11) {
            rowNum = row - 10 - 1;
        } else if (size == 4 || size == 8) {
            rowNum = row - 4 * 10 - 3;
        } else {
            rowNum = row;
        }
        if (size <= 7) {
            col = 3;
        } else if (size >= 8 && size <= 11) {
            col = 5;
        } else {
            col = 1;
        }

        sheet.addCell(new Label(col, rowNum, "罐表结束"));
    }

    /**
     * 数据四组分片
     *
     * @param sub
     * @param sheet
     * @param startRow
     * @throws WriteException
     */
    private static void stepGroupTwo(List sub, WritableSheet sheet, int startRow) throws WriteException {
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
            recursiveFill(sub1, sheet, startRow, colNum);

            subOffSet += 4;
            colNum += 2;
        }

    }


    /**
     * 设置页尾分页信息
     *
     * @param row
     * @param sheet
     * @param page
     */
    private static void stepTailInfo(int row, WritableSheet sheet, int page) throws WriteException {
        System.out.println("tail info " + row);
        WritableFont bodyFont = new WritableFont(WritableFont.TIMES, 12, WritableFont.NO_BOLD);
        WritableCellFormat bodyFormat = new WritableCellFormat(bodyFont);
        bodyFormat.setAlignment(Alignment.getAlignment(2));
        bodyFormat.setVerticalAlignment(VerticalAlignment.CENTRE);


        sheet.mergeCells(1, row, 2, row);
        sheet.addCell(new Label(1, row, "有效日期:", bodyFormat));
        sheet.mergeCells(3, row, 5, row);
        sheet.addCell(new Label(3, row, "2017年7月27日至2021年7月26日", bodyFormat));
        sheet.addCell(new Label(6, row, String.format("第%d页", page), bodyFormat));
    }


    /**
     * 行列转置数据填充
     *
     * @param sourceData
     * @param sheet
     * @param rowNum
     * @param col
     * @throws WriteException
     */
    private static void recursiveFill(List sourceData, WritableSheet sheet, int rowNum, int col) throws WriteException {
        WritableCellFormat right = new WritableCellFormat();
        right.setAlignment(Alignment.getAlignment(2));
        right.setBorder(Border.RIGHT, BorderLineStyle.THIN);

        WritableCellFormat left = new WritableCellFormat();
        left.setBorder(Border.LEFT, BorderLineStyle.THIN);
        left.setAlignment(Alignment.getAlignment(2));

        WritableCellFormat leftBottom = new WritableCellFormat();
        leftBottom.setBorder(Border.LEFT, BorderLineStyle.THIN);
        leftBottom.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
        leftBottom.setAlignment(Alignment.getAlignment(2));

        WritableCellFormat rightBottom = new WritableCellFormat();
        rightBottom.setBorder(Border.RIGHT, BorderLineStyle.THIN);
        rightBottom.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
        rightBottom.setAlignment(Alignment.getAlignment(2));

        int width = 10;
        for (int i = 0; i < sourceData.size(); i++) {
            List row1 = (List) sourceData.get(i);
//                每条行数据转置为列数据
            for (int j = 0; j < row1.size(); j++) {
                if (j % 2 == 0) {
                    sheet.setColumnView(col, width);
                    if (i == sourceData.size() - 1 && j == row1.size() - 2) {
                        sheet.addCell(new Label(col, rowNum, row1.get(j).toString(),leftBottom));//leftBottom
                    } else {
                        sheet.addCell(new Label(col, rowNum, row1.get(j).toString(),left ));//left
                    }
                } else {
                    sheet.setColumnView(col + 1, width);
                    if (i == sourceData.size() - 1 && j == row1.size() - 1) {
                        sheet.addCell(new Label(col + 1, rowNum, row1.get(j).toString() ,rightBottom));//rightBottom
                    } else {
                        sheet.addCell(new Label(col + 1, rowNum, row1.get(j).toString(),right));//right
                    }
                    rowNum++;
                }
            }
            rowNum++;
            if (i < sourceData.size() - 1) {
                sheet.addCell(new Label(col, rowNum - 1, "",left));//, left
                sheet.addCell(new Label(col + 1, rowNum - 1, "",right));//, right
            }
        }
    }


    /**
     * 添加尾页说明信息
     *
     * @param rows
     * @param sheet
     * @throws WriteException
     */
    private static void stepTableTailComment(int rows, WritableSheet sheet) throws WriteException {

        WritableFont font = new WritableFont(WritableFont.TIMES, 11, WritableFont.NO_BOLD);
        WritableCellFormat bodyFormat = new WritableCellFormat(font);
        bodyFormat.setAlignment(Alignment.getAlignment(2));
        bodyFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
        bodyFormat.setWrap(true);

        sheet.addCell(new Label(1, rows, commentList.get(0).toString(), bodyFormat));
        sheet.mergeCells(2, rows, 6, rows);
        sheet.addCell(new Label(2, rows, commentList.get(1).toString(), bodyFormat));
        sheet.mergeCells(2, rows + 1, 6, rows + 1);
        sheet.addCell(new Label(2, rows + 1, commentList.get(2).toString(), bodyFormat));
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
            Cell[] c1 = sheet.getRow(rowTotal - 11);
            for (int i = 0; i < c1.length; i++) {
                String val = c1[i].getContents();
                if (val != null && val.length() > 0) {
                    commentList.add(val);
                }
            }
            Cell[] c2 = sheet.getRow(rowTotal - 10);
            for (int i = 0; i < c2.length; i++) {
                String val = c2[i].getContents();
                if (val != null && val.length() > 0) {
                    commentList.add(val);
                }
            }
            Cell[] c3 = sheet.getRow(4);
            for (int i = 0; i < c3.length; i++) {
                String val = c3[i].getContents();
                if ("单 位:".equals(val)) {
                    customerName = c3[i + 1].getContents();
                    System.out.println(customerName);
                    break;
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


    public static WritableSheet copySheetSettingToSheet(WritableSheet sheet, SheetSettings setting) {
        //      设置原Sheet打印属性到新Sheet页
        SheetSettings sheetSettings = sheet.getSettings();

        sheetSettings.setAutomaticFormulaCalculation(setting.getAutomaticFormulaCalculation());
        sheetSettings.setBottomMargin(setting.getBottomMargin());
        sheetSettings.setCopies(setting.getCopies());
        sheetSettings.setDefaultColumnWidth(setting.getDefaultColumnWidth());
        sheetSettings.setDefaultRowHeight(setting.getDefaultRowHeight());
        sheetSettings.setDisplayZeroValues(setting.getDisplayZeroValues());
        sheetSettings.setFitHeight(setting.getFitHeight());
        sheetSettings.setFitToPages(setting.getFitToPages());
        sheetSettings.setFitWidth(setting.getFitWidth());

        HeaderFooter footer = setting.getFooter();
        if (footer != null) {
            sheetSettings.setFooter(footer);
        }
        sheetSettings.setFooterMargin(setting.getFooterMargin());
        HeaderFooter header = setting.getHeader();
        if (header != null) {
            sheetSettings.setHeader(header);
        }
        sheetSettings.setHeaderMargin(setting.getHeaderMargin());
        sheetSettings.setHidden(setting.isHidden());
        sheetSettings.setHorizontalCentre(setting.isHorizontalCentre());
        sheetSettings.setHorizontalFreeze(setting.getHorizontalFreeze());
        sheetSettings.setHorizontalPrintResolution(setting.getHorizontalPrintResolution());
        sheetSettings.setLeftMargin(setting.getLeftMargin());
        sheetSettings.setNormalMagnification(setting.getNormalMagnification());
        PageOrientation pageOrientation = setting.getOrientation();
        if (pageOrientation != null) {
            sheetSettings.setOrientation(pageOrientation);
        }
        sheetSettings.setPageBreakPreviewMagnification(setting.getPageBreakPreviewMagnification());
        sheetSettings.setPageBreakPreviewMode(setting.getPageBreakPreviewMode());
        sheetSettings.setPageStart(setting.getPageStart());
        PaperSize paperSize = setting.getPaperSize();
        if (paperSize != null) {
            sheetSettings.setPaperSize(setting.getPaperSize());
        }

        sheetSettings.setPassword(setting.getPassword());
        sheetSettings.setPasswordHash(setting.getPasswordHash());
        Range printArea = setting.getPrintArea();
        if (printArea != null) {
            sheetSettings.setPrintArea(printArea.getTopLeft() == null ? 0 : printArea.getTopLeft().getColumn(),
                    printArea.getTopLeft() == null ? 0 : printArea.getTopLeft().getRow(),
                    printArea.getBottomRight() == null ? 0 : printArea.getBottomRight().getColumn(),
                    printArea.getBottomRight() == null ? 0 : printArea.getBottomRight().getRow());
        }

        sheetSettings.setPrintGridLines(setting.getPrintGridLines());
        sheetSettings.setPrintHeaders(setting.getPrintHeaders());

        Range printTitlesCol = setting.getPrintTitlesCol();
        if (printTitlesCol != null) {
            sheetSettings.setPrintTitlesCol(printTitlesCol.getTopLeft() == null ? 0 : printTitlesCol.getTopLeft().getColumn(),
                    printTitlesCol.getBottomRight() == null ? 0 : printTitlesCol.getBottomRight().getColumn());
        }
        Range printTitlesRow = setting.getPrintTitlesRow();
        if (printTitlesRow != null) {
            sheetSettings.setPrintTitlesRow(printTitlesRow.getTopLeft() == null ? 0 : printTitlesRow.getTopLeft().getRow(),
                    printTitlesRow.getBottomRight() == null ? 0 : printTitlesRow.getBottomRight().getRow());
        }

        sheetSettings.setProtected(setting.isProtected());
        sheetSettings.setRecalculateFormulasBeforeSave(setting.getRecalculateFormulasBeforeSave());
        sheetSettings.setRightMargin(setting.getRightMargin());
        sheetSettings.setScaleFactor(setting.getScaleFactor());
        sheetSettings.setSelected(setting.isSelected());
        sheetSettings.setShowGridLines(setting.getShowGridLines());
        sheetSettings.setTopMargin(setting.getTopMargin());
        sheetSettings.setVerticalCentre(setting.isVerticalCentre());
        sheetSettings.setVerticalFreeze(setting.getVerticalFreeze());
        sheetSettings.setVerticalPrintResolution(setting.getVerticalPrintResolution());
        sheetSettings.setZoomFactor(setting.getZoomFactor());
        return sheet;
    }

}
