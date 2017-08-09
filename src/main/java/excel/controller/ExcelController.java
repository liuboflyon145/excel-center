package excel.controller;

import java.io.*;
import java.util.List;


import excel.center.ExcelTools;
import excel.utils.Constants;
import excel.utils.Utils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

@Controller
@RequestMapping("excel")
public class ExcelController {

    private static final Logger log = LoggerFactory.getLogger(ExcelController.class);

    private final String dest = "temp/upload";

    private final String export = "temp/export";

    private String downloadFile = null;

    private String type = "";

    @RequestMapping(value = "/hide", method = RequestMethod.GET)
    public String excelHide(Model model) {
        model.addAttribute("downloadFile", downloadFile);
        return "excel/hide";
    }

    @RequestMapping(value = "/convert", method = RequestMethod.GET)
    public String excelConvert(Model model) {
        model.addAttribute("downloadFile", downloadFile);
        return "excel/convert";
    }

    @RequestMapping(value = "/upload", method = RequestMethod.POST)
    public String handleFileUpload(Model model,@RequestParam("type") String type, @RequestParam("file") MultipartFile file) {
        String name = file.getName();

        log.info("上传文件名称：{},文件大小为：{} kb", name, file.getSize() / 1000);
        if (!file.isEmpty()) {
            try {
                String dateStr = Utils.getDate(Utils.DateFormatter.DAY_FORMATTER);
                String timeStr = Utils.getDate(Utils.DateFormatter.TIME_FORMATTER);

                String fileName = dest + Constants.SYSTEM_SEPARATOR + dateStr + Constants.SYSTEM_SEPARATOR + timeStr + "-uploaded.xls";
                Utils.checkFile(fileName);
                File uploadFile = new File(fileName);
                byte[] bytes = file.getBytes();
                BufferedOutputStream stream = new BufferedOutputStream(new FileOutputStream(uploadFile));
                stream.write(bytes);
                stream.close();
                log.info("上传文件成功");

                String exportFile = export + Constants.SYSTEM_SEPARATOR + dateStr + Constants.SYSTEM_SEPARATOR + timeStr + "_export.xls";
                Utils.checkFile(exportFile);
                log.info("开始解析excel文件");
                if ("hide".equals(type)) {
                    List dataList = ExcelTools.readHideExcel(fileName);
                    ExcelTools.writeHideExcel(dataList, exportFile);
                } else {
                    List data = ExcelTools.readSourceExcel(fileName);
                    ExcelTools.writeConvertExcel(data, exportFile);
                }
                downloadFile = exportFile;
                log.info("excel隐藏内容完成");
                model.addAttribute("path",exportFile);
                model.addAttribute("type",type);
                return "/excel/download";

            } catch (Exception e) {
                e.getStackTrace();
                return "error";
            }
        } else {
            return "error";
        }
    }


    //文件下载相关代码
    @RequestMapping("/download")
    public String downloadFile(@RequestParam("filePath") String filePath,@RequestParam("type") String type,HttpServletResponse response) {
        FileInputStream fis = null;
        BufferedInputStream bis = null;
        OutputStream output = null;
        try {
            File file = new File(filePath);
            if (file.exists()) {
                String fileName = new String("excel处理结果.xls".getBytes("utf-8"), "iso-8859-1");
                response.setCharacterEncoding("UTF-8");
                response.setContentType("application/force-download");// 设置强制下载不打开
                response.addHeader("Content-Disposition", "attachment;fileName=" + fileName);// 设置文件名
                byte[] buffer = new byte[1024];


                fis = new FileInputStream(file);
                bis = new BufferedInputStream(fis);
                output = response.getOutputStream();
                int i = bis.read(buffer);
                while (i != -1) {
                    output.write(buffer, 0, i);
                    i = bis.read(buffer);
                }
                output.flush();
                log.info("文件下载成功");
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (bis != null) {
                try {
                    bis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

        }
        return "/";
    }

}