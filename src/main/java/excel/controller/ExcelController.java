package excel.controller;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.Future;


import excel.center.ExcelConvert;
import excel.center.ExcelHide;
import excel.center.PoiExcelConvert;
import excel.center.PoiExcelHide;
import excel.current.ReadHideHandler;
import excel.current.TaskService;
import excel.current.UploadHandler;
import excel.current.WriteHideHandler;
import excel.utils.Constants;
import excel.utils.Utils;
import excel.utils.ZipUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

@Controller
@RequestMapping("excel")
public class ExcelController {

    private static final Logger log = LoggerFactory.getLogger(ExcelController.class);

    private final String dest = "temp/upload";

    private final String export = "temp/export";

    private String downloadFile = null;

    private List<String> mfilePath;

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

    //    http://blog.csdn.net/coding13/article/details/54577076
    @RequestMapping(value = "/baupload", method = RequestMethod.POST)
    public String handleBatchFileUpload(Model model, @RequestParam("type") String type, HttpServletRequest request) {
        List<MultipartFile> files = ((MultipartHttpServletRequest) request).getFiles("file");
        final TaskService[] service = new TaskService[1];
        files.stream().forEach(multipartFile -> {
            UploadHandler upload = new UploadHandler(multipartFile);
            service[0] = new TaskService();
            try {
                Future<String> future = service[0].doTask(upload);
                ReadHideHandler read = new ReadHideHandler(future.get());
                List data = (List) service[0].doTask(read);

                String dateStr = Utils.getDate(Utils.DateFormatter.DAY_FORMATTER);
                String timeStr = Utils.getDate(Utils.DateFormatter.TIME_FORMATTER);
                String exportFile = export + Constants.SYSTEM_SEPARATOR + dateStr + Constants.SYSTEM_SEPARATOR + timeStr + "_export.xls";
                Utils.checkFile(exportFile);
                WriteHideHandler write = new WriteHideHandler(exportFile, data);
                service[0].doTask(write);
            } catch (ExecutionException e) {
                e.printStackTrace();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        });
        return "";
    }

    @RequestMapping(value = "/upload", method = RequestMethod.POST)
    public String handleFileUpload(Model model, @RequestParam("type") String type, @RequestParam("file") MultipartFile file, HttpServletRequest request) {

        List<MultipartFile> files = ((MultipartHttpServletRequest) request).getFiles("file");
        mfilePath = new ArrayList<>();
        String dateStr = Utils.getDate(Utils.DateFormatter.DAY_FORMATTER);
        String timeStr = Utils.getDate(Utils.DateFormatter.TIME_FORMATTER);

        String commonPath = Constants.SYSTEM_SEPARATOR + dateStr + Constants.SYSTEM_SEPARATOR + timeStr + Constants.SYSTEM_SEPARATOR;

        files.stream().forEach(multipartFile -> {
            loopUpload(type, commonPath, multipartFile);
        });
        String path = "export.zip";
        try {
            ZipUtils.doCompress(export + commonPath, path);
        } catch (IOException e) {
            e.printStackTrace();
        }

        model.addAttribute("path", path);//downloadFile
        model.addAttribute("type", type);
        return "/excel/download";
    }

    private void loopUpload(String type, String commonPath, MultipartFile file) {
        String name = file.getOriginalFilename();
        log.info("上传文件名称：{},文件大小为：{} kb", name, file.getSize() / 1000);
        if (!file.isEmpty()) {
            try {
                String fileName = dest + commonPath + name;//上传目录

                Utils.checkFile(fileName);
                File uploadFile = new File(fileName);
                byte[] bytes = file.getBytes();
                BufferedOutputStream stream = new BufferedOutputStream(new FileOutputStream(uploadFile));
                stream.write(bytes);
                stream.close();
                log.info("上传文件成功");

                String exportFile = export + commonPath + name;//下载目录
                Utils.checkFile(exportFile);
                log.info("开始解析excel文件");

                if ("hide".equals(type)) {
                    List dataList = ExcelHide.readHideExcel(fileName);
                    PoiExcelHide.writeHideExcel(dataList, exportFile);
                } else {
                    List data = PoiExcelConvert.readSourceExcel(fileName);
                    PoiExcelConvert.writeConvertExcel(data, exportFile);
                }

                log.info("excel隐藏内容完成");


            } catch (Exception e) {
                e.getStackTrace();

            }
        }
    }


    //文件下载相关代码
    @RequestMapping("/download")
    public String downloadFile(HttpServletRequest request, @RequestParam("filePath") String filePath, @RequestParam("type") String type, HttpServletResponse response) {
        FileInputStream fis = null;
        BufferedInputStream bis = null;
        OutputStream output = null;
        try {
            String name = new String(filePath.getBytes("iso-8859-1"), "utf-8");
            File file = new File(name);

            if (file.exists()) {
                String fileName = new String(file.getName().getBytes("utf-8"), "iso-8859-1");
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