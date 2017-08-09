package excel.utils;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by liubo on 2017/8/8.
 */
public class Utils {

    private static final Logger  log = LoggerFactory.getLogger(Utils.class);

    public static boolean checkFile(String path){
        String str = path.substring(0,path.lastIndexOf(Constants.SYSTEM_SEPARATOR));
        File dir = new File(str);
        if (!dir.exists()) {
            log.info("文件目录{}正在创建",path);
            return dir.mkdirs();
        }
        return false;
    }

    /**
     * 根据format 格式化日期
     *
     * @param format
     * @return
     */
    public static String getDate(final DateFormatter format) {
        SimpleDateFormat sdf = new SimpleDateFormat(format.formatter);
        String dateStr = sdf.format(new Date());
        return dateStr;
    }

    public static String getDate() {
        return getDate(DateFormatter.DEFAULT_FORMATTER);
    }

    public static enum DateFormatter {

        DEFAULT_FORMATTER(0, "yyyyMMddHHmmss", "默认格式化日期到时分秒"),
        DAY_FORMATTER(1, "yyyyMMdd", "格式日期到天"),
        TIME_FORMATTER(2, "HHmmss", "格式化为时间");


        private int index;
        private String formatter;
        private String desc;

        DateFormatter(int index, String formatter, String desc) {
            this.index = index;
            this.formatter = formatter;
            this.desc = desc;
        }

        public int getIndex() {
            return index;
        }

        public void setIndex(int index) {
            this.index = index;
        }

        public String getFormatter() {
            return formatter;
        }

        public void setFormatter(String formatter) {
            this.formatter = formatter;
        }

        public String getDesc() {
            return desc;
        }

        public void setDesc(String desc) {
            this.desc = desc;
        }
    }
}
