package cn.ylj;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

/**
 * @author : yanglujian
 * create at:  2020/9/20  10:34 下午
 */
public class readDemo {

    public static void main(String[] args) throws IOException {
        String path = readDemo.class.getClassLoader().getResource("").getPath();
        System.out.println(path);
    }
}