package cn.ylj;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;

/**
 * @author : yanglujian
 * create at:  2020/9/20  10:34 下午
 */
public class readDemo {

    public static void main(String[] args) throws IOException {
        //获取maven目录resource下的资源目录
        String path = readDemo.class.getClassLoader().getResource("read.xlsx").getPath();
        //1.获取工作簿对象
        XSSFWorkbook wb = new XSSFWorkbook(path);
        //2.获取工作表
        //XSSFSheet sheet1 = wb.getSheet("sheet1");//根据工作表名称获取
        XSSFSheet sheet = wb.getSheetAt(0);//根据工作表索引获取
        //3.获取行
        for (Row row : sheet) {
            //4.获取单元格
            for (Cell cell : row) {
                String value = cell.getStringCellValue();
                System.out.println(value);
            }
        }
        //5.释放资源
        wb.close();
    }
}