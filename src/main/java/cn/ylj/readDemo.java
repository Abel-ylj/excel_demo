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

/**
 * @author : yanglujian
 * create at:  2020/9/20  10:34 下午
 */
public class readDemo {

    public static void main(String[] args) throws IOException {
        //获取maven目录resource下的资源目录
        String path = readDemo.class.getClassLoader().getResource("test.xlsx").getPath();
        //1.获取工作簿对象
        XSSFWorkbook wb = new XSSFWorkbook(path);
        //2.获取工作表
        //XSSFSheet sheet1 = wb.getSheet("sheet1");//根据工作表名称获取
        XSSFSheet sheet = wb.getSheetAt(0);//根据工作表索引获取

        //3.获取行， (获取的是最后一行的索引从0开始)
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            XSSFRow row = sheet.getRow(i);
            //4.获取列 (获取的列的数量从1开始)
            short lastCellNum = row.getLastCellNum();
            for (int j = 0; j < lastCellNum; j++) {
                XSSFCell cell = row.getCell(j);
                System.out.println(cell.getStringCellValue());
            }
            System.out.println();
        }
        //5.释放资源
        wb.close();
    }
}