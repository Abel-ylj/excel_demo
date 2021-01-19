import cn.ylj.readDemo;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author : yanglujian
 * create at:  2021/1/19  6:22 下午
 */
public class ExcelTest {

    @Test
    public void writeTest() throws IOException {
        //1. 在内存中创建xlsx文件
        XSSFWorkbook wb = new XSSFWorkbook();
        //2. 创建sheet对象
        XSSFSheet sheet = wb.createSheet();
        //3. 创建row对象
        XSSFRow titleRow = sheet.createRow(0);
        //4. 创建cell对象
        titleRow.createCell(0).setCellValue("姓名");
        titleRow.createCell(1).setCellValue("年龄");
        //5. 创建输出流
        String root = readDemo.class.getClassLoader().getResource("").getPath();
        File file = new File(root, "out.xlsx");
        wb.write(new FileOutputStream(file));
    }

    @Test
    public void readTest() throws IOException {
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