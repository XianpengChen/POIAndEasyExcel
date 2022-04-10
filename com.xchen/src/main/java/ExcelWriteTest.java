import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriteTest {
    String PATH = "C:\\Users\\x'chen\\Desktop\\excel读写练习\\";
    @Test
    public void testWrite03() throws IOException {
//        1.创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
//        2.创建一个工作表
        Sheet sheet = workbook.createSheet("狂神观众统计表");
//        3.创建第一行
        Row row1 = sheet.createRow(0);
//        4.创建一个单元格,(1,1)
        Cell cell = row1.createCell(0);
        cell.setCellValue("今日新增观众");
//        创建单元格（1，2）
        Cell cell2 = row1.createCell(1);
        cell2.setCellValue(666);

//        创建第二行
        Row row2 = sheet.createRow(1);
        Cell row2Cell = row2.createCell(0);
        row2Cell.setCellValue("统计时间");
        Cell row2Cell1 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        row2Cell1.setCellValue(time);

//        生成一张表（IO流）03版本就是使用xls结尾！
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "狂神观众统计表03.xls");
//        工作簿写到流里去
        workbook.write(fileOutputStream);
//        关闭流
        fileOutputStream.close();
        System.out.println("excel表生成完毕");

    }
    @Test
    public void testWrite07() throws IOException {
//        1.创建一个工作簿
        Workbook workbook = new XSSFWorkbook();
//        2.创建一个工作表
        Sheet sheet = workbook.createSheet("狂神观众统计表");
//        3.创建第一行
        Row row1 = sheet.createRow(0);
//        4.创建一个单元格,(1,1)
        Cell cell = row1.createCell(0);
        cell.setCellValue("今日新增观众");
//        创建单元格（1，2）
        Cell cell2 = row1.createCell(1);
        cell2.setCellValue(666);

//        创建第二行
        Row row2 = sheet.createRow(1);
        Cell row2Cell = row2.createCell(0);
        row2Cell.setCellValue("统计时间");
        Cell row2Cell1 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        row2Cell1.setCellValue(time);

//        生成一张表（IO流）07版本就是使用xlsx结尾！
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "狂神观众统计表07.xlsx");
//        工作簿写到流里去
        workbook.write(fileOutputStream);
//        关闭流
        fileOutputStream.close();
        System.out.println("excel表生成完毕");

    }
}
