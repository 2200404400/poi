import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class Main {
    public static final String PATH = "D:\\Office\\Project\\poi\\poi\\src\\main\\file\\";
    // 03版本的Excel后缀
    public static final String EXCEL03 = ".xls";
    // 07版本的Excel后缀
    public static final String EXCEL07 = ".xlsx";

    public static void main(String[] args) throws Exception {
        excel03();
        excel07();
    }

    /**
     * 03版excle
     */
    public static void excel03() throws Exception{
        // 创建工作簿
        Workbook workbook = new HSSFWorkbook();
        // 创建工作表
        Sheet sheet = workbook.createSheet();
        // 创建行
        Row row1 = sheet.createRow(0);
        // 创建单元格（1，1）
        Cell cell11 = row1.createCell(0);
        // 给单元格填数据
        cell11.setCellValue("这是第一行第一个单元格");
        // 创建单元格（1，2）
        Cell cell12 = row1.createCell(1);

        cell12.setCellValue("这里是第一行第二个单元格");

        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        Cell cell22 = row2.createCell(1);
        cell21.setCellValue("21");
        cell22.setCellValue("22");


        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "ApachePoi03测试" + EXCEL03);
        workbook.write(fileOutputStream);
        fileOutputStream.close();

        System.out.println("文件输出完毕");
    }

    /**
     * 07版excel
     */
    public static void excel07() throws Exception{
        // 创建工作簿
        Workbook workbook = new XSSFWorkbook();
        // 创建工作表
        Sheet sheet = workbook.createSheet();
        // 创建行
        Row row1 = sheet.createRow(0);
        // 创建单元格（1，1）
        Cell cell11 = row1.createCell(0);
        // 给单元格填数据
        cell11.setCellValue("这是第一行第一个单元格");
        // 创建单元格（1，2）
        Cell cell12 = row1.createCell(1);

        cell12.setCellValue("这里是第一行第二个单元格");

        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        Cell cell22 = row2.createCell(1);
        cell21.setCellValue("21");
        cell22.setCellValue("22");


        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "ApachePoi03测试" + EXCEL07);
        workbook.write(fileOutputStream);
        fileOutputStream.close();

        System.out.println("文件输出完毕");
    }
}
