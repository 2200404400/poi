import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class ExcelWrite {
    public static final String PATH = "D:\\Office\\Project\\poi\\poi\\src\\main\\file\\";
    // 03版本的Excel后缀
    public static final String EXCEL03 = ".xls";
    // 07版本的Excel后缀
    public static final String EXCEL07 = ".xlsx";

    public static void main(String[] args) throws Exception {
        excel03();
        excel07();
        excelProBig07();
    }

    /**
     * 03版excle，写的比07版快，但是列最高未65535
     * 它首先将所有的数据加载到内存中，再一次性写入，所以很快
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
     * 07版excel，写的速度比03慢，但是列的数量没有限制
     * 它每一条数据都进行一次io操作所以很耗时间
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

    /**
     * 07版excle的升级版，写入速度更快
     * 测试写入大数据的速度
     * 它每次写入100条数据，多余的数据将放入临时文件，当写入完毕后需要删除临时文件
     */
    public static void excelProBig07() throws Exception{
        // 创建工作簿
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        // 创建工作表
        Sheet sheet = workbook.createSheet();

        for (int i = 0; i < 100000; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 9; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "ApachePoiProBig03测试" + EXCEL07);
        workbook.write(fileOutputStream);
        // 删除临时文件
        workbook.dispose();
        fileOutputStream.close();

        System.out.println("文件输出完毕");
    }
}
