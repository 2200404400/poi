import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ExcelRead {
    public static final String PATH = "D:\\Office\\Project\\poi\\poi\\src\\main\\file\\";
    // 03版本的Excel后缀
    public static final String EXCEL03 = ".xls";
    // 07版本的Excel后缀
    public static final String EXCEL07 = ".xlsx";

    public static void main(String[] args) throws Exception {
        //excel03();
        //excel07();
        excelDataType03();
    }

    /**
     * 03版excle，读的方法和写很像，只是将set变成get
     */
    public static void excel03() throws Exception{
        // 获取到文件
        FileInputStream input = new FileInputStream(PATH + "ApachePoi03测试" + EXCEL03);
        // 将文件属性获取到HSSF里
        HSSFWorkbook workbook = new HSSFWorkbook(input);
        // 获取工作表
        HSSFSheet sheetAt = workbook.getSheetAt(0);
        // 获取列
        HSSFRow row = sheetAt.getRow(0);
        // 获取单元格
        HSSFCell cell = row.getCell(0);
        // 获取数据
        String stringCellValue = cell.getStringCellValue();
        // 输出数据
        System.out.println(stringCellValue);
        input.close();
    }

    /**
     * 07版excel
     */
    public static void excel07() throws Exception{
        FileInputStream input = new FileInputStream(PATH + "ApachePoi03测试" + EXCEL07);
        XSSFWorkbook sheets = new XSSFWorkbook(input);
        XSSFSheet sheetAt = sheets.getSheetAt(0);
        XSSFRow row = sheetAt.getRow(0);
        XSSFCell cell = row.getCell(0);
        String stringCellValue = cell.getStringCellValue();
        System.out.println(stringCellValue);
        input.close();
    }

    /**
     * 03版判断数据类型
     */
    public static void excelDataType03() throws Exception {
        FileInputStream input = new FileInputStream(PATH + "ApachePoi03测试" + EXCEL03);
        HSSFWorkbook workbook = new HSSFWorkbook(input);
        // 获取工作表
        HSSFSheet sheetAt = workbook.getSheetAt(0);
        // 获取列（名称）
        HSSFRow row = sheetAt.getRow(0);
        if (row != null) {
            // 获取这个列有多少个单元格
            int physicalNumberOfCells = row.getPhysicalNumberOfCells();
            for (int i = 0; i < physicalNumberOfCells; i++) {
                // 获取单元格
                HSSFCell cell = row.getCell(i);
                if (cell != null){
                    // 获取单元格里的数据类型
                    CellType cellType = cell.getCellType();
                    // 获取单元格数据
                    String stringCellValue = cell.getStringCellValue();
                    System.out.print(cellType + " | " + stringCellValue);
                }
            }
        }

        // 一共有多少列
        int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows();
        for (int i = 1; i < physicalNumberOfRows; i++) {
            // 判断列不是空
            HSSFRow row1 = sheetAt.getRow(i);
            if (row1 != null){
                // 这个列有多少个单元格
                int physicalNumberOfCells = row1.getPhysicalNumberOfCells();
                for (int j = 0; j < physicalNumberOfCells; j++) {
                    HSSFCell cell = row1.getCell(j);
                    if (cell != null){
                        CellType cellType = cell.getCellType();
                        String value = "";
                        switch (cellType) {
                            case NUMERIC:
                                // 判断是否是日期
                                if (DateUtil.isCellDateFormatted(cell)){
                                    System.out.println("日期");
                                    value = String.valueOf(cell.getDateCellValue());

                                }else {
                                    System.out.println("数字");
                                    value = String.valueOf(cell.getNumericCellValue());
                                }

                                break;

                            case STRING:
                                System.out.println("字符串");
                                value = cell.getStringCellValue();
                                break;

                            case BOOLEAN:
                                System.out.println("布尔");
                                value = String.valueOf(cell.getBooleanCellValue());
                                break;

                            case _NONE:
                                System.out.println("没有");
                                break;

                            case FORMULA:
                                System.out.println("公式");
                                break;

                            case BLANK:
                                System.out.println("空白");
                                break;

                            case ERROR:
                                System.out.println("错误");
                                break;

                            default:
                                System.out.println("未知数据类型");
                        }
                        System.out.println("第" + i + "行" + j + "单元格，数据：" + value);
                    }

                }
            }
        }
        input.close();
    }

}
