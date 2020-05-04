package top.jipson.poi;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.util.Date;

public class ExcelReadTest {
    String PATH = "F:\\Study\\exceldemo\\";
    @Test
    public void testRead03() throws Exception {
        
        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "统计表.xls");

        //1.通过输入流创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
        //2.得到sheet表，第一个sheet
        HSSFSheet sheet = workbook.getSheetAt(0);
        //3.得到行，第一行
        HSSFRow row = sheet.getRow(0);
        //4.得到列，
        // 第一行第一列
        HSSFCell cell11 = row.getCell(0);
        // 第一行第二列
        HSSFCell cell12 = row.getCell(1);
        //打印输出单元格内容
        System.out.print(cell11.getStringCellValue());
        System.out.print("|");
        System.out.print(cell12.getStringCellValue());//注意值的类型
        //关闭流
        fileInputStream.close();
    }

    @Test
    public void testRead07() throws Exception {

        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "统计表.xlsx");

        //1.通过输入流创建一个工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        //2.得到sheet表，第一个sheet
        XSSFSheet sheet = workbook.getSheetAt(0);
        //3.得到行，第一行
        XSSFRow row = sheet.getRow(0);
        //4.得到列，
        // 第一行第一列
        XSSFCell cell11 = row.getCell(0);
        // 第一行第二列
        XSSFCell cell12 = row.getCell(1);
        //打印输出单元格内容
        System.out.print(cell11.getStringCellValue());
        System.out.print("|");
        System.out.print(cell12.getStringCellValue());//注意值的类型
        //关闭流
        fileInputStream.close();
    }

    /**
     * 测试单元格类型
     * @throws Exception
     */
    @Test
    public void testCellType() throws Exception {
        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "明细表.xls");

        //1.通过输入流创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);

        HSSFSheet sheet = workbook.getSheetAt(0);

        //获取标题内容
        HSSFRow rowTitle = sheet.getRow(0);
        if(rowTitle!=null){
            int cellCount = rowTitle.getPhysicalNumberOfCells();//获取多少列
            for(int cellNum = 0; cellNum < cellCount; cellNum++ ){
                HSSFCell cell = rowTitle.getCell(cellNum);
                if(cell!=null){
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + "|");
                }
            }
            System.out.println();//换行
        }

        //获取表内容
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {//表内容从第二行开始（第一行为表头）
            HSSFRow rowData = sheet.getRow(rowNum);//第rowNum行数据
            if(rowData!=null){
                //读取列
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    System.out.print("["+(rowNum+1)+"-"+(cellNum+1)+"]");//虚拟数据

                    HSSFCell cell = rowData.getCell(cellNum);//拿到单元格数
                    //匹配列的数据类型
                    if(cell!=null){
                        int cellType = cell.getCellType();//获取单元格类型
                        String cellValue = "";
                        switch (cellType){
                            case HSSFCell.CELL_TYPE_STRING:
                                System.out.print("[String]");
                                cellValue = cell.getStringCellValue();//读取值
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN:
                                System.out.print("[BOOLEAN]");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK:
                                System.out.print("[BLANK]");
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC: //数字（日期、普通数字）
                                System.out.print("[NUMERIC]");
                                if(HSSFDateUtil.isCellDateFormatted(cell)){//日期
                                    System.out.print("[日期]");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                }else{
                                    //普通数字，防止数字过长
                                    System.out.print("[数字转换为字符串输出]");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case HSSFCell.CELL_TYPE_ERROR:
                                System.out.print("[数据类型错误]");
                                break;
                        }
                        System.out.println(cellValue);
                    }

                }
                System.out.println();//换行
            }
        }
        fileInputStream.close();
    }

    /**
     * 测试单元格公式
     * @throws Exception
     */
    @Test
    public void testFormula() throws Exception {
        FileInputStream inputStream = new FileInputStream(PATH + "公式.xls");
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        HSSFSheet sheet = workbook.getSheetAt(0);
        //公式在所在单元格为A5
        HSSFRow row = sheet.getRow(4);
        HSSFCell cell = row.getCell(0);
        int cellType = cell.getCellType();
        switch (cellType){
            case Cell.CELL_TYPE_FORMULA://单元格的数据类型为公式
                String formula = cell.getCellFormula();
                System.out.println(formula);//则把公式打印出来
                //计算
//                CellValue evaluate = FormulaEvaluator.evaluate(cell);
//                String cellValue = evaluate.formatAsString();
                double cellValue = cell.getNumericCellValue();
                System.out.println(cellValue);
                break;
        }
    }
}
