package top.jipson.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriteTest {
    String PATH = "F:\\Study\\exceldemo\\";

    /**
     * 03版excel最大行数为65536
     * @throws Exception
     */
    @Test
    public void testWrite03() throws Exception {
        //1.创建一个工作簿（excel文件）对象
        HSSFWorkbook workbook = new HSSFWorkbook();
        //2.创建一个工作表对象
        HSSFSheet sheet = workbook.createSheet("统计表");

        //3.创建一行,第一行（姓名，年龄，入职日期）
        // 行号为第1行
        HSSFRow row1 = sheet.createRow(0);

        //创建单元格并赋值
        //创建一个单元格(第1行，第1列)
        HSSFCell celll1 = row1.createCell(0);
        //设置单元格值(第1行，第1列)
        celll1.setCellValue("姓名");

        //创建一个单元格(第1行，第2列)
        HSSFCell cell12 = row1.createCell(1);
        //设置单元格值(第1行，第1列)
        cell12.setCellValue("年龄");

        HSSFCell cell13 = row1.createCell(2);
        cell13.setCellValue("入职日期");

        //4.创建一行,第二行（姓名，年龄，入职日期）
        // 行号为第1行
        HSSFRow row2 = sheet.createRow(1);

        //创建单元格并赋值
        //创建一个单元格(第2行，第1列)
        HSSFCell cell2l = row2.createCell(0);
        //设置单元格值(第1行，第1列)
        cell2l.setCellValue("张三");

        //创建一个单元格(第2行，第2列)
        HSSFCell cell22 = row2.createCell(1);
        //设置单元格值(第1行，第1列)
        cell22.setCellValue(18);

        HSSFCell cell23 = row2.createCell(2);
        String time = new DateTime().toString("yyyy-MM-dd");
        cell23.setCellValue(time);


        //新建一个输出流，数据输出到哪里，注意后缀名位xls
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "统计表.xls");
        //将workbook写到输出流中
        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();

        System.out.println("完成");

    }

    /**
     * 07版excel最大行数1048576
     * @throws Exception
     */
    @Test
    public void testWrite07() throws Exception {
        //1.创建一个工作簿（excel文件）对象
        XSSFWorkbook workbook = new XSSFWorkbook();
        //2.创建一个工作表对象
        XSSFSheet sheet = workbook.createSheet("统计表");

        //3.创建一行,第一行（姓名，年龄，入职日期）
        // 行号为第1行
        XSSFRow row1 = sheet.createRow(0);

        //创建单元格并赋值
        //创建一个单元格(第1行，第1列)
        XSSFCell celll1 = row1.createCell(0);
        //设置单元格值(第1行，第1列)
        celll1.setCellValue("姓名");

        //创建一个单元格(第1行，第2列)
        XSSFCell cell12 = row1.createCell(1);
        //设置单元格值(第1行，第1列)
        cell12.setCellValue("年龄");

        XSSFCell cell13 = row1.createCell(2);
        cell13.setCellValue("入职日期");

        //4.创建一行,第二行（姓名，年龄，入职日期）
        // 行号为第1行
        XSSFRow row2 = sheet.createRow(1);

        //创建单元格并赋值
        //创建一个单元格(第2行，第1列)
        XSSFCell cell2l = row2.createCell(0);
        //设置单元格值(第1行，第1列)
        cell2l.setCellValue("张三");

        //创建一个单元格(第2行，第2列)
        XSSFCell cell22 = row2.createCell(1);
        //设置单元格值(第1行，第1列)
        cell22.setCellValue(18);

        XSSFCell cell23 = row2.createCell(2);
        String time = new DateTime().toString("yyyy-MM-dd");
        cell23.setCellValue(time);


        //新建一个输出流，数据输出到哪里，注意后缀名位xlsx
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "统计表.xlsx");
        //将workbook写到输出流中
        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();

        System.out.println("完成");

    }

    @Test
    public void testBatchWrite03(){
        //开始时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建表,使用默认的sheet名
        HSSFSheet sheet = workbook.createSheet();
        //批量生成数据,生成03版最大的数据量
        for(int rowNum = 0; rowNum < 65536; rowNum++){//行
            HSSFRow row = sheet.createRow(rowNum);
            for(int cellNum = 0; cellNum < 10; cellNum++){//列
                HSSFCell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        try {
            FileOutputStream fileOutputStream = new FileOutputStream(PATH + "批量写入表.xls");
            workbook.write(fileOutputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        //结束时间
        long end = System.currentTimeMillis();
        //计算时间差
        System.out.println((double) (end-begin)/1000);//执行时间（秒） 测试结果：2.891
    }

    @Test
    public void testBatchWrite07(){
        //开始时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //创建表,使用默认的sheet名
        XSSFSheet sheet = workbook.createSheet();
        //批量生成数据,生成03版最大的数据量
        for(int rowNum = 0; rowNum < 65537; rowNum++){//行
            XSSFRow row = sheet.createRow(rowNum);
            for(int cellNum = 0; cellNum < 10; cellNum++){//列
                XSSFCell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        try {
            FileOutputStream fileOutputStream = new FileOutputStream(PATH + "批量写入表.xlsx");
            workbook.write(fileOutputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        //结束时间
        long end = System.currentTimeMillis();
        //计算时间差
        System.out.println((double) (end-begin)/1000);//执行时间（秒）执行结果：17.213
    }


    /**
     * 优化07版excel写入，使用SXSSF
     */
    @Test
    public void testBatchWriteOptimize07(){
        //开始时间
        long begin = System.currentTimeMillis();

        //创建一个工作簿
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        //创建表,使用默认的sheet名
        SXSSFSheet sheet = workbook.createSheet();
        //批量生成数据,生成03版最大的数据量
        for(int rowNum = 0; rowNum < 65537; rowNum++){//行
            SXSSFRow row = sheet.createRow(rowNum);
            for(int cellNum = 0; cellNum < 10; cellNum++){//列
                SXSSFCell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        try {
            FileOutputStream fileOutputStream = new FileOutputStream(PATH + "优化批量写入表.xlsx");
            workbook.write(fileOutputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        //清除临时文件
//        workbook.dispose();（不用清除，运行完成后没有看到临时文件，可能时版本原因）
        //结束时间
        long end = System.currentTimeMillis();
        //计算时间差
        System.out.println((double) (end-begin)/1000);//执行时间（秒）执行结果：优化前17.213--->优化后2.286
    }
}
