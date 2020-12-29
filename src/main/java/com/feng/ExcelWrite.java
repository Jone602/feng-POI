package com.feng;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWrite {

    static  String Path = "D:\\javaT\\feng-POI\\src\\main\\resources\\";
//    03版 excel表的生成方式 new HSSFWorkbook()
    @Test
    public void test03() throws Exception {
//        1.创建工作簿
        Workbook workbook = new HSSFWorkbook();
//        2.创建工作表
        Sheet sheet = workbook.createSheet("工作表名");
//        3.创建行
        Row row1 = sheet.createRow(0);
//        4.创建单元
//        5.写入数据
        Cell cell0 = row1.createCell(0);
        cell0.setCellValue("今日新增");
        Cell cell1 = row1.createCell(1);
        cell1.setCellValue(666);
//        3.创建第二行
        Row row2 = sheet.createRow(1);
//        4.创建单元
//        5.写入数据
        Cell cell3 = row2.createCell(0);
        cell3.setCellValue("时间");
        Cell cell4 = row2.createCell(1);
        String  time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell4.setCellValue(time);
        //生成一张表（IO流技术），03版本就是以xls结尾
        FileOutputStream fileOutputStream = new FileOutputStream(Path+"统计表.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("生成完毕！！！");
    }
//    07版excel表的生成方式：XSSFWorkbook()
    @Test
    public void test07() throws Exception {
//        1.创建工作簿
        Workbook workbook = new XSSFWorkbook();
//        2.创建工作表
        Sheet sheet = workbook.createSheet("工作表名");
//        3.创建行
        Row row1 = sheet.createRow(0);
//        4.创建单元
//        5.写入数据
        Cell cell0 = row1.createCell(0);
        cell0.setCellValue("今日新增");
        Cell cell1 = row1.createCell(1);
        cell1.setCellValue(666);
//        3.创建第二行
        Row row2 = sheet.createRow(1);
//        4.创建单元
//        5.写入数据
        Cell cell3 = row2.createCell(0);
        cell3.setCellValue("时间");
        Cell cell4 = row2.createCell(1);
        String  time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell4.setCellValue(time);
        //生成一张表（IO流技术），03版本就是以xls结尾
        FileOutputStream fileOutputStream = new FileOutputStream(Path+"统计表07.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("生成完毕！！！");
    }

    @Test
    public void testWrite03BigData() throws IOException {
        long begin = System.currentTimeMillis();
        //创建工作簿
        Workbook workbook = new HSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum <65536 ; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum <10 ; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        System.out.println("运行完毕！！");
        FileOutputStream stream = new FileOutputStream(Path+"testWrite03BigData.xls");
        workbook.write(stream);
        stream.close();
//        Row row = sheet.createRow(0);
//        Cell cell = row.createCell(0);
//        cell.setCellValue("110");
        long end = System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);

    }
    //    SXSSFWorkbook()
    @Test
    public void testWrite07BigData() throws IOException {
        long begin = System.currentTimeMillis();
        //创建工作簿
        Workbook workbook = new XSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum <100000 ; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum <10 ; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("运行完毕！！");
        FileOutputStream stream = new FileOutputStream(Path+"testWrite07BigData.xlsx");
        workbook.write(stream);
        stream.close();
        long end = System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);
    }
//    SXSSFWorkbook()  优化版
    @Test
    public void testWrite07BigDataSXSSF() throws IOException {
        long begin = System.currentTimeMillis();
        //创建工作簿
        Workbook workbook = new SXSSFWorkbook(100000);
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum <100000 ; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum <10 ; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("运行完毕！！");
        FileOutputStream stream = new FileOutputStream(Path+"testWrite07BigData.xlsx");
        workbook.write(stream);
        stream.close();
        //清除临时文件。
        ((SXSSFWorkbook) workbook).dispose();
        long end = System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);
    }
}