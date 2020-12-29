package com.feng;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class ExcelRead {

    static  String Path = "D:\\javaT\\feng-POI\\src\\main\\resources\\";
    @Test
    public void testRead03() throws IOException {
        //1.获取文件流
        FileInputStream inputStream = new FileInputStream(Path+"统计表.xls");
        //2.创建工作簿  excel能操作的，都可以操作
        Workbook workbook = new HSSFWorkbook(inputStream);
        //3.拿到那个表
        Sheet sheetAt = workbook.getSheetAt(0);
        //4.拿到那个行
        Row row = sheetAt.getRow(0);
        //5.得到列
        Cell cell = row.getCell(1);
        //得到值
        //读取值得时候，要注意类型，不同得类型，有不同得方法。
        double value = cell.getNumericCellValue();
        System.out.println(value);
        inputStream.close();
    }

    @Test
    public void testRead07() throws IOException {
        //1.获取文件流
        FileInputStream inputStream = new FileInputStream(Path+"统计表07.xlsx");
        //2.创建工作簿  excel能操作的，都可以操作
        Workbook workbook = new XSSFWorkbook(inputStream);
        //3.拿到那个表
        Sheet sheetAt = workbook.getSheetAt(0);
        //4.拿到那个行x
        Row row = sheetAt.getRow(0);
        //5.得到列
        Cell cell = row.getCell(1);
        //得到值
        //读取值得时候，要注意类型，不同得类型，有不同得方法。
        double value = cell.getNumericCellValue();
        System.out.println(value);
        inputStream.close();
    }

    @Test
    public void testReadTitle03() throws IOException {
        //1.获取文件流
        FileInputStream inputStream = new FileInputStream(Path+"统计表.xls");
        //2.创建工作簿  excel能操作的，都可以操作
        Workbook workbook = new HSSFWorkbook(inputStream);
        //3.拿到表
        Sheet sheet = workbook.getSheetAt(0);
        //获得标题内容
        Row rowTile = sheet.getRow(0);
        if (rowTile!=null){
            //获得总数量
            int count = rowTile.getPhysicalNumberOfCells();
            for (int rowNum = 0; rowNum <count ; rowNum++) {
                Cell cell = rowTile.getCell(rowNum);
                if (cell!=null){
                    System.out.print(cell.getStringCellValue()+"|");
                }
            }
        }
        System.out.println();
        //读取表中的内容
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum <rowCount ; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData!=null){
                //读取列
                int cellCount = rowTile.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum <cellCount ; cellNum++) {
                    System.out.print("["+(rowNum+1)+"-"+(cellNum-1)+"]");
                    Cell cell = rowData.getCell(cellNum);
                    if (cell!=null){
                        int cellType = cell.getCellType();
                        String cellValue = "";
                        switch (cellType){
                            case HSSFCell
                                    .CELL_TYPE_STRING : //字符串
                                System.out.print("[STRING]");
                                cellValue = cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN : //布尔值
                                System.out.print("[BOOLEAN]");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK:
                                System.out.print("[BLANK]");
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC:
                                System.out.print("[判断：日期or数字]");
                                if (HSSFDateUtil.isCellDateFormatted(cell)){
                                    System.out.print("[Date]");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                }else{
                                    System.out.print("是数字：转换成字符串输出！！");
                                    cell.setCellValue(HSSFCell.CELL_TYPE_STRING);
                                    cell.toString();
                                }
                            case HSSFCell.CELL_TYPE_ERROR:
                                System.out.print("[ERROR]");
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }
        inputStream.close();
        //得到值
        //读取值得时候，要注意类型，不同得类型，有不同得方法。

    }


    @Test
    public void testFormula() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(Path+"公式表.xls");
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        Row row = sheet.getRow(4);
        Cell cell = row.getCell(0);

        //拿到计算公式
       FormulaEvaluator FormulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook)workbook);
       //输出单元格的内容
        int cellType = cell.getCellType();
        switch (cellType){
            case Cell.CELL_TYPE_FORMULA:
                String formula = cell.getCellFormula();
                //拿到公式
                System.out.println(formula);
                //计算
                CellValue evaluate = FormulaEvaluator.evaluate(cell);
                String cellValue = evaluate.formatAsString();
                System.out.println(cellValue);
                break;
        }
    }
}