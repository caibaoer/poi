package com.poi;

import org.apache.commons.collections4.bag.SynchronizedSortedBag;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

@SpringBootTest
class PoiApplicationTests {

    //创建一个excel文件test.xls到D盘
    //内容  3行4列 A1:中国 D3:美国
    @Test
    void contextLoads() {
        //创建一个工作簿
        XSSFWorkbook workbook=new XSSFWorkbook();
        //创建一个表格
        XSSFSheet sheet=workbook.createSheet("world");
        //创建第一行第一列并设置“中国”
        sheet.createRow(0).createCell(0).setCellValue("中国");
        //创建第三行第四列并设置“美国”
        sheet.createRow(2).createCell(3).setCellValue("中国");
        String fileName="D://test.xls";
        try {
            FileOutputStream fo=new FileOutputStream(fileName);
            workbook.write(fo);
            workbook.close();
            fo.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


    //把D盘的test.xls文档解析，计算该文档有多少行 多少列，并把有值的单元格的值取出来

    @Test
    void contextLoads2() throws  Exception{
        FileInputStream fi=new FileInputStream("D://test.xls");
        XSSFWorkbook workbook =new XSSFWorkbook(fi);
        //获取表格
        XSSFSheet sheet=workbook.getSheetAt(0);
        int rowNum=sheet.getLastRowNum();
        for(int i=0;i<rowNum;i++){
          XSSFRow row= sheet.getRow(i);
          int cellNum=row.getLastCellNum();
          for(int b=0;b<cellNum;b++){
             XSSFCell cell= row.getCell(b);
              System.out.println(cell.getStringCellValue());
          }

        }


    }

}
