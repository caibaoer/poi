package com.poi;



import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.extractor.ExcelExtractor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.UUID;


public class POITest {

    public static void main(String[] args) {

        toExcel();
        //fromExcel0();
        //fromExcel1();
        //fromExcel2();
        //t();
        //getFromexcel();
        //getText();
    }

    /**
     * 程序组装excel文件到硬盘
     */
    public static void toExcel() {
        //创建一个工作簿
        XSSFWorkbook workbook=new XSSFWorkbook();
        //创建一个表格
        XSSFSheet sheet=workbook.createSheet("world");
        //统一设置默认的列宽(一般情况下，我们一个excel表格的有些列要比其他列宽些，这种情况下就需要另外针对专门的列设置宽度，使用sheet.setColumnWidth(2, 20 * 256))  这里的2是列的索引值
        sheet.setDefaultColumnWidth(20);
        sheet.setDefaultRowHeightInPoints(20);
        //创建第一行
        XSSFRow row0=sheet.createRow(0);
        //单独设置行高
        row0.setHeightInPoints(30);
        //在第一行创建第一列
        XSSFCell cell0=row0.createCell(0);
        //在第一行第一列设置内容为“中国”
        cell0.setCellValue("中国CHINA \n hello");
        XSSFCellStyle cellstyle=workbook.createCellStyle();
        cellstyle.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 设置内容水平对齐方式：居中
        cellstyle.setVerticalAlignment(XSSFCellStyle.ALIGN_CENTER);//// 设置内容垂直对齐方式：居中
        cellstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //设置背景颜色（与setFillForegroundColor一起配合使用）  -------》设置颜色建议使用IndexedColors
        cellstyle.setFillForegroundColor(IndexedColors.YELLOW.index);// 设置前背景颜色（与setFillPattern一起配合使用）  -------》设置颜色建议使用IndexedColors
        cellstyle.setBorderBottom(XSSFCellStyle.BORDER_THIN); //设置下边框 小实线
        cellstyle.setBottomBorderColor(IndexedColors.BLUE.getIndex());//设置下边框颜色
        cellstyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);//设置左边框 小实线
        cellstyle.setLeftBorderColor(IndexedColors.RED.getIndex());//设置左边框 颜色
        cellstyle.setBorderTop(XSSFCellStyle.BORDER_THIN);//设置上边框 小实线
        cellstyle.setTopBorderColor(IndexedColors.BROWN.getIndex());//设置上边框 颜色
        cellstyle.setBorderRight(XSSFCellStyle.BORDER_THIN);//设置右边框 小实线
        cellstyle.setRightBorderColor(IndexedColors.RED.getIndex());//设置右边框 颜色
        cellstyle.setWrapText(true);//设置换行，只有这里设置了，在cellValue里面的 "\n"  才起效果
        Font font=workbook.createFont(); //设置字体
        font.setFontName("楷体");//设置字体
        font.setFontHeightInPoints((short) 12);//设置字体大小  一般12就够了
        font.setColor(IndexedColors.RED.index);//设置字体颜色    -------》设置颜色建议使用IndexedColors
        cellstyle.setFont(font);
        cell0.setCellStyle(cellstyle);
        //创建第三行第四列并设置“美国”
        sheet.createRow(2).createCell(3).setCellValue("中国");
        XSSFRow row1=sheet.createRow(5);
        //row1.createCell(0).setCellValue("123");
        row1.setRowStyle(cellstyle);//行也可以使用样式

        Row row3=sheet.createRow(3);
        row3.createCell(1).setCellValue(new Date());//直接设置Date的话，excel显示的是43880.9129211806这种值，不是时间字符串
        Cell cell2=row3.createCell(2);
        cell2.setCellValue(new Date());
        CellStyle csy=workbook.createCellStyle();//设置时间格式
        XSSFDataFormat df = workbook.createDataFormat();
        csy.setDataFormat(df.getFormat("yyyy/MM/dd"));
        cell2.setCellStyle(csy);

        String excelFileName=UUID.randomUUID().toString();
        String fileName="D://"+excelFileName+".xlsx";
        try {
            FileOutputStream fo=new FileOutputStream(fileName);
            workbook.write(fo);
            workbook.close();
            fo.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 从硬盘读取文件到内存解析
     * 有几种方式都可以的
     * 1、通过文件流的形式fromExcel0
     * 2、查看源码得知，直接构造函数传入文件地址字符串也可以，自己测试时可以的。fromExcel1
     * 3、查看源码得知，直接构造函数传入File类型的文件也可以，自己测试时可以的。fromExcel2
     */
    public static void fromExcel0() {
        String file="D://test.xlsx";
        try {
            FileInputStream fileInputStream=new FileInputStream(file);
            XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);
            int sheetNumber=workbook.getNumberOfSheets();//工作表的个数
            Sheet sheet=workbook.getSheetAt(0);
            Row row=sheet.getRow(0);
            int physicalNumberOfCells=row.getPhysicalNumberOfCells();//获得实际的有内容的列的列数
            int physicalNumberOfRows=sheet.getPhysicalNumberOfRows();//获得实际的有内容的行的行数
            int firstRowNum=sheet.getFirstRowNum();
            int lastRowNum=sheet.getLastRowNum();//这个是最后一行的索引值，+1就是该表有多少行（包括有数据无数据）
            int lastCellNum=row.getLastCellNum();//返回的是最后一列的列数，即等于总列数  不用+1   本来的值就是该表的该行有多少列（包括有数据无数据）
            System.out.println(1);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void fromExcel1() {
        String file="D://test.xlsx";
        try {
            XSSFWorkbook workbook=new XSSFWorkbook(file);
            System.out.println(1);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static void fromExcel2() {
        String file="D://test.xlsx";
        try {
            File filedemo=new File(file);
            XSSFWorkbook workbook=new XSSFWorkbook(file);
            System.out.println(1);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 循环获取excel得行和列数据
     */
    public  static  void  getFromexcel(){
        try{
            String file="D://test2.xlsx";
            XSSFWorkbook workbook=new XSSFWorkbook(file);
            Sheet sheet=workbook.getSheetAt(0);
            for(int rowIndex=0;rowIndex<=sheet.getLastRowNum();rowIndex++){
                Row row=sheet.getRow(rowIndex);
                if(row==null){
                  continue;
                }
                for(int cellIndex=0;cellIndex<row.getLastCellNum();cellIndex++){
                    Cell cell=row.getCell(cellIndex);
                    if(cell==null){
                        continue;
                    }
                    System.out.println(getValue(cell));
                }
            }
        }
        catch (Exception e){
            System.out.println(1);
        }

    }

    /**
     * 获取值
     * @param cell
     * @return
     */
    public static  String getValue(Cell cell){
        DateFormat format=new SimpleDateFormat("yyyy-MM-dd");
        if(cell.getCellType()==Cell.CELL_TYPE_STRING){
            return String.valueOf(cell.getStringCellValue());
        }
        /**
         * 通过查询资料发现，poi在Cell.CELL_TYPE_NUMERIC中又具体区分了类型，Date类型就是其中一种，把代码再做处理
         */
        if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                Date date = cell.getDateCellValue();
                return  format.format(date);
            } else {
                return  String.valueOf(cell.getNumericCellValue());
            }
        }
        if(cell.getCellType()==Cell.CELL_TYPE_BOOLEAN){
            return String.valueOf(cell.getBooleanCellValue());
           // return cell.getBooleanCellValue();
        }
        if(HSSFDateUtil.isCellDateFormatted(cell)){
            Date date = cell.getDateCellValue();
            return format.format(date);
        }
        return "";

    }

    /**
     * 抽取text
     */
    public static  void getText(){
        try{
            String file="D://test2.xlsx";
            FileInputStream in=new FileInputStream(file);
            XSSFWorkbook workbook=new XSSFWorkbook(in);
            ExcelExtractor extractor=new XSSFExcelExtractor(workbook);
            extractor.setIncludeSheetNames(false);//不需要sheet页名字
            System.out.println(extractor.getText());
        }catch (Exception e){

        }
    }

}
