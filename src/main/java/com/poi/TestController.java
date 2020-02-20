package com.poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;

@Controller
public class TestController {
    @RequestMapping("t")
    @ResponseBody
    public String t(){
        return  "1";
    }

    /**
     * 系统里面本来就有excel表格，浏览器下载该表格到本地
     * @param request
     * @param response
     */
    @RequestMapping("getExcel")
    public void getExcel(HttpServletRequest request, HttpServletResponse response){
        try{
            response.setContentType("application/octet-stream");//设置流传输
            response.addHeader("Content-Disposition", "attachment;filename="+"test.xlsx");//设置下载的文件名称
            InputStream inStream = new FileInputStream("D:/test.xlsx");
            // 循环取出流中的数据
            byte[] b = new byte[100];
            int len;
            try {
                while ((len = inStream.read(b)) > 0)
                    response.getOutputStream().write(b, 0, len);//向客户端响应
                inStream.close();//关闭流
            } catch (IOException e) {
                e.printStackTrace();
            }
        }catch (Exception e){
        }
    }

    /**
     * 在程序内部生成一个excel表格，然后下载到本地
     * @param request
     * @param response
     */
    @RequestMapping("getExcel2")
    public void getExcel2(HttpServletRequest request, HttpServletResponse response){
        try{
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
        response.setContentType("application/octet-stream");//设置流传输
        response.addHeader("Content-Disposition", "attachment;filename="+"test.xlsx");//设置下载的文件名称
        workbook.write(response.getOutputStream());
        }catch (Exception e){
        }
    }
}
