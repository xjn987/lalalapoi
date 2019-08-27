package com.baizhi;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiApplicationTests {

    @Test
    public void contextLoads() {
        //创建工作溥
        HSSFWorkbook workbook = new HSSFWorkbook();
        //通过工作溥创建工作表
        HSSFSheet sheet = workbook.createSheet("测试");
        //通过工作表创建行
        HSSFRow row = sheet.createRow(0);
        //通过行创建单元格
        HSSFCell cell = row.createCell(0);
        //给单元格赋值
        cell.setCellValue("第一个单元格");
        //导出文件
        try {
            workbook.write(new FileOutputStream(new File("D:/a.xls")));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void name() {
        //创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //根据工作簿创建工作表
        HSSFSheet sheet = workbook.getSheet("on");
        //设置单元格宽度
        sheet.setColumnWidth(2,15*256);

        //设置日期格式
        HSSFDataFormat dataFormat = workbook.createDataFormat();
        short format = dataFormat.getFormat("yyyy年mm月dd日");
        //把日期格式交个样式对象
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(format);
        //创建单元格样式对象
        HSSFCellStyle fontStyle = workbook.createCellStyle();
        fontStyle.setAlignment(HorizontalAlignment.CENTER);
        //创建字体样式对象
        HSSFFont font = workbook.createFont();
        font.setBold(true);//字体加粗
        font.setColor(Font.COLOR_RED);//字体颜色
        font.setItalic(true);//字体样式
        font.setFontName("楷体");//字体类型
        fontStyle.setFont(font);//放回对象

        //创建标题行
        HSSFRow tiltRow = sheet.createRow(0);
        String[] str = {"id","姓名","生日"};
        for (int i=0;i<str.length;i++){
            HSSFCell cell = tiltRow.createCell(i);//的到第I个单元格
            cell.setCellStyle(fontStyle);//设置样式
            cell.setCellValue(str[i]);//往单元格填充字段
        }
    }
}
