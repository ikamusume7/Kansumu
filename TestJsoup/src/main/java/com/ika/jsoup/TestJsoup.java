/**
 * 	Copyright 2018 www.arbexpress.cn
 * 
 * 	All right reserved
 * 
 * 	Create on 2018年4月28日下午1:46:54
 */
package com.ika.jsoup;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

/**
 * @File: TestJsoup.java
 * @Author: yangLaiYao
 * @Description: TODO
 */
public class TestJsoup {
    
    public static void main(String[] args) throws IOException {
        Document doc = Jsoup.connect("http://wiki.joyme.com/blhx/%E7%BD%97%E6%81%A9").data("query", "Java")
                        .userAgent("Mozilla")
                        .cookie("auth", "token")
                        .timeout(3000)
                        .post();
//        System.out.println(doc);
        List<Element> tables = doc.select("table[class^=wikitable]");
        for(Element table : tables) {
            List<Element> trs = table.select("tr");
            for(Element tr : trs) {
                List<Element> tds = tr.select("td");
                for(Element td : tds) {
                    System.out.println(td.text());
                }
            }
        }
        //创建excel工作簿
        HSSFWorkbook workbook=new HSSFWorkbook();
        //创建工作表sheet
        HSSFSheet sheet=workbook.createSheet();
        //写入数据
        int i = 0;
        for(Element table : tables) {
            List<Element> trs = table.select("tr");
            for(Element tr : trs) {
                HSSFRow row = sheet.createRow(i++);
                List<Element> tds = tr.select("td");
                int j = 0;
                for(Element td : tds) {
                    HSSFCell cell= row.createCell(j++);
                    if(StringUtils.isNotBlank(td.select("img").attr("src"))) {
                        cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                        cell.setCellFormula("HYPERLINK(\"" + td.select("img").attr("src")+ "\",\"" + td.text() + "\")");
                        HSSFCellStyle linkStyle = workbook.createCellStyle();
                        HSSFFont cellFont= workbook.createFont();
                        cellFont.setUnderline((byte) 1);
                        cellFont.setColor(HSSFColor.BLUE.index);
                        linkStyle.setFont(cellFont);
                        cell.setCellStyle(linkStyle);
                    } else {
                        cell.setCellValue(td.text());
                    }
                }
            }
            i = i + 2;
        }
        //创建excel文件
        File file=new File("e://luo_en.xls");
        try {
            file.createNewFile();
            //将excel写入
            FileOutputStream stream= FileUtils.openOutputStream(file);
            workbook.write(stream);
            stream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
