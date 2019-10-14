package com.baizhi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.IOException;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiApplicationTests {

    @Test
    public void contextLoads() {
        //创建 Excel 工作薄对象
        HSSFWorkbook workbook = new HSSFWorkbook();

        //创建工作表 HSSFSheet
        HSSFSheet sheet = workbook.createSheet("用户表");
        //创建标题行
        HSSFRow row = sheet.createRow(0);
        String[] title = {"id", "姓名", "生日"};
        for (int i = 0; i < title.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellValue(title[i]);
        }
        try {
            workbook.write(new File("D:/a.xls"));
            System.out.println("成功");
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

}
