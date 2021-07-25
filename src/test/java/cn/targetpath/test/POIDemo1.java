package cn.targetpath.test;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 需求: 创建一个低版本的excel.并向其中的一个单元格随便写一句话
 * @author ISC
 * @version V1.0
 * @date 2021/7/24 15:07
 */
public class POIDemo1 {
    public static void main(String[] args) throws IOException {
        // 创建一个全新工作簿
       Workbook workbook = new HSSFWorkbook();
        // 在工作簿中创建新的工作表
        Sheet sheet = workbook.createSheet("POI操作Excel");
        // 在工作表中创建行
        Row row = sheet.createRow(0);
        //在行中创建单元格
        Cell cell = row.createCell(1);
        //在单元格中写入内容
        cell.setCellValue("创建一个低版本的excel.并向其中的一个单元格随便写一句话");

        workbook.write(new FileOutputStream("d:/test.xls"));


    }
}
