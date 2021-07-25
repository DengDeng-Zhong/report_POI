package cn.targetpath.test;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

/**
 * 导入数据
 *
 * @author ISC
 * @version V1.0
 * @date 2021/7/24 15:30
 */
public class POIDemo3 {
    public static void main(String[] args) throws Exception{
        // 用户名 手机号 省份 城市 工资 入职日期 出生日期 现住址

        // src/main/resources/excel_template/用户导入测试数据.xlsx
        Workbook workbook = new XSSFWorkbook(new FileInputStream("src/main/resources/excel_template/用户导入测试数据.xlsx"));
        //获取第一个工作表
        Sheet sheet = workbook.getSheetAt(0);
        //读取工作表中的内容
        int rowNum = sheet.getLastRowNum();
        Row row = null;
        for (int i = 1; i <= rowNum; i++) {
            row = sheet.getRow(i);
            String username = row.getCell(0).getStringCellValue();
            String phone = row.getCell(1).getStringCellValue();
            String province = row.getCell(2).getStringCellValue();
            String city = row.getCell(3).getStringCellValue();
            String salary = row.getCell(4).getStringCellValue();
            String hireDate = row.getCell(5).getStringCellValue();
            String birthDay = row.getCell(6).getStringCellValue();
            String address = row.getCell(7).getStringCellValue();

        }


        workbook.close();
    }
}
