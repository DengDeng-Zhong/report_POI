package cn.targetpath.test;

import com.itheima.pojo.User;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

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
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");

        // src/main/resources/excel_template/用户导入测试数据.xlsx
        Workbook workbook = new XSSFWorkbook(new FileInputStream("src/main/resources/excel_template/用户导入测试数据.xlsx"));
        //获取第一个工作表
        Sheet sheet = workbook.getSheetAt(0);
        //读取工作表中的内容
        int rowNum = sheet.getLastRowNum();
        Row row = null;
        User user = null;
        for (int i = 1; i <= rowNum; i++) {
            row = sheet.getRow(i);
            String username = row.getCell(0).getStringCellValue();
            String phone = row.getCell(1).getStringCellValue();
            String province = row.getCell(2).getStringCellValue();
            String city = row.getCell(3).getStringCellValue();
            Integer salary = ((Double)row.getCell(4).getNumericCellValue()).intValue();
            Date hireDate = simpleDateFormat.parse(row.getCell(5).getStringCellValue());
            Date birthDay = simpleDateFormat.parse(row.getCell(6).getStringCellValue());
            String address = row.getCell(7).getStringCellValue();
            user.setUserName(username);
            user.setPhone(phone);
            user.setProvince(province);
            user.setCity(city);
            user.setSalary(salary);
            user.setHireDate(hireDate);
            user.setBirthday(birthDay);
            user.setAddress(address);

            user.toString();
        }


        workbook.close();
    }
}
