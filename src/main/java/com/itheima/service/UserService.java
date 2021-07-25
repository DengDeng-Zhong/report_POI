package com.itheima.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.itheima.mapper.UserMapper;
import com.itheima.pojo.User;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;


import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 *
 */
@Service
public class UserService {

    @Autowired
    private UserMapper userMapper;

    private SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");

    public List<User> findAll() {
        return userMapper.selectAll();
    }

    public List<User> findPage(Integer page, Integer pageSize) {
        PageHelper.startPage(page,pageSize);  //开启分页
        Page<User> userPage = (Page<User>) userMapper.selectAll(); //实现查询
        return userPage.getResult();
    }

    public void downLoadXlsByJxl(HttpServletResponse response) throws Exception{
        //编号,姓名,手机号,入职日期,现住址
        ServletOutputStream outputStream = response.getOutputStream();

        //创建一个全新的工作簿
        WritableWorkbook workbook = Workbook.createWorkbook(outputStream);
        //创建一个工作表
        WritableSheet sheet = workbook.createSheet("一个JXL入门", 0);

        Label label = null;

        //调整列宽
        sheet.setColumnView(0,5);
        sheet.setColumnView(1,8);
        sheet.setColumnView(2,15);
        sheet.setColumnView(3,15);
        sheet.setColumnView(4,30);

        //处理标题
        String[] titles = new String[]{"编号","姓名","手机号","入职日期","现住址"};
        for (int i = 0; i < titles.length; i++) {
            label = new Label(i,0, titles[i]);//列脚标,行脚标,单元格中的内容
            sheet.addCell(label);
        }

        //查询所有用户数据
        List<User> users = userMapper.selectAll();
        int rowIndex = 1;
        for (User user: users ) {
            //列脚标,行脚标,单元格中的内容 编号
            label = new Label(0, rowIndex, user.getId().toString());
            sheet.addCell(label);
//列脚标,行脚标,单元格中的内容用户名
            label = new Label(1, rowIndex, user.getUserName());
            sheet.addCell(label);
//列脚标,行脚标,单元格中的内容手机号
            label = new Label(2, rowIndex, user.getPhone());
            sheet.addCell(label);
//列脚标,行脚标,单元格中的内容 入职日期
            label = new Label(3, rowIndex, sdf.format(user.getHireDate()));
            sheet.addCell(label);
//列脚标,行脚标,单元格中的内容 地址
            label = new Label(4, rowIndex, user.getAddress());
            sheet.addCell(label);

            rowIndex++;
        }

        //文件的导出  一个流(outputstream)两个头(文件的打开方式in-line attachment,文件的下载时mime-type类型)
        String fileName = "一个JXL入门.xls";
        response.setHeader("content-disposition","attachment:filename="+new String(fileName.getBytes(),"ISO8859-1"));
        response.setContentType("application/vnd.ms-excel");
        workbook.write();
        workbook.close();
        outputStream.close();

    }

    public void uploadExcel(MultipartFile file) throws Exception{
        // 用户名 手机号 省份 城市 工资 入职日期 出生日期 现住址

        // src/main/resources/excel_template/用户导入测试数据.xlsx
        org.apache.poi.ss.usermodel.Workbook workbook = new XSSFWorkbook(new FileInputStream("src/main/resources/excel_template/用户导入测试数据.xlsx"));
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
            Date hireDate = sdf.parse(row.getCell(5).getStringCellValue());
            Date birthDay = sdf.parse(row.getCell(6).getStringCellValue());
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
