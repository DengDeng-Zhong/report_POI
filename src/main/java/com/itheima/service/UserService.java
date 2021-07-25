package com.itheima.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.itheima.mapper.UserMapper;
import com.itheima.pojo.User;
//import jxl.Workbook;
//import org.apache.poi.ss.usermodel.Workbook;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;


import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.text.SimpleDateFormat;
import java.util.*;

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
        sheet.setColumnView(4,20);

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
            label = new Label(0, rowIndex, user.getId().toString());//列脚标,行脚标,单元格中的内容 编号
            sheet.addCell(label);

            label = new Label(1, rowIndex, user.getUserName());//列脚标,行脚标,单元格中的内容用户名
            sheet.addCell(label);

            label = new Label(2, rowIndex, user.getPhone());//列脚标,行脚标,单元格中的内容手机号
            sheet.addCell(label);

            label = new Label(3, rowIndex, sdf.format(user.getHireDate()));//列脚标,行脚标,单元格中的内容 入职日期
            sheet.addCell(label);

            label = new Label(4, rowIndex, user.getAddress());//列脚标,行脚标,单元格中的内容 地址
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
}
