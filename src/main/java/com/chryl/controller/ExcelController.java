package com.chryl.controller;

import com.chryl.po.JxLoanPortfolio;
import com.chryl.util.UseCaseExcelUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Created by Chryl on 2019/10/22.
 */
@Controller
@RequestMapping("/ex")
public class ExcelController {


    //导入 导出
    @GetMapping("/cel")
    public void show(HttpServletRequest request, HttpServletResponse response) throws IOException {

        String fileName = "chrylcs";
        String columnNames[] = {"机构号", "营销代号", "营销名字"};// 列名
        String keys[] = {"orgid", "empcode", "empname"};// map中的key
        //数据信息list
        List<JxLoanPortfolio> list = new ArrayList<>();
        JxLoanPortfolio j1 = new JxLoanPortfolio();
        JxLoanPortfolio j2 = new JxLoanPortfolio();
        j1.setOrgid("123id");
        j1.setEmpcode("123code");
        j1.setEmpname("123name");
        j2.setOrgid("124id");
        j2.setEmpcode("124code");
        j2.setEmpname("124name");
        list.add(j1);
        list.add(j2);
        //list.map--为创建excel所用
        List<Map<String, Object>> lm = UseCaseExcelUtil.createExcelRecord(list);
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        ServletOutputStream out = response.getOutputStream();
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;
        try {
            //创建Workbook ,并写出
            UseCaseExcelUtil.createWorkBook(lm, keys, columnNames).write(os);
            // 设置response参数，可以打开下载页面
            response.reset();
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + new String((fileName + ".xls").getBytes(), "iso-8859-1"));

            byte[] content = os.toByteArray();
            InputStream is = new ByteArrayInputStream(content);
            out = response.getOutputStream();
            bis = new BufferedInputStream(is);
            bos = new BufferedOutputStream(out);
            byte[] buff = new byte[2048];
            int bytesRead;
            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                bos.write(buff, 0, bytesRead);
            }
        } catch (final Exception e) {
            e.printStackTrace();
        } finally {
            if (bis != null)
                bis.close();
            if (bos != null)
                bos.close();
        }
        //return null;
    }

}
