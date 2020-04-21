package com.chryl.controller;

import com.chryl.po.LoanInfo;
import com.chryl.service.LoanInfoService;
import com.chryl.util.ExcelUtil;
import com.chryl.util.UseCaseExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.ibatis.reflection.ExceptionUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.*;

/**
 * Created by Chryl on 2019/10/22.
 */
@RestController
@Slf4j
@RequestMapping("/excel")
public class ExcelController {

    //导出模板
    @GetMapping(value = "/template")
    public void template(HttpServletRequest request, HttpServletResponse response, LoanInfo loanInfo) throws Exception {
        String fileName = "测试模板";

        String columnNames[] = {"行销号码", "营销姓名", "时间1", "时间2"};//列名

        ByteArrayOutputStream os = new ByteArrayOutputStream();
        ServletOutputStream out = null;
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;

        try {
            UseCaseExcelUtil.createWorkBook(null, null, columnNames).write(os);
            // 设置response参数，可以打开下载页面
            response.reset();
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + new String((fileName + ".xls").getBytes(), "iso-8859-1"));

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
    }


    //导出
    @GetMapping("/export")
    public void export(HttpServletRequest request, HttpServletResponse response) throws IOException {

        String fileName = "测试导出";
        String columnNames[] = {"营销代号", "营销名字", "日期1", "日期2"};// 列名
        String keys[] = {"empcode", "empname", "strDate", "date"};// map中的key
        //数据信息list
        List<LoanInfo> list = new ArrayList<>();
        LoanInfo j1 = new LoanInfo();
        LoanInfo j2 = new LoanInfo();
        j1.setEmpId("123id");
        j1.setEmpCode("123code");
        j1.setEmpName("123name");
        j2.setEmpCode("124code");
        j2.setEmpName("124name");
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
    }

    @Autowired
    private LoanInfoService loanInfoService;

    //模板文件导入
    @PostMapping(value = "/import")
    public Object importExc(@RequestParam("uploadFile") MultipartFile uploadFile) {
        String fileName = uploadFile.getOriginalFilename();
        String fileType = fileName.substring(fileName.lastIndexOf(".") + 1);
        String result = "";
        try {
            System.out.println(fileName + "." + fileType);
            InputStream is = uploadFile.getInputStream();
            if (fileType.equals("xlsx") || fileType.equals("xls")) {
                String keys[] = {"empCode", "empName", "strDate", "date"};
                //属性对应的列
                Map<String, Integer> columnIndexMap = new HashMap<String, Integer>();
                for (int i = 0; i < keys.length; i++) {
                    columnIndexMap.put(keys[i], i);
                }
                List<LoanInfo> list = ExcelUtil.excle2Object(LoanInfo.class, is, fileType, columnIndexMap);
                if (list.size() > 0) {
                    for (LoanInfo loanInfo : list) {
                        try {
                            loanInfo.setEmpId(UUID.randomUUID().toString().replaceAll("-", ""));
                            loanInfoService.add(loanInfo);
                        } catch (Exception e) {
                            return "导入失败";
                        }
                    }
                    result = "success";
                } else {
                    return "文件内容为空！";
                }
            } else {
                //文件类型错误，不是xlsx或者xls格式
                result = "所选文件不是.xlsx或者.xls格式文件！";
            }
            is.close();
        } catch (Exception e) {
            log.error("上传异常", e);
            result = e.getMessage();
        }
        return result;
    }

}
