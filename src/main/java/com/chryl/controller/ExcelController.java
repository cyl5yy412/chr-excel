package com.chryl.controller;

import com.chryl.po.LoanInfo;
import com.chryl.util.ExcelUtil;
import com.chryl.util.UseCaseExcelUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Chryl on 2019/10/22.
 */
@Controller
@RequestMapping("/excel")
public class ExcelController {


    //导出
    @GetMapping("/export")
    public void show(HttpServletRequest request, HttpServletResponse response) throws IOException {

        String fileName = "chrylcs";
        String columnNames[] = {"id", "营销代号", "营销名字", "日期1", "日期2"};// 列名
        String keys[] = {"id", "empcode", "empname", "strDate", "date"};// map中的key
        //数据信息list
        List<LoanInfo> list = new ArrayList<>();
        LoanInfo j1 = new LoanInfo();
        LoanInfo j2 = new LoanInfo();
        j1.setId("123id");
        j1.setEmpcode("123code");
        j1.setEmpname("123name");
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

    //模板文件导入
    @RequestMapping(value = "/import_excel", method = RequestMethod.POST)
    public Object import_excel(@RequestParam("uploadFile") MultipartFile uploadFile, String date) {
        String fileName = uploadFile.getOriginalFilename();
        String fileType = fileName.substring(fileName.lastIndexOf(".") + 1);
        try {
            System.out.println(fileName + "." + fileType);
            InputStream is = uploadFile.getInputStream();
            if (fileType.equals("xlsx") || fileType.equals("xls")) {

//                String keys[] = {"orgname", "frorgname", "brhid", "businame", "businum", "acctno", "tradeCount", "tradeBal", "tradeCharge", "balamt"};//map中的key
                String keys[] = {"orgid", "empcode", "empname"};
                //属性对应的列
                Map<String, Integer> columnIndexMap = new HashMap<String, Integer>();
                for (int i = 0; i < keys.length; i++) {
                    columnIndexMap.put(keys[i], i);
                }
                List<LoanInfo> list = ExcelUtil.excle2Object(LoanInfo.class, is, fileType, columnIndexMap);
                if (list.size() > 0) {
                    for (LoanInfo jxBusitraCount : list) {
                        try {
//                            jxBusitraCountService.insert(jxBusitraCount, true);
                        } catch (Exception e) {
//                            return ResponseResult.build(500, jxBusitraCount.getBusiname()+"交易信息已导入，请检查后重新导入！");
                        }
                    }
//                    result = ResponseResult.ok();
                } else {
//                    return ResponseResult.build(ResponseResult.ERROR, "文件内容为空！");
                }
            } else {
                //文件类型错误，不是xlsx或者xls格式
//                result = ResponseResult.build(ResponseResult.ERROR, "所选文件不是.xlsx或者.xls格式文件！");
            }
            is.close();
        } catch (Exception e) {
//            logger.error("上传异常",e);
//            result = ResponseResult.build(ResponseResult.ERROR, ExceptionUtil.getStackTrace(e));
        }
//        return result;
        return "";
    }

}
