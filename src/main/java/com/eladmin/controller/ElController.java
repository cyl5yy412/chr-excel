package com.eladmin.controller;

import com.alibaba.fastjson.JSON;
import com.chryl.po.LoanInfo;
import com.eladmin.util.FileUtil;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * eladmin导出
 * Created by Chr.yl on 2021/1/10.
 *
 * @author Chr.yl
 */
@RestController
@RequestMapping("/eladmin")
public class ElController {

    @GetMapping("/excel")
    public void show(HttpServletResponse response) throws IOException {
        LoanInfo j1 = new LoanInfo();
        LoanInfo j2 = new LoanInfo();
        j1.setEmpId("123id");
        j1.setEmpCode("123code");
        j1.setEmpName("123name");
        j2.setEmpCode("124code");
        j2.setEmpName("124name");

        List<Map<String, Object>> objList = new ArrayList<>();

        Map<String, Object> map1 = JSON.parseObject(JSON.toJSONString(j1), Map.class);
        Map<String, Object> map2 = JSON.parseObject(JSON.toJSONString(j2), Map.class);
        objList.add(map1);
        objList.add(map2);
        FileUtil.downloadExcel(objList, response);


    }


    /**
     * eladmin 导出excel : list<Obj> 直接导出
     *
     * @param response
     * @throws IOException
     */
    @GetMapping("/excel222")
    public void show222(HttpServletResponse response) throws IOException {

        List<Map<String, Object>> objList = new ArrayList<>();//导出的list map

        List<LoanInfo> loanInfoList = new ArrayList<>();
        LoanInfo j1 = new LoanInfo();
        LoanInfo j2 = new LoanInfo();
        j1.setEmpId("123id");
        j1.setEmpCode("123code");
        j1.setEmpName("123name");
        j2.setEmpCode("124code");
        j2.setEmpName("124name");
        loanInfoList.add(j1);
        loanInfoList.add(j2);


        for (LoanInfo loanInfo : loanInfoList) {
            Map<String, Object> map = new LinkedHashMap<>();
            map.put("职工ID编号", loanInfo.getEmpId());
            map.put("职工名", loanInfo.getEmpName());
            map.put("职工编号", loanInfo.getEmpCode());
            map.put("STR日期", loanInfo.getStrDate());
            map.put("日期", loanInfo.getDate());
            objList.add(map);
        }
        FileUtil.downloadExcel(objList, response);


    }
}
