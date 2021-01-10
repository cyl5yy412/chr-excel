package com.chryl.controller;

import com.alibaba.fastjson.JSON;
import com.chryl.po.LoanInfo;
import com.eladmin.util.FileUtil;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
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
}
