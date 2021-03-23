package com.chryl.module.easyExcel.controller;

import com.chryl.module.easyExcel.service.AliLoanInfoService;
import com.chryl.po.AliLoanInfo;
import com.chryl.module.easyExcel.model.AjaxResult;
import com.chryl.module.easyExcel.poi.ExcelUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Chr.yl on 2021/3/23.
 *
 * @author Chr.yl
 */
@Controller
@RequestMapping("easy")
public class RuoyiController {

    @GetMapping("export")
    @ResponseBody
    public AjaxResult export(HttpServletResponse response) {

        List<AliLoanInfo> AliLoanInfoList = new ArrayList<>();
        AliLoanInfo j1 = new AliLoanInfo();
        AliLoanInfo j2 = new AliLoanInfo();
        j1.setEmpId("123id");
        j1.setEmpCode("123code");
        j1.setEmpName("123name");
        j2.setEmpCode("124code");
        j2.setEmpName("124name");
        AliLoanInfoList.add(j1);
        AliLoanInfoList.add(j2);


        ExcelUtil<AliLoanInfo> util = new ExcelUtil<>(AliLoanInfo.class);
        return util.exportEasyExcel(AliLoanInfoList, "测试");

    }

    @Autowired
    private AliLoanInfoService aliLoanInfoService;

    @PostMapping("/importData")
    @ResponseBody
    public AjaxResult importData(@RequestParam("uploadFileeasyExcel") MultipartFile uploadFileeasyExcel, boolean updateSupport) throws Exception {
        ExcelUtil<AliLoanInfo> util = new ExcelUtil<>(AliLoanInfo.class);
        List<AliLoanInfo> aliLoanInfos = util.importEasyExcel(uploadFileeasyExcel.getInputStream());
        String message = aliLoanInfoService.add(aliLoanInfos);

        return AjaxResult.success(message);
    }

}
