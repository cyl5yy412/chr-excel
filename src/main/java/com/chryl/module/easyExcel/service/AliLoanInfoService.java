package com.chryl.module.easyExcel.service;

import com.chryl.module.easyExcel.mapper.AliLoaniInfoMapper;
import com.chryl.po.AliLoanInfo;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.UUID;

/**
 * Created by Chr.yl on 2021/3/23.
 *
 * @author Chr.yl
 */
@Service
public class AliLoanInfoService {
    @Autowired
    private AliLoaniInfoMapper aliLoaniInfoMapper;

    public String add(List<AliLoanInfo> aliLoanInfoList) {
        int successNum = 0;
        for (AliLoanInfo aliLoanInfo : aliLoanInfoList) {
            aliLoanInfo.setEmpId(UUID.randomUUID().toString().replaceAll("-", "").substring(15));
            successNum = aliLoaniInfoMapper.insert(aliLoanInfo);
            successNum++;
        }

        StringBuilder successMsg = new StringBuilder();
        StringBuilder failureMsg = new StringBuilder();

        if (successNum > 0) {
            successMsg.insert(0, "恭喜您，数据已全部导入成功！共 " + successNum + " 条，数据如下：");
            return successMsg.toString();
        } else {
            failureMsg.insert(0, "很抱歉，导入失败！数据格式不正确，错误如下：");
            return failureMsg.toString();
        }
    }
}
