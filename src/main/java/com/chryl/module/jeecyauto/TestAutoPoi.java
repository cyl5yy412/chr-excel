package com.chryl.module.jeecyauto;

import com.chryl.module.jeecyauto.po.SysUser;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Created by Chr.yl on 2020/11/1.
 *
 * @author Chr.yl
 */
public class TestAutoPoi {
    public static void main(String[] args) {

        List<SysUser> list = new ArrayList<>();
        SysUser sysUser = new SysUser("123", "123", "123", "123"
                , new Date(), new Integer(2), "123", "123", new Integer(2));
        list.add(sysUser);

        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams(
                "2412312", "测试", "测试"), SysUser.class, list);
    }
}
