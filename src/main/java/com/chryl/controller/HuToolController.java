package com.chryl.controller;

import cn.hutool.core.io.IORuntimeException;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.baomidou.mybatisplus.core.metadata.OrderItem;
import com.chryl.po.LoanInfo;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * hutool工具类导出excel
 * Created by Chr.yl on 2020/12/21.
 *
 * @author Chr.yl
 */
@RestController
public class HuToolController {


    @GetMapping("huExport")
    public void show(HttpServletResponse response) {
        //数据信息list
        List<LoanInfo> list = new ArrayList<>();
        LoanInfo j1 = new LoanInfo();
        LoanInfo j2 = new LoanInfo();
        j1.setEmpId("123id");
        j1.setEmpCode("123code");
        j1.setEmpName("123name");
        j1.setStrDate("20201220");
        j1.setDate(new Date());

        j2.setEmpId("usgifakfal12321o");
        j2.setEmpCode("124code");
        j2.setEmpName("124name");
        j2.setStrDate("20201110");
        j2.setDate(new Date());
        list.add(j1);
        list.add(j2);

        //通过工具类创建writer
        ExcelWriter writer = ExcelUtil.getWriter();
        // 待发货
        String[] hearder = {"编号", "编码", "姓名", "开始时间", "时间"};
        Sheet sheet = writer.getSheet();
        sheet.setColumnWidth(0, 20 * 256);
        sheet.setColumnWidth(1, 20 * 256);
        sheet.setColumnWidth(3, 20 * 256);
        sheet.setColumnWidth(4, 60 * 256);
        sheet.setColumnWidth(5, 60 * 256);

        writer.merge(hearder.length - 1, "测试导出Excel");
        writer.writeRow(Arrays.asList(hearder));


        int row = 1;//行
        int col = 0;//列

        for (LoanInfo info : list) {

            row++;

            writer.writeCellValue(col++, row, info.getEmpId());
            writer.writeCellValue(col++, row, info.getEmpCode());
            writer.writeCellValue(col++, row, info.getEmpName());
            writer.writeCellValue(col++, row, info.getStrDate());
            writer.writeCellValue(col++, row, info.getDate());


            col = 0;
        }
        writeExcel(response, writer);


    }

    /**
     * 如果需要合并的话，就合并
     */
    private void mergeIfNeed(ExcelWriter writer, int firstRow, int lastRow, int firstColumn, int lastColumn, Object content) {
        if (lastRow - firstRow > 0 || lastColumn - firstColumn > 0) {
            writer.merge(firstRow, lastRow, firstColumn, lastColumn, content, false);
        } else {
            writer.writeCellValue(firstColumn, firstRow, content);
        }

    }

    //hutool导出excel
    private void writeExcel(HttpServletResponse response, ExcelWriter writer) {
        //response为HttpServletResponse对象
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        //test.xls是弹出下载对话框的文件名，不能为中文，中文请自行编码
        response.setHeader("Content-Disposition", "attachment;filename=1.xls");

        ServletOutputStream servletOutputStream = null;
        try {
            servletOutputStream = response.getOutputStream();
            writer.flush(servletOutputStream);
            servletOutputStream.flush();
        } catch (IORuntimeException | IOException e) {
            e.printStackTrace();
        } finally {
            writer.close();
            try {
                if (servletOutputStream != null) {
                    servletOutputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

}
