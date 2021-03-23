package com.chryl.po;

import com.alibaba.excel.annotation.ExcelIgnoreUnannotated;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;
import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
@ExcelIgnoreUnannotated
@ColumnWidth(16)
@HeadRowHeight(14)
@HeadFontStyle(fontHeightInPoints = 11)
public class AliLoanInfo implements Serializable {

    private static final long serialVersionUID = 374037782462545488L;

    @ExcelProperty(value = "主键")
    private String empId;

    @ExcelProperty(value = "编号")
    private String empCode;

    @ExcelProperty(value = "姓名")
    private String empName;

    @ExcelProperty(value = "r日期")
    private String strDate;

    @ExcelProperty(value = "日期")
    private Date date;


}