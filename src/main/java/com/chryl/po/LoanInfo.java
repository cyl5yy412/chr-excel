package com.chryl.po;

import lombok.*;

import java.io.Serializable;
import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class LoanInfo implements Serializable {

    private static final long serialVersionUID = 374037782462545488L;
    private String empId;

    private String empCode;

    private String empName;

    private String strDate;

    private Date date;


}