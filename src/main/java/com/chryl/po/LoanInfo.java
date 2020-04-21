package com.chryl.po;

import java.io.Serializable;
import java.util.Date;

public class LoanInfo implements Serializable {

    private static final long serialVersionUID = 374037782462545488L;
    private String id;

    private String empcode;

    private String empname;

    private String strDate;

    private Date date;

    public LoanInfo() {
    }

    public LoanInfo(String id, String empcode, String empname, String strDate, Date date) {
        this.id = id;
        this.empcode = empcode;
        this.empname = empname;
        this.strDate = strDate;
        this.date = date;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getEmpcode() {
        return empcode;
    }

    public void setEmpcode(String empcode) {
        this.empcode = empcode;
    }

    public String getEmpname() {
        return empname;
    }

    public void setEmpname(String empname) {
        this.empname = empname;
    }

    public String getStrDate() {
        return strDate;
    }

    public void setStrDate(String strDate) {
        this.strDate = strDate;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }
}