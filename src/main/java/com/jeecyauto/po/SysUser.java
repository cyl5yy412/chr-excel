package com.jeecyauto.po;

import lombok.Data;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;

import java.io.Serializable;
import java.util.Date;

/**
 * Created by Chr.yl on 2020/11/1.
 *
 * @author Chr.yl
 */
@Data
@ExcelTarget("SysUser")
public class SysUser implements Serializable {

    /**
     * id
     */
    private String id;

    /**
     * 登录账号
     */
    @Excel(name = "登录账号", width = 15)
    private String username;

    /**
     * 真实姓名
     */
    @Excel(name = "真实姓名", width = 15)
    private String realname;

    /**
     * 头像
     */
    @Excel(name = "头像", width = 15)
    private String avatar;

    /**
     * 生日
     */
    @Excel(name = "生日", width = 15, format = "yyyy-MM-dd")
    private Date birthday;

    /**
     * 性别（1：男 2：女）
     */
    @Excel(name = "性别", width = 15, dicCode = "sex")
    private Integer sex;

    /**
     * 电子邮件
     */
    @Excel(name = "电子邮件", width = 15)
    private String email;

    /**
     * 电话
     */
    @Excel(name = "电话", width = 15)
    private String phone;

    /**
     * 状态(1：正常  2：冻结 ）
     */
    @Excel(name = "状态", width = 15, replace = {"正常_1", "冻结_0"})
    private Integer status;

    public SysUser(String id, String username, String realname, String avatar, Date birthday,
                   Integer sex, String email, String phone, Integer status) {
        this.id = id;
        this.username = username;
        this.realname = realname;
        this.avatar = avatar;
        this.birthday = birthday;
        this.sex = sex;
        this.email = email;
        this.phone = phone;
        this.status = status;
    }
}
