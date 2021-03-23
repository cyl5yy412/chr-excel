//package com.chryl.module.jeecyauto.po;
//
//import org.jeecgframework.poi.excel.annotation.*;
//
///**
// * Created by Chr.yl on 2020/11/1.
// *
// * @author Chr.yl
// */
//@ExcelTarget("courseEntity")
//public class CourseEntity implements java.io.Serializable {
//    /**
//     * 主键
//     */
//    private String id;
//    /**
//     * 课程名称
//     */
//    @Excel(name = "课程名称", orderNum = "1", needMerge = true)
//    private String name;
//    /**
//     * 老师主键
//     */
//    @ExcelEntity(id = "yuwen")
//    @ExcelVerify()
//    private TeacherEntity teacher;
//    /**
//     * 老师主键
//     */
//    @ExcelEntity(id = "shuxue")
//    private TeacherEntity shuxueteacher;
//
//    @ExcelCollection(name = "选课学生", orderNum = "4")
//    private List<StudentEntity> students;
//
//
//
//}
