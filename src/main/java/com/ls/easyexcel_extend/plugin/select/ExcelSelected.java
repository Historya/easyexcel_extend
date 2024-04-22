package com.ls.easyexcel_extend.plugin.select;

import java.lang.annotation.*;

/**
 * excel下拉选择
 * @author ls
 * @version 1.0
 */
@Documented
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSelected {
    /**
     * 类型
     */
    Type type();
    /**
     * 固定下拉内容
     */
    String[] source() default {};
    /**
     * 动态数据类
     * 动态下拉内容
     */
    Class<? extends ExcelDynamicDataSource> sourceHandle() default DefaultDataSource.class;

    /**
     * 动态数据 参数
     */
    String[] sourceParams() default {};

    /**
     *父索引
     */
    int parentColumnIndex() default -1;

    /**
     * 设置下拉框的起始行，默认为第二行
     */
    int firstRow() default 1;
 
    /**
     * 设置下拉框的结束行，默认为最后一行
     * 65536
     */
    int lastRow() default 65536;

    enum Type{
        //序列
        SEQUENCE,
        //自定义
        CUSTOMER,
        //联级
        CASCADE
    }

}