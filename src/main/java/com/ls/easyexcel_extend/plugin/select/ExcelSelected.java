package com.ls.easyexcel_extend.plugin.select;

import java.lang.annotation.*;

/**
 * excel下拉选择
 * @author ls
 * @version 1.0
 */
@Documented
@Target({ElementType.FIELD})//用此注解用在属性上。
@Retention(RetentionPolicy.RUNTIME)//注解不仅被保存到class文件中，jvm加载class文件之后，仍然存在；
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
    Class<? extends ExcelDynamicDataSource> sourceHandle();

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

    static enum Type{
        SEQUENCE,
        CUSTOMER
    }

}