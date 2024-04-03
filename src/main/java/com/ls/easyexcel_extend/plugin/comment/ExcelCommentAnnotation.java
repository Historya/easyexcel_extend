package com.ls.easyexcel_extend.plugin.comment;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;

/**
 * excel comment
 *
 * @author ls
 * @version 1.0
 * date: 2024-03-20
 */
@Target(FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCommentAnnotation {

    /**
     * 批注文本内容
     */
    String value() default "";

    /**
     * 批注行高, 一般不用设置
     * 这个参数可以设置不同字段 批注显示框的高度
     */
    int remarkRowHigh() default 2;

    /**
     * 批注列宽, 根据导出情况调整
     * 这个参数可以设置不同字段 批注显示框的宽度
     */
    int remarkColumnWide() default 4;
}

