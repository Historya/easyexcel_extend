package com.ls.easyexcel_extend.base;

import lombok.Data;

import java.lang.reflect.Field;

/**
 * excel table head model
 * <br/>
 * date: 2024/3/28<br/>
 * @author ls
 * @version 1.0
 *
 * @author ls<br />
 */
@Data
public class ModelField {
    /**
     * 列索引
     */
    private Integer index;

    /**
     * model entity field
     */
    private Field field;

    public ModelField(){

    }

    public ModelField(Integer index, Field field) {
        this.index = index;
        this.field = field;
    }
}
