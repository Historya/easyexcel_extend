package com.ls.easyexcel_extend.base;

import java.lang.reflect.Field;

/**
 *
 * <br/>
 * date: 2024/3/28<br/>
 * @author ls
 * @version 1.0
 *
 * @author ls<br />
 */
public class ModelColumn {
    private Integer index;

    private Field field;

    public ModelColumn(){

    }

    public ModelColumn(Integer index, Field field) {
        this.index = index;
        this.field = field;
    }

    public Integer getIndex() {
        return index;
    }

    public void setIndex(Integer index) {
        this.index = index;
    }

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }
}
