package com.ls.easyexcel_extend.base;

import com.alibaba.excel.annotation.ExcelIgnore;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * 基础处理器
 * <br/>
 * date: 2024/3/28<br/>
 *
 * @author ls
 * @version 1.0
 */
public interface BaseHandler<E extends Model>{

    /**
     * 模型
     */
    Class<E> getModelClass();

    default List<ModelField> getModelFields(){
        Field[] declaredFields = this.getModelClass().getDeclaredFields();
        List<ModelField> fieldList = new ArrayList<>();
        for (int i = 0; i < declaredFields.length; i++) {
            Field itemField = declaredFields[i];
            if(itemField.isAnnotationPresent(ExcelIgnore.class)){
                continue;
            }

            fieldList.add(new ModelField(i,itemField));
        }

        return fieldList;
    }
}
