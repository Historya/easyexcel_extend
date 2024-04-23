package com.ls.easyexcel_extend.plugin.select;

import lombok.Data;
import lombok.EqualsAndHashCode;

import java.util.Map;

/**
 * 联级下拉选择列
 * @author ls
 * @version 1.0
 */
@EqualsAndHashCode(callSuper = true)
@Data
public final class CascadeExcelSelectColumn extends BaseExcelSelectColumn<Map<String, String[]>> {
    /**
     * 判断source是否是空的
     *
     * @return 是否
     */
    @Override
    public boolean isSourceEmpty() {
        Map<String, String[]> source = super.getSource();
        return null == source || source.isEmpty();
    }

}