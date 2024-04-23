package com.ls.easyexcel_extend.plugin.select;

import lombok.Data;
import lombok.EqualsAndHashCode;

import java.util.Map;

/**
 * 常规下拉选择列
 * @author ls
 * @version 1.0
 */
@EqualsAndHashCode(callSuper = true)
@Data
public final class CommonExcelSelectColumn extends BaseExcelSelectColumn<String[]> {

    /**
     * 判断source是否是空的
     *
     * @return 是否
     */
    @Override
    public boolean isSourceEmpty() {
        String[] source = super.getSource();
        return null == source || source.length == 0;
    }
}