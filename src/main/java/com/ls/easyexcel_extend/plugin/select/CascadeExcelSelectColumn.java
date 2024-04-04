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
public final class CascadeExcelSelectColumn extends BaseExcelSelectColumn {
    /**
     * 下拉内容
     */
    private Map<String, String[]> source;
}