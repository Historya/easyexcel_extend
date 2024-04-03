package com.ls.easyexcel_extend.plugin.select;

import lombok.Data;
import lombok.EqualsAndHashCode;

import java.util.Map;

/**
 * 联级
 */
@EqualsAndHashCode(callSuper = true)
@Data
public final class CascadeExcelSelectColumn extends ExcelSelectColumn {
    /**
     * 下拉内容
     */
    private Map<String, String[]> source;
}