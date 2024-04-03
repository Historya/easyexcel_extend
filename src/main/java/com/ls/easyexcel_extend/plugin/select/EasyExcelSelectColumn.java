package com.ls.easyexcel_extend.plugin.select;

import lombok.Data;
import lombok.EqualsAndHashCode;

@EqualsAndHashCode(callSuper = true)
@Data
public final class EasyExcelSelectColumn extends ExcelSelectColumn {
    /**
     * 下拉内容
     */
    private String[] source;
}