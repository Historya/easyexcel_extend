package com.ls.easyexcel_extend.exception;

import com.alibaba.excel.exception.ExcelAnalysisException;

import java.text.MessageFormat;
import java.util.Objects;

/**
 * excel 处理异常
 * <br/>
 * date: 2024/4/7<br/>
 * version 1.0
 *
 * @author ls<br />
 */
public class ExcelHandlerException extends ExcelAnalysisException {

    private Integer rowIndex;

    private Integer columnIndex;

    private String messageInfo;

    public ExcelHandlerException(Integer rowIndex, Integer columnIndex, String messageInfo, Throwable cause){
        super(cause);
        this.rowIndex = rowIndex;
        this.columnIndex = columnIndex;
        this.messageInfo = messageInfo;
    }

    public ExcelHandlerException(String message, Throwable cause){
        super(message,cause);
    }

    /**
     * 生成标准信息描述
     * @return 信息描述
     */
    public String generationMessage(){
        if(Objects.nonNull(this.rowIndex) && Objects.nonNull(this.messageInfo)){
            return MessageFormat.format("处理Excel 文件错误，请检查Excel 第{0}行,{1}",this.rowIndex,this.messageInfo);

        }
        if(Objects.nonNull(this.rowIndex)){
            return MessageFormat.format("处理Excel 文件错误，请检查Excel 第{0}行",this.rowIndex);
        }
        return this.getMessage();
    }
}
