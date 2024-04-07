package com.ls.easyexcel_extend.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.ls.easyexcel_extend.exception.ExcelHandlerException;

/**
 * 可控的解析监听器
 * <br/>
 * date: 2024/4/7<br/>
 * version 1.0
 *
 * @author ls<br />
 */
public abstract class ControlledAnalysisEventListener<T> extends AnalysisEventListener<T> {
    /**
     * All listeners receive this method when any one Listener does an error report. If an exception is thrown here, the
     * entire read will terminate.
     *
     * @param exception exception
     * @param context  上下文
     * @see AnalysisContext
     */
    @Override
     public final void onException(Exception exception, AnalysisContext context) throws ExcelHandlerException {
        this.beforeException();
        throw new ExcelHandlerException(exception.getMessage(),exception);
    }

    /**
     * 异常之前执行
     */
    public void beforeException(){}
}
