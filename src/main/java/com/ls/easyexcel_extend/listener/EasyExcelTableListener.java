package com.ls.easyexcel_extend.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellExtra;
import com.alibaba.excel.metadata.data.ReadCellData;

import javax.annotation.Nonnull;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Consumer;

/**
 * 简单的表格分析监听器
 * <br/>
 * date: 2024/4/10<br/>
 * version 0.1
 *
 * @author ls<br />
 */
public class EasyExcelTableListener<T> extends ControlledAnalysisEventListener<T> {

    private final BiConsumer<T, AnalysisContext> invokeFun;

    private final Consumer<AnalysisContext> afterAllAnalysedFun;

    private BiConsumer<CellExtra, AnalysisContext> extraFun = (CellExtra cellExtra, AnalysisContext analysisContext) -> {
    };

    private BiConsumer<Map<Integer, ReadCellData<?>>, AnalysisContext> invokeHeadFun = (Map<Integer, ReadCellData<?>> headMap, AnalysisContext analysisContext) -> {
    };

    public EasyExcelTableListener(@Nonnull BiConsumer<T, AnalysisContext> readInvokeFun,@Nonnull Consumer<AnalysisContext> redAfterFun) {
        this.invokeFun = readInvokeFun;
        this.afterAllAnalysedFun = redAfterFun;
    }

    public EasyExcelTableListener(@Nonnull BiConsumer<T, AnalysisContext> readInvokeFun,@Nonnull Consumer<AnalysisContext> redAfterFun,@Nonnull BiConsumer<CellExtra, AnalysisContext> extraFun) {
        this.invokeFun = readInvokeFun;
        this.afterAllAnalysedFun = redAfterFun;
        this.extraFun = extraFun;
    }

    public EasyExcelTableListener(@Nonnull BiConsumer<T, AnalysisContext> readInvokeFun,@Nonnull  Consumer<AnalysisContext> redAfterFun,@Nonnull BiConsumer<CellExtra, AnalysisContext> extraFun,@Nonnull BiConsumer<Map<Integer, ReadCellData<?>>,AnalysisContext> invokeHeadFun) {
        this.invokeFun = readInvokeFun;
        this.afterAllAnalysedFun = redAfterFun;
        this.extraFun = extraFun;
        this.invokeHeadFun = invokeHeadFun;
    }

    /**
     * The current method is called when extra information is returned
     *
     * @param extra   extra information
     * @param context analysis context
     */
    @Override
    public void extra(CellExtra extra, AnalysisContext context) {
        this.extraFun.accept(extra, context);
    }

    @Override
    public void invokeHead(Map<Integer, ReadCellData<?>> headMap, AnalysisContext context) {
        this.invokeHeadFun.accept(headMap, context);
    }

    /**
     * When analysis one row trigger invoke function.
     *
     * @param data    one row value. It is same as {@link AnalysisContext#readRowHolder()}
     * @param context analysis context
     */
    @Override
    public void invoke(T data, AnalysisContext context) {
        this.invokeFun.accept(data, context);
    }

    /**
     * if have something to do after all analysis
     *
     * @param context
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        this.afterAllAnalysedFun.accept(context);
    }
    
}
