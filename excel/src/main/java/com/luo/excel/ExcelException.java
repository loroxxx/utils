package com.luo.excel;

/**
 * 自定义excel解析异常
 * Created by Administrator on 2017/9/30 0030.
 */
public class ExcelException extends Exception{


    private static final long serialVersionUID = 1L;

    private String msg;
    private int code = 500;

    public ExcelException(String msg) {
        super(msg);
        this.msg = msg;
    }

    public ExcelException(String msg, Throwable e) {
        super(msg, e);
        this.msg = msg;
    }





}
