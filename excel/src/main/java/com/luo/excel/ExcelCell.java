package com.luo.excel;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 配合Excel2List使用
 * index用于匹配每列的数据，从0开始
 * allowNull=true 表示该列可以为空
 * 如果属性的类型是Data，会通过SimpleDateFormat自动解析成Date
 * 如果属性的类型是数字，或通过DecimalFormat自动解析成不带小数的格式
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelCell {

    public int index();

    public String name();

    public boolean allowNull() default false;

    public String format() default "yyyyMMdd";

    public String decimal() default "0";


}
