package com.luo.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

/**
 * excel文件解析类，把文件解析成List<Bean>形式.
 * 配合@ExcelCell注解使用。
 * <p>
 * Created by luoyuanq on 2017/9/25 0025.
 */
public class Excel2List<T> {

    private Class resolveTemplate;

    private int startPage;

    private int starRow;

    private Class<T> clazz;


    private HashMap<String, templateField> templateFieldMap = new HashMap<>();//按index保存每个field的注解信息

    private static Logger logger = LoggerFactory.getLogger(Excel2List.class);


    /**
     * * @param starPage-从第几页解析
     *
     * @param starRow-从第几行解析
     * @param resolveTemplate 解析目标类
     */
    public Excel2List(Class resolveTemplate, int starPage, int starRow) {
        this.startPage = starPage;
        this.starRow = starRow;
        this.resolveTemplate = resolveTemplate;
        this.clazz = resolveTemplate;

    }


    /**
     * 把excel文件转为List<Bean>
     *
     * @param inputStream
     * @return
     */
    public List<T> resolve2List(InputStream inputStream) throws  ExcelException {

        resolveAnnotation();

        ArrayList list = null;
        try {
            Workbook book = WorkbookFactory.create(inputStream);
            Sheet sheet = book.getSheetAt(startPage);
            int rowNum = sheet.getLastRowNum();// 获取行数
            if (rowNum < 1) {
                return new ArrayList<T>();
            }

            list = new ArrayList();
            for (int i = starRow; i <= rowNum; i++) {
                Row row = sheet.getRow(i);

                if (checkRowIsNull(row)) {
                    continue;
                }

                T vo = createVo(row);
                list.add(vo);
            }

        } catch (InvalidFormatException |IOException e) {
            logger.error("文件读取失败或者文件格式错误",e);
            throw new ExcelException("文件读取失败或者文件格式错误",e);
        }

        if(logger.isDebugEnabled()){
            logger.debug("解析出来的list为:"+list.toString());
        }
        return list;


    }


    /**
     * 传入Row，检查该行是否空行
     *
     * @param row
     * @return
     */
    private boolean checkRowIsNull(Row row) {
        if (row == null) {//有空行会跳过
            return true;
        }

        boolean isBlank = true;
        for (Cell cell : row) {
            int index = cell.getColumnIndex();
            templateField templateField = templateFieldMap.get(String.valueOf(index));
            if (templateField == null) {
                continue;
            }

            String cellValue = getCellValue(cell, templateField.simpleDateFormat,templateField.decimal);
            if (cellValue != null && !cellValue.equals("")) {
                isBlank = false;
            }
        }

        return isBlank;
    }


    private T createVo(Row row) throws ExcelException {


        T instance = null;
        try {
            instance = clazz.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            throw new RuntimeException("Bean实例化失败:" + e);
        }

        for (Cell cell : row) {
            int index = cell.getColumnIndex();
            templateField templateField = templateFieldMap.get(String.valueOf(index));
            if (templateField == null) {
                continue;
            }


            String cellValue = getCellValue(cell, templateField.simpleDateFormat,templateField.decimal);

            if (isBlank(cellValue) && templateField.allowNull == false) {
                throw new ExcelException("第"+(row.getRowNum()+1)+"行的"  +templateField.name + "不能为空");
            }


            Method method = null;//构造setXXX方法
            try {
                method = clazz.getMethod(templateField.setMethod, templateField.returnType);
            } catch (NoSuchMethodException e) {
                throw new RuntimeException("找不到" + templateField.fieldName + "的 getXXX,setXXX方法");
            }

            //赋值
            try {
                loadValue(method, instance, cellValue, templateField.returnType, templateField.simpleDateFormat);
            } catch (Exception e) {
                logger.error("解析"+templateField.name+"失败，请检查该列数据格式是否都为字符串或者数字",e);
                throw new ExcelException("解析"+templateField.name+"失败，请检查该列数据格式是否都为字符串或者数字",e);
            }

        }

        return instance;

    }


    /**
     * 注入属性值,先构造set与get方法，再通过反射调用
     */
    private <T> void loadValue(Method method, T instance, String value, Class type, SimpleDateFormat simpleDateFormat) throws InvocationTargetException, IllegalAccessException, ParseException {


            if (type == String.class) {
                method.invoke(instance, value);
            } else if (type == int.class || type == Integer.class) {
                method.invoke(instance, Integer.parseInt(value));

            } else if (type == long.class || type == Long.class) {
                method.invoke(instance, Long.parseLong(value));


            } else if (type == float.class || type == Float.class) {
                method.invoke(instance, Float.parseFloat(value));

            } else if (type == double.class || type == Double.class) {
                method.invoke(instance, Double.parseDouble(value));

            } else if (type == Date.class) {
                if (!value.equals(""))
                    method.invoke(instance, simpleDateFormat.parse(value));
            } else if (type == BigDecimal.class) {
                if (!value.equals(""))
                    method.invoke(instance, new BigDecimal(value));
            } else if (type == Object.class) {
                method.invoke(instance, value);
            }

    }

    /**
     * 传入field的名字，构造setXXX的方法
     *
     * @param field
     * @return
     */
    private String initSetMethod(String field) {
        return "set" + field.substring(0, 1).toUpperCase() + field.substring(1);
    }

    /**
     * 传入field的名字，构造getXXX的方法
     *
     * @param field
     * @return
     */
    private String initGetMethod(String field) {
        return "get" + field.substring(0, 1).toUpperCase() + field.substring(1);
    }


    /**
     * 把cell的值转换为String
     */
    private String getCellValue(Cell cell, SimpleDateFormat format, DecimalFormat decimalFormat) {
        Object value = null;
        if (null == cell) {
            return "";
        }
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = format.format(cell.getDateCellValue());
                } else {
                    value =  decimalFormat.format(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_FORMULA:
                value = cell.getCellFormula();
                break;
            default:
                value = cell.getStringCellValue();
                break;
        }
        return String.valueOf(value);
    }


    /**
     *
     */
    private void resolveAnnotation() {

        //根据注解解析模板类
        Field[] fieldsArr = resolveTemplate.getDeclaredFields();
        for (Field field : fieldsArr) {
            ExcelCell ec = field.getAnnotation(ExcelCell.class);
            if (ec == null) {
                // 没有ExcelCell Annotation 视为不汇入
                continue;
            }
            templateField template = new templateField();
            template.name = ec.name().trim();
            template.allowNull = ec.allowNull();
            template.fieldName = field.getName();
            template.simpleDateFormat = new SimpleDateFormat(ec.format());
            template.decimal = new DecimalFormat(ec.decimal());
            template.returnType = field.getType();
            template.setMethod = initSetMethod(field.getName());
            templateFieldMap.put(String.valueOf(ec.index()), template);//按index保存每个field的信息

        }

    }

    private boolean isBlank(String str) {
        int strLen;
        if (str != null && (strLen = str.length()) != 0) {
            for (int i = 0; i < strLen; ++i) {
                if (!Character.isWhitespace(str.charAt(i))) {
                    return false;
                }
            }

            return true;
        } else {
            return true;
        }
    }


    private class templateField {

        private String name;
        private boolean allowNull;
        private String fieldName;
        private SimpleDateFormat simpleDateFormat;
        private DecimalFormat decimal;
        private String setMethod;
        private Class returnType;


    }


}
