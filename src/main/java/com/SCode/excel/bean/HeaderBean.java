package com.SCode.excel.bean;

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;

public class HeaderBean {
    /**
     * 属性
     */
    private String field;
    /**
     * 标题
     */
    private String name;
    /**
     * 下拉框枚举
     */
    private Map<String,String> rowEnum;
    /**
     * excel 列数据类型   
     */
    private int type = Cell.CELL_TYPE_STRING;
    
    private int sort;
    
    
    
    public int getSort() {
        return sort;
    }

    
    public void setSort(int sort) {
        this.sort = sort;
    }

    public String getField() {
        return field;
    }
    
    public String getName() {
        return name;
    }
    
    public Map<String, String> getRowEnum() {
        return rowEnum;
    }
    
    public int getType() {
        return type;
    }
    
    public void setField(String field) {
        this.field = field;
    }
    
    public void setName(String name) {
        this.name = name;
    }
    
    public void setRowEnum(Map<String, String> rowEnum) {
        this.rowEnum = rowEnum;
    }
    
    public void setType(int type) {
        this.type = type;
    }
    
    
    
}
