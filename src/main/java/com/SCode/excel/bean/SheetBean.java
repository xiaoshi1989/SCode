package com.SCode.excel.bean;

import java.util.List;
import java.util.Map;

public class SheetBean <T> {
    /**
     * 工作表名称
     */
    private String sheetName;
    /**
     * 工作表数据
     */
    private List<T> data;
    /**
     * 工作表头
     */
    private List<Map<String, String>> header;
    
    public String getSheetName() {
        return sheetName;
    }
    
    public List<T> getData() {
        return data;
    }
    
    public List<Map<String, String>> getHeader() {
        return header;
    }
    
    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }
    
    public void setData(List<T> data) {
        this.data = data;
    }
    
    public void setHeader(List<Map<String, String>> header) {
        this.header = header;
    }
    
}
