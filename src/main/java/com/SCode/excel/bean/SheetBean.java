package com.SCode.excel.bean;

import java.util.List;
/**
 * 
 * 工作表
 * @author  shizj
 * @version  [1.0, 2016年5月17日]
 */
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
    private List<HeaderBean> header;
    
    public String getSheetName() {
        return sheetName;
    }
    
    public List<T> getData() {
        return data;
    }
    
    public List<HeaderBean> getHeader() {
        return header;
    }
    
    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }
    
    public void setData(List<T> data) {
        this.data = data;
    }
    
    public void setHeader(List<HeaderBean> header) {
        this.header = header;
    }

}
