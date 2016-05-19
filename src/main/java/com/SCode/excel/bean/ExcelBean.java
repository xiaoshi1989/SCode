package com.SCode.excel.bean;

import java.util.List;
/**
 * 
 * EXCEL类
 * @author  shizj
 * @version  [1.0, 2016年5月17日]
 */
public class ExcelBean <T> {

    public static final String EXCEL_VERSION_2003 = "2003";
    
    public static final String EXCEL_VERSION_2007 = "2007";
    
    /**
     * excel 最大行数
     */
    public static final int   EXCEL_MAX_CLOUMN_2007 = 1048576;
    /**
     * 下拉列表最大值
     */
    public static final int   EXCEL_MAX_ROWENUM = 10;
    
    /**
     * 文件名称
     */
    private String fileName;
    /**
     * 工作表
     */
    private List<SheetBean<T>> sheets;
    /**
     * 文件保存目录
     */
    private String filePath;
    /**
     * Excel版本
     * 2003/2007
     */
    private String version = EXCEL_VERSION_2007 ;
    
    public String getFileName() {
        return fileName;
    }
    
    public List<SheetBean<T>> getSheets() {
        return sheets;
    }
    
    public String getFilePath() {
        return filePath;
    }
    
    public String getVersion() {
        return version;
    }
    
    public void setFileName(String fileName) {
        this.fileName = fileName;
    }
    
    public void setSheets(List<SheetBean<T>> sheets) {
        this.sheets = sheets;
    }
    
    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }
    
    public void setVersion(String version) {
        this.version = version;
    }
    
   
}
