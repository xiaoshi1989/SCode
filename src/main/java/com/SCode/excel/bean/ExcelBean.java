package com.SCode.excel.bean;

import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
/**
 * 
 * excel bean
 * @author  shizj
 * @version  [1.0, 2016年5月10日]
 * @since  [ERP/模块版本]
 */
public class ExcelBean <T> {

    public static final String EXCEL_VERSION_2003 = "2003";
    
    public static final String EXCEL_VERSION_2007 = "2007";
    
    /**
     * 文件名称
     */
    private String fileName;
    /**
     * 工作表
     */
    private List<Sheet> sheets;
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
    
    public List<Sheet> getSheets() {
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
    
    public void setSheets(List<Sheet> sheets) {
        this.sheets = sheets;
    }
    
    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }
    
    public void setVersion(String version) {
        this.version = version;
    }
    
}
