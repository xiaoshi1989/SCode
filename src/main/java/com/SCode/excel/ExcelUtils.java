package com.SCode.excel;

import java.beans.BeanInfo;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.OpenXML4JRuntimeException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.CollectionUtils;

import com.SCode.excel.bean.ExcelBean;
import com.SCode.excel.bean.HeaderBean;
import com.SCode.excel.bean.SheetBean;
import com.SCode.excel.exception.ExcelException;
import com.SCode.excel.util.StringUtil;
/**
 * 
 * excel 相关操作
 * @since POI 3.13
 * 
 * @author  shizj
 * @version  [1.0, 2016年5月10日]
 * @since  [ERP/模块版本]
 */
public class ExcelUtils{
	
	public final static String SUCCESS ="Excel导出成功";
	
	public final static String EXCEL2003 = "2003";
	
	public final static String EXCEL2007 = "2007";
	
	public final static String DEFAULT_SHEET_NAME = "SHEET";
	
	
	/**
	 * 
	 * 导出全部属性
	 * @param data 数据
	 * @return
	 * @throws ExcelException
	 * @see [类、类#方法、类#成员]
	 */
	public  static <T> Workbook creatWorkbook(List<T> data){
	    if(data==null){
	        return null;
	    }
	    return creatWorkbook(data, null, null, null, null);
	}
	/**
	 * 
	 * 导出全部属性或者按注解导出
	 * 在T的成员表里上增加@ExportAnnotation(name="名称",type=0,sort=1)
	 * <功能详细描述>
	 * @param data 数据
	 * @param userAnnotation 是否使用注解导出
	 * @return
	 * @see [类、类#方法、类#成员]
	 */
	public  static <T> Workbook creatWorkbook(List<T> data,boolean userAnnotation){
       if(userAnnotation){
           return  creatWorkbook(data, null, null);
       }else{
           return creatWorkbook(data);
       }
    }
	/**
	 * 根据数据类上的注解获取默认表头
	 * <一句话功能简述>
	 * <功能详细描述>
	 * @param data
	 * @param sheetName
	 * @param version
	 * @return
	 * @see [类、类#方法、类#成员]
	 */
	public static <T> Workbook creatWorkbook(List<T> data, String sheetName,String version) {
        if(data.size()<1){
            return null;
        }
        return creatWorkbook(data, sheetName, version, getDefaultHeader(data.get(0)));
    }
	
	/**
     * 
     * 根据置顶表头导出
     * <功能详细描述>
     * @param data 数据
     * @param sheetName 工作表名称
     * @param version excel版本
     * @param headerName 表头名称
     * @param headerField 表头
     * @return Workbook
     * @see [类、类#方法、类#成员]
     */
    public  static <T> Workbook creatWorkbook(List<T> data,String sheetName,String version,String[] headerName,String[] headerField){
            if(headerField==null)return null;
            List<HeaderBean> headers = new ArrayList<>();
            for (int i = 0; i < headerField.length; i++) {
                HeaderBean herderBean = new HeaderBean();
                if(StringUtil.isEmpty(headerField[i])){
                    continue;
                }
                herderBean.setField( headerField[i]);
                if(headerName.length>=i+1&&StringUtil.isNotEmpty(headerName[i])){
                    herderBean.setName(headerName[i]);
                }else{
                    herderBean.setName(headerField[i]);
                }
                headers.add(herderBean);
            }
            return creatWorkbook(data, sheetName, version, headers);
    }
	

	private static <T> Workbook creatWorkbook(List<T> data, String sheetName,String version,List<HeaderBean> header) {
	    Workbook wb = null;
        ExcelBean<T> excel = new ExcelBean<>();
        
        if(StringUtil.isNotEmpty(version)){
            if(EXCEL2003.equalsIgnoreCase(version)){
                excel.setVersion(EXCEL2003);
            }else{
                excel.setVersion(EXCEL2007);
            }
        }
        
        SheetBean<T> sheet = new SheetBean<>();
        if(StringUtil.isEmpty(sheetName)){
            sheet.setSheetName(DEFAULT_SHEET_NAME);
        }
        if(CollectionUtils.isEmpty(header)){
            if(CollectionUtils.isEmpty(data)){
                throw new ExcelException(ExcelException.ExcelExceptionEnum.DATA_EMPTY);
            }else{
                sheet.setData(data);
            }
        }else{
            sheet.setHeader(header);
            sheet.setData(data);
        }
        List<SheetBean<T>> sheets = new ArrayList<>();
        sheets.add(sheet);
        excel.setSheets(sheets);
        wb = creatWorkbook(excel);
        return wb;
    }
	/**
	 * 
	 * 根据数据类上的注解获取默认表头
	 * <功能详细描述>
	 * @param t
	 * @return
	 * @see [类、类#方法、类#成员]
	 */
    private static <T> List<HeaderBean> getDefaultHeader(T t) {
        Field[] fields = t.getClass().getFields();
        List<HeaderBean> headers = new ArrayList<>();
        for(Field f : fields){  
            //获取字段中包含fieldMeta的注解  
            try {
               ExportAnnotation meta = f.getAnnotation(ExportAnnotation.class);
               if(meta!=null){
                   HeaderBean header = new HeaderBean();
                   header.setField(f.getName());
                   if(StringUtil.isEmpty(meta.name())){
                       header.setName(f.getName());
                   }else{
                       header.setName(meta.name());
                   }
                   header.setType(meta.tpye());
                   header.setSort(meta.sort());
                   headers.add(header);
               }
            }
            catch (Exception e) {
                e.printStackTrace();
            } 
        }
        Collections.sort(headers, new Comparator<HeaderBean>() {
            public int compare(HeaderBean arg0, HeaderBean arg1) {
                return arg0.getSort()-arg1.getSort();
            }
        });
        return headers;
    }

	
	private  static <T> Workbook creatWorkbook(ExcelBean<T> excel) throws ExcelException{
	    Workbook wb = null;
	    if(excel==null){
	        throw new ExcelException(ExcelException.ExcelExceptionEnum.PARAMETER_MISSING);
	    }
	    if(StringUtil.isEmpty(excel.getVersion())){
	        excel.setVersion(ExcelBean.EXCEL_VERSION_2007);
	    }
	    if(excel.getSheets()==null||excel.getSheets().size()<1){
	        throw new ExcelException(ExcelException.ExcelExceptionEnum.SHEET_EMPTY);
	    }
	    if(excel.getVersion().equalsIgnoreCase(ExcelBean.EXCEL_VERSION_2007)){
//	        if(!(excel.getFileName().endsWith(".xlsx")||excel.getFileName().endsWith(".XLSX"))){
//	            excel.setFileName(excel.getFileName()+".xlsx");
//	        }
	        return creatWorkbook2007(excel);
	    }else{
	        if(!(excel.getFileName().endsWith(".xls")||excel.getFileName().endsWith(".XLS"))){
                excel.setFileName(excel.getFileName()+".xls");
            }
	        //TODO
	        //creatWorkbook2003();
	    }
	    return wb;
	}
	
	private static <T> Workbook creatWorkbook2007(ExcelBean<T> excel) {
	    SXSSFWorkbook wb = null;
	    
	    List<SheetBean<T>> sheets = excel.getSheets();
	    for (int i = 0; i < sheets.size(); i++) {
	        SheetBean<T> sheetBean = sheets.get(i);
            if(sheetBean==null||(CollectionUtils.isEmpty(sheetBean.getData())
                    &&CollectionUtils.isEmpty(sheetBean.getHeader()))){
                continue;
            }
	        if(StringUtil.isEmpty(sheetBean.getSheetName())){
	            sheetBean.setSheetName("Sheet"+(i+1));
	        }
	        wb = new SXSSFWorkbook(500);  //每500行缓存到内存
	        SXSSFSheet sheet = wb.createSheet(sheetBean.getSheetName());
	        
	        if(sheetBean.getHeader()==null||sheetBean.getHeader().size()==0){
	            sheetBean.setHeader(getDefaultHeader(sheetBean));
	        }
	        addSheetHeader2007(wb,sheet,sheetBean.getHeader());
	        addSheetData2007(wb,sheet,sheetBean.getData(),sheetBean.getHeader());
	        
        }
	    return wb;
	}  
	    
    private static <T> List<HeaderBean> getDefaultHeader(SheetBean<T> sheetBean) {
        if (sheetBean == null
            || (CollectionUtils.isEmpty(sheetBean.getData()) && CollectionUtils.isEmpty(sheetBean.getHeader()))) {
            return null;
        }
        List<HeaderBean> headers = new ArrayList<>();
        try {  
            
            BeanInfo beanInfo = Introspector.getBeanInfo(sheetBean.getData().get(0).getClass()); 
            PropertyDescriptor[] propertyDescriptors = beanInfo.getPropertyDescriptors();  
            for (PropertyDescriptor property : propertyDescriptors) {  
                String key = property.getName();  
                
                // 过滤class属性  
                if (!key.equals("class")) {  
                    HeaderBean sheet = new HeaderBean();
                    sheet.setField(key);
                    headers.add(sheet);
                }  
            }  
        } catch (Exception e) {  
            System.out.println("getDefaultHeader Error " + e);  
            e.printStackTrace();
        }  
        return headers;
    }

    /**
	 *     
	 * 添加excel头
	 * <功能详细描述>
	 * @param wb 
	 * @param sheet
	 * @param headers
	 * @see [类、类#方法、类#成员]
	 */
	private static void addSheetHeader2007(SXSSFWorkbook wb, SXSSFSheet sheet, List<HeaderBean> headers) {
        XSSFCellStyle greenStyle = null;
        SXSSFRow row = sheet.createRow(0);
        
        for (int i = 0; i < headers.size(); i++) {
            
            HeaderBean header = headers.get(i);
            if (header == null || header.getField() == null) {
                continue;
            }
            Cell cell = row.createCell(i);
            
            if (StringUtil.isEmpty(header.getName())) {
                cell.setCellValue(header.getField());
            } else {
                cell.setCellValue(header.getName());
            }
            if (greenStyle == null) {
                greenStyle = (XSSFCellStyle)createGreenStyle(wb);
            }
            cell.setCellStyle(greenStyle);
            cell.setCellType(Cell.CELL_TYPE_STRING);
            sheet.setColumnWidth(i, 4000);
	        //增加数据校验
	        if(header.getRowEnum()==null||header.getRowEnum().isEmpty()){
	            continue;
	        }else if(header.getRowEnum().size()<ExcelBean.EXCEL_MAX_ROWENUM){
	          //下拉列表
	          DataValidation data_validation_list = ExcelUtils.setDataValidationList(sheet,convertMap2Strs(header.getRowEnum()), 1,ExcelBean.EXCEL_MAX_CLOUMN_2007,i,i);
              //设置提示内容,标题,内容  
              data_validation_list.createPromptBox("提示", "请选择");  
              data_validation_list.createErrorBox("错误", "请输入有效值");
              data_validation_list.setEmptyCellAllowed(false);
              data_validation_list.setShowErrorBox(true);
              data_validation_list.setShowPromptBox(true);
              //工作表添加验证数据  
              sheet.addValidationData(data_validation_list);
	        }else{
	              //下拉列表过长时,单独增加sheet枚举
                  SXSSFSheet sheet2 = wb.createSheet(header.getName()==null?header.getField():header.getName());
                  //设置头
                  SXSSFRow row2 = sheet2.createRow(0);
                  Cell cell1  = row2.createCell(0);
                  cell1.setCellValue("代码");
                  cell1.setCellStyle(greenStyle);
                  Cell cell2  = row2.createCell(1);
                  cell2.setCellValue("名称");
                  cell2.setCellStyle(greenStyle);
                  String[] nameList= convertMap2Strs(header.getRowEnum());
                  for (int j = 0; j < nameList.length; j++)
                  {
                      SXSSFRow rowJ = sheet2.createRow(j+1);
                      Cell cellA  = rowJ.createCell(0);
                      if(nameList[j].split("-").length!=2){
                          continue;
                      }
                      cellA.setCellValue(nameList[j].split("-")[0]);
                      Cell cellB  = rowJ.createCell(1);
                      cellB.setCellValue(nameList[j].split("-")[1]);
                  }
	        }
	        
	        
        }
    }
	/**
	 * 
	 * 添加EXCEl内容
	 * @param wb
	 * @param sheet
	 * @param list
	 * @param header
	 * @see [类、类#方法、类#成员]
	 */
    private static <T> void addSheetData2007(SXSSFWorkbook wb, SXSSFSheet sheet, List<T> list, List<HeaderBean> header) {
        CellStyle style = null;
        
        for (int j = 0; j < list.size(); j++) {
            if (list.get(j) == null)
                continue;
            Map<String, Object> map = ExcelUtils.transBean2Map(list.get(j));
            if (map.isEmpty())
                return;
            SXSSFRow row = sheet.createRow(j + 1);
            for (int i = 0; i < header.size(); i++) {
                Cell cell = null;
                HeaderBean headerBean = header.get(i);
                if (headerBean == null) {
                    continue;
                }
                if (map.containsKey(headerBean.getField())) {
                    if (headerBean.getType() == Cell.CELL_TYPE_NUMERIC) {
                        cell = row.createCell(i, Cell.CELL_TYPE_NUMERIC);
                        if(map.get(headerBean.getField())==null){
                            cell.setCellValue(0);
                        }else{
                            String val = map.get(headerBean.getField()).toString();
                            cell.setCellValue(Double.valueOf(val));
                        }
                    }
                    else {
                        cell = row.createCell(i, Cell.CELL_TYPE_STRING);
                        if(map.get(headerBean.getField())==null){
                            cell.setCellValue("");
                        }else{
                            String val = map.get(headerBean.getField()).toString();
                            cell.setCellValue(val);
                        }
                    }
                    if (style == null) {
                        style = sheet.getWorkbook().createCellStyle();
                        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
                    }
                    cell.setCellStyle(style);
                }
            }
        }
    }


    /**
     * 
     * 构建模版/导出数据
     * <功能详细描述>
     * @deprecated
     * @param list  初始数据
     * @param response httpresponse
     * @param tableName 表头名称
     * @param tableFild 表头对应数据字段  导出数据类型默认为字符串，field后加 '@'表示该列未数字
     * @param headerCheckList 表头对应的下拉列表数据
     * @param fileName  表格名称
     * @param excelVersion excel版本 07/03
     * @param fileName 文件名称
     * @return
     * @see [类、类#方法、类#成员]
     */
    public static <T> Boolean buildModel(List<T> list,HttpServletResponse response,String[] tableName,String[] tableFild,Map<String,String[]> headerCheckList,String fileName,String excelVersion){
        try {
            //HSSFWorkbook wb = service.getTudimiaoJingyingHssfWorkbook(rp);
           
            //HSSFWorkbook wb = BaseExcelUtil.writeXlsData2003(dataList, header);
            Workbook wb=null;
            if(EXCEL2007.equals(excelVersion)){
                fileName+=".xlsx";
                wb = writeXlsxData2007(list, tableName,tableFild,headerCheckList,fileName);
             }else{
                fileName+=".xls";
                wb = writeXlsData2003(list, tableName,tableFild);
             }
            if(wb==null){
                return false;
            }
            response.setContentType("application/vnd.ms-excel"); 
            //以保存或者直接打开的方式把Excel返回到页面
            response.setHeader("Content-disposition", "attachment; filename=" + URLEncoder.encode(fileName, "UTF-8"));
            ServletOutputStream os =  response.getOutputStream();
            try {
                wb.write(os);
            } catch (OpenXML4JRuntimeException e) {
                os.close();
                return false;
            }
            os.flush();
            os.close();
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        } 
        
    }
    
    /**
     * 
     * 读取文件中数据
     * <功能详细描述>
     * @deprecated
     * @param is 输入流
     * @param cls  bean类型
     * @param field  对应字段
     * @return
     * @throws IOException
     * @throws InstantiationException
     * @throws IllegalAccessException
     * @see [类、类#方法、类#成员]
     */
    @SuppressWarnings("resource")
   public  static <T> List<T>  readGoodsItemFromXls(InputStream is,Class<T> cls,String[] field) throws IOException, InstantiationException, IllegalAccessException {
           
           XSSFWorkbook hssfWorkbook = new XSSFWorkbook(is);
           List<T> list = new ArrayList<T>();
           // 循环工作表Sheet
               XSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
               if (hssfSheet == null) {
                   return null;
               }
               // 循环行Row-从数据行开始
               for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                   T t=cls.newInstance();
                   XSSFRow hssfRow = hssfSheet.getRow(rowNum);
                   HashMap<String, Object> map=new HashMap<String, Object>();
                   //循环row中的每一个单元格
                   for (int i = 0; i < hssfRow.getLastCellNum(); i++) {
                       XSSFCell cell = hssfRow.getCell(i);
                       //格式转换
                       String val="";
                       if(cell!=null){
                           if(cell.getCellType()==Cell.CELL_TYPE_STRING){
                               val=cell.getStringCellValue();
                           }else if(cell.getCellType()==Cell.CELL_TYPE_BOOLEAN){
                               val=cell.getBooleanCellValue()==true?"true":"false";
                           }else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
                               val=new BigDecimal(cell.getNumericCellValue()).toPlainString();
                           }else{
                               cell.setCellType(Cell.CELL_TYPE_STRING);
                               val=cell.getStringCellValue();
                           }
                       }
                       for(int n=0;n<field.length;n++){
                           if(i==n&&!field[n].contains("&")){
                               map.put(field[n], cell==null?"":val);
                           }else if(i==n){
                               map.put(field[n].split("&")[0], cell==null?"":val.split("-")[0]);
                           }
                       }
                   }
                   transMap2Bean(map,t);
                   list.add(t);

               }
           return list;
       }
	
    /**
     * 
     * 导出excel2003
     * <功能详细描述>
     * @param dataList
     * @param tableName
     * @param tableFeild
     * @return
     * @see [类、类#方法、类#成员]
     */
    @SuppressWarnings("unused")
    private static  <T> HSSFWorkbook writeXlsData2003(List<T> dataList,String[] tableName,String[] tableFeild){
        
        
        List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
        if(dataList!=null){
            for (T t : dataList) {
                Map<String, Object> map = new HashMap<String, Object>();
                //BeanUtils.populate(r, map);
                map = ExcelUtils.transBean2Map(t);
                list.add(map);
            }
        }
        List<Map<String, String>> header = new ArrayList<Map<String, String>>();
        for (int i = 0; i < tableName.length; i++) {
            Map<String,String> map = new HashMap<String,String>();
            map.put("name", tableName[i]);
            map.put("field", tableFeild[i]);
            header.add(map);
        }
        
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet();
        HSSFCellStyle style = sheet.getWorkbook().createCellStyle();
        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        HSSFRow row = null;
        // 添加excel头
        row = sheet.createRow(0);
        HSSFCellStyle greenStyle = createGreenStyle(wb);
        if(header!=null){
            Cell cell = null;
            for (int i = 0; i < header.size(); i++) {
                cell = row.createCell(i);
                cell.setCellValue(header.get(i).get("name")==null?"":header.get(i).get("name").toString());
                cell.setCellStyle(greenStyle);
                sheet.setColumnWidth(i, 4000);
            }
        }else if(list.size()>0){
            Cell cell = null;
            Object[] keys = list.get(0).keySet().toArray();
            for (int i = 0; i < keys.length; i++) {
                cell = row.createCell(i);
                cell.setCellValue(keys[i].toString());
            }
        }
        // 添加excel内容
        for (int i = 0; i < list.size(); i++) {
            Map<String, Object> map = list.get(i);
            Set<String> set = map.keySet();
            Object[] keys = set.toArray();
            row = sheet.createRow(i+1);
            Cell cell = null;

            for (int j = 0; j < header.size(); j++) {
                cell = row.createCell(j);
                if(header!=null){
                    String value="";
                    try {
                        value = map.get(header.get(j).get("field")).toString();
                    } catch (Exception e) {
                        value="";
                        //e.printStackTrace();
                    }
                    
                    cell.setCellValue(StringUtil.notNull(value));
                    cell.setCellStyle(style);
                }else{
                    cell.setCellValue(map.get(keys[j].toString()).toString());
                    cell.setCellStyle(style);
                }
        }
        }
       
        return wb;
    }
    /**
     * 
     * 导出2007
     * @param dataList 导出数据
     * @param tableName 文件头
     * @param tableFeild 对应bean属性
     * @param headerCheckList 下拉列表map<Feild,Value>
     * @param fileName 导出文件名
     * @return
     * @see [类、类#方法、类#成员]
     */
    @SuppressWarnings("unused")
    private  static <T> SXSSFWorkbook writeXlsxData2007(List<T> dataList,String[] tableName,String[] tableFeild,Map<String,String[]> headerCheckList,String fileName){
        
        List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
        if(tableName==null){
            return null;  
        }
        if(tableFeild==null){
            return null;  
        }
        if(tableName.length!=tableFeild.length){
            return null;
        }
        
        if(dataList!=null){
            for (T t : dataList) {
                Map<String, Object> map = new HashMap<String, Object>();
                map = ExcelUtils.transBean2Map(t);
                list.add(map);
            }
        }   
        SXSSFWorkbook wb = new SXSSFWorkbook(500);
        SXSSFSheet sheet = wb.createSheet(fileName.split("\\.")[0]);
        XSSFCellStyle greenStyle = (XSSFCellStyle)createGreenStyle(wb);
        List<Map<String, String>> header = new ArrayList<Map<String, String>>();
        for (int i = 0; i < tableName.length; i++) {
            Map<String,String> map = new HashMap<String,String>();
            map.put("name", tableName[i]);
            map.put("field", tableFeild[i]);
            header.add(map);
            if(headerCheckList.containsKey(tableFeild[i])&&headerCheckList.get(tableFeild[i]).length<10){
                //添加验证
                DataValidation data_validation_list = ExcelUtils.setDataValidationList(sheet,headerCheckList.get(tableFeild[i]), 1,1000000,i,i);  
                //设置提示内容,标题,内容  
                data_validation_list.createPromptBox("提示", "请选择");  
                data_validation_list.createErrorBox("错误", "请输入有效值");
                data_validation_list.setEmptyCellAllowed(false);
                data_validation_list.setShowErrorBox(true);
                data_validation_list.setShowPromptBox(true);
                //工作表添加验证数据  
                sheet.addValidationData(data_validation_list);  
            }else if(headerCheckList.containsKey(tableFeild[i])){
                SXSSFSheet sheetName = wb.createSheet(tableName[i]);
                //设置头
                SXSSFRow row = sheetName.createRow(0);
                Cell cell1  = row.createCell(0);
                cell1.setCellValue("代码");
                cell1.setCellStyle(greenStyle);
                Cell cell2  = row.createCell(1);
                cell2.setCellValue("名称");
                cell2.setCellStyle(greenStyle);
                String[] nameList=headerCheckList.get(tableFeild[i]);
                for (int j = 0; j < nameList.length; j++)
                {
                    SXSSFRow rowJ = sheetName.createRow(j+1);
                    Cell cellA  = rowJ.createCell(0);
                    if(nameList[j].split("-").length<1){
                        continue;
                    }
                    cellA.setCellValue(nameList[j].split("-")[0]);
                    Cell cellB  = rowJ.createCell(1);
                    cellB.setCellValue(nameList[j].split("-")[1]);
                }
                
            }
        }
        
        
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        SXSSFRow row = null;
        // 添加excel头
        row = sheet.createRow(0);
        
        if(header!=null){
            Cell cell = null;
            for (int i = 0; i < header.size(); i++) {
                cell = row.createCell(i);
                cell.setCellValue(header.get(i).get("name")==null?"":header.get(i).get("name").toString());
                cell.setCellStyle(greenStyle);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                sheet.setColumnWidth(i, 4000);
            }
        }else if(list.size()>0){
            Cell cell = null;
            Object[] keys = list.get(0).keySet().toArray();
            for (int i = 0; i < keys.length; i++) {
                cell = row.createCell(i);
                cell.setCellValue(keys[i].toString());
                cell.setCellType(Cell.CELL_TYPE_STRING);
            }
        }
       
        
        // 添加excel内容
        for (int i = 0; i < list.size(); i++) {
            Map<String, Object> map = list.get(i);
            Set<String> set = map.keySet();
            Object[] keys = set.toArray();
            row = sheet.createRow(i+1);
            row.setRowStyle(style);
            Cell cell = null;
            
            for (int j = 0; j < header.size(); j++) {
                if(header!=null){
                    String value="";
                    try {
                        String key = header.get(j).get("field");
                        if(key.contains("@")){
                            cell = row.createCell(j, Cell.CELL_TYPE_NUMERIC);
                            value = map.get(key.replace("@", "")).toString();
                            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                            cell.setCellValue(Double.valueOf(value));
                            cell.setCellStyle(style);
                        }else{
                            cell = row.createCell(j, Cell.CELL_TYPE_STRING);
                            value = map.get(key).toString();
                            cell.setCellValue(value);
                            cell.setCellStyle(style);
                        }
                    } catch (Exception e) {
                        value="";
                        //e.printStackTrace();
                    }
                    

                }else{
                    cell.setCellValue(map.get(keys[j].toString()).toString());
                    cell.setCellStyle(style);
                }
        }
        }
       
        return wb;
    }
	/**
	 * 
	 * 保存数据到excel文件
	 * @param list 数据
	 * @param filePath 导保存文件目录
	 * @param out 文件输出流
	 * @param result 返回结果
	 * @param header excel头
	 * @return
	 * @see [类、类#方法、类#成员]
	 */
	@SuppressWarnings("unused")
    private static Map<String, Object> writeXlsxData(List<Map<String, Object>> list,String filePath,FileOutputStream out,Map<String, Object> result,ArrayList<String> header){
		XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();
        
        XSSFRow row = null;
        // 添加excel头
        row = sheet.createRow(0);
        if(header!=null&&header.size()>=list.get(0).keySet().size()){
        	Cell cell = null;
        	for (int i = 0; i < header.size(); i++) {
        		cell = row.createCell(i);
    			cell.setCellValue(header.get(i));
			}
        }else{
        	Cell cell = null;
        	Object[] keys = list.get(0).keySet().toArray();
        	for (int i = 0; i < keys.length; i++) {
        		cell = row.createCell(i);
    			cell.setCellValue(keys[i].toString());
			}
        }
        // 添加excel内容
        for (int i = 0; i < list.size(); i++) {
        	Map<String, Object> map = list.get(i);
        	Set<String> set = map.keySet();
        	Object[] keys = set.toArray();
        	row = sheet.createRow(i+1);
        	Cell cell = null;
        	for (int j = 0; j < keys.length; j++) {
        			cell = row.createCell(j);
        			if(header!=null&&header.size()>=list.get(0).keySet().size()){
        				cell.setCellValue(map.get(header.get(j)).toString());
        			}else{
        				cell.setCellValue(map.get(keys[j].toString()).toString());
        			}
			}
		}
        
		try {
			 wb.write(out);
			 wb.close();
			 out.flush();
		     out.close();
		     
		     result.put("code", 1);
			 result.put("message", "成功导出"+list.size()+"条记录到"+filePath);
		} catch (Exception e) {
			 result.put("code", -200);
			 result.put("message", e.getMessage());
		}
       
        return result;
	}
	/**
	 * 保存数据到excel文件
     * @param list 数据
     * @param filePath 导保存文件目录
     * @param out 文件输出流
     * @param result 返回结果
     * @param header excel头
	 * @return
	 * @see [类、类#方法、类#成员]
	 */
	@SuppressWarnings("unused")
    private static Map<String, Object> writeXlsData(List<Map<String, Object>> list,String filePath,FileOutputStream out,Map<String, Object> result,ArrayList<String> header){
		HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet();
        
        HSSFRow row = null;
        // 添加excel头
        row = sheet.createRow(0);
        if(header!=null&&header.size()>=list.get(0).keySet().size()){
        	Cell cell = null;
        	for (int i = 0; i < header.size(); i++) {
        		cell = row.createCell(i);
    			cell.setCellValue(header.get(i));
			}
        }else{
        	Cell cell = null;
        	Object[] keys = list.get(0).keySet().toArray();
        	for (int i = 0; i < keys.length; i++) {
        		cell = row.createCell(i);
    			cell.setCellValue(keys[i].toString());
			}
        }
        // 添加excel内容
        for (int i = 0; i < list.size(); i++) {
        	Map<String, Object> map = list.get(i);
        	Set<String> set = map.keySet();
        	Object[] keys = set.toArray();
        	row = sheet.createRow(i+1);
        	Cell cell = null;
        	for (int j = 0; j < keys.length; j++) {
        			cell = row.createCell(j);
        			if(header!=null&&header.size()>=list.get(0).keySet().size()){
        				cell.setCellValue(map.get(header.get(j)).toString());
        			}else{
        				cell.setCellValue(map.get(keys[j].toString()).toString());
        			}
			}
		}
        
		try {
			 wb.write(out);
			 wb.close();
			 out.flush();
		     out.close();
		     
		     result.put("code", 1);
			 result.put("message", "成功导出"+list.size()+"条记录到"+filePath);
		} catch (Exception e) {
			 result.put("code", -200);
			 result.put("message", e.getMessage());
		}
		return result;
	}
	
	
	
	/**
	 * 
	 * 对象转换称map
	 * @param obj
	 * @return
	 * @see [类、类#方法、类#成员]
	 */
    private static Map<String, Object> transBean2Map(Object obj) {  
    	  
        if(obj == null){  
            return null;  
        }          
        Map<String, Object> map = new HashMap<String, Object>();  
        try {  
            BeanInfo beanInfo = Introspector.getBeanInfo(obj.getClass());  
            PropertyDescriptor[] propertyDescriptors = beanInfo.getPropertyDescriptors();  
            for (PropertyDescriptor property : propertyDescriptors) {  
                String key = property.getName();  
  
                // 过滤class属性  
                if (!key.equals("class")) {  
                    // 得到property对应的getter方法  
                    Method getter = property.getReadMethod();  
                    Object value = getter.invoke(obj);  
                    if(value==null){
                    	continue;
                    }
                    map.put(key, value);  
                }  
  
            }  
        } catch (Exception e) {  
            System.out.println("transBean2Map Error " + e);  
        }  
  
        return map;  
  
    }
    /**
     * 
     * 设置样式
     * @param wb
     * @return
     * @see [类、类#方法、类#成员]
     */
    private  static CellStyle createGreenStyle(SXSSFWorkbook wb) {
		//设置字体
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 11); // 字体高度
        font.setFontName("宋体"); // 字体
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体显示
        
     	CellStyle greenStyle = wb.createCellStyle();  
     	greenStyle.setFillBackgroundColor(HSSFCellStyle.LEAST_DOTS);
     	greenStyle.setFillPattern(HSSFCellStyle.LEAST_DOTS);
     	greenStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式  
     	greenStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
     	greenStyle.setBottomBorderColor(HSSFColor.BLACK.index);
     	greenStyle.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
     	greenStyle.setLeftBorderColor(HSSFColor.BLACK.index);
     	greenStyle.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
     	greenStyle.setRightBorderColor(HSSFColor.BLACK.index);
     	greenStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
     	greenStyle.setTopBorderColor(HSSFColor.BLACK.index);
        greenStyle.setFont(font);
        greenStyle.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
        greenStyle.setFillBackgroundColor(HSSFColor.LIGHT_GREEN.index);
        greenStyle.setWrapText(true); 
        
        return greenStyle;
	}
    /**
     * 
     * 设置Excel样式2003
     * @param wb
     * @return
     * @see [类、类#方法、类#成员]
     */
    private  static HSSFCellStyle createGreenStyle(HSSFWorkbook wb) {
        //设置字体
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 11); // 字体高度
        font.setFontName("宋体"); // 字体
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体显示
        
        HSSFCellStyle greenStyle = wb.createCellStyle();  
        greenStyle.setFillBackgroundColor(HSSFCellStyle.LEAST_DOTS);
        greenStyle.setFillPattern(HSSFCellStyle.LEAST_DOTS);
        greenStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式  
        greenStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
        greenStyle.setBottomBorderColor(HSSFColor.BLACK.index);
        greenStyle.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
        greenStyle.setLeftBorderColor(HSSFColor.BLACK.index);
        greenStyle.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
        greenStyle.setRightBorderColor(HSSFColor.BLACK.index);
        greenStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
        greenStyle.setTopBorderColor(HSSFColor.BLACK.index);
        greenStyle.setFont(font);
        greenStyle.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
        greenStyle.setFillBackgroundColor(HSSFColor.LIGHT_GREEN.index);
        greenStyle.setWrapText(true); 
        
        return greenStyle;
    }
    /**
     * 
     * 设置Excel样式2007
     * @param wb
     * @return
     * @see [类、类#方法、类#成员]
     */
    @SuppressWarnings("unused")
    private  static XSSFCellStyle createGreenStyle(XSSFWorkbook wb) {
		//设置字体
        XSSFFont font = wb.createFont();
        font.setFontHeightInPoints((short) 11); // 字体高度
        font.setFontName("宋体"); // 字体
        font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);//粗体显示
        
     	XSSFCellStyle greenStyle = wb.createCellStyle();  
     	greenStyle.setFillBackgroundColor(XSSFCellStyle.LEAST_DOTS);
     	greenStyle.setFillPattern(XSSFCellStyle.LEAST_DOTS);
     	greenStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式  
     	greenStyle.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);
     	greenStyle.setBottomBorderColor(HSSFColor.BLACK.index);
     	greenStyle.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
     	greenStyle.setLeftBorderColor(HSSFColor.BLACK.index);
     	greenStyle.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
     	greenStyle.setRightBorderColor(HSSFColor.BLACK.index);
     	greenStyle.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
     	greenStyle.setTopBorderColor(HSSFColor.BLACK.index);
        greenStyle.setFont(font);
        greenStyle.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
        greenStyle.setFillBackgroundColor(HSSFColor.LIGHT_GREEN.index);
        greenStyle.setWrapText(true); 
        
        return greenStyle;
	}

    
	
	 /**
	  * 
	  * 设置excel数据有效性
	  * <功能详细描述>
	  * @param sheet
	  * @param textlist
	  * @param firstRow 起始行
	  * @param firstCol 终止行
	  * @param endRow   起始列
	  * @param endCol   终止列
	  * @return
	  * @see [类、类#方法、类#成员]
	  */
	 private static DataValidation setDataValidationList(SXSSFSheet sheet,String[] textlist,int firstRow,int endRow, int firstCol,int endCol){  
	        //设置下拉列表的内容  
//	        String[] textlist={"列表1","列表2","列表3","列表4","列表5"};  
	        // 加载下拉列表内容
	        
	        DataValidationHelper helper = sheet.getDataValidationHelper();
	        DataValidationConstraint constraint = helper.createExplicitListConstraint(textlist);
	        constraint.setExplicitListValues(textlist);
	        //设置数据有效性加载在哪个单元格上。  
	          
	     // 设置数据有效性加载在哪个单元格上。
	     // 四个参数分别是：起始行、终止行、起始列、终止列
	     CellRangeAddressList regions = new CellRangeAddressList(firstRow,endRow,firstCol,endCol);

	     // 数据有效性对象
	     DataValidation data_validation = helper.createValidation(constraint, regions);
	          
	     return data_validation;  
	} 
	 
	/**
	 *  
	 * map转对象
	 * @param map
	 * @param obj
	 * @see [类、类#方法、类#成员]
	 */
    private static void transMap2Bean(Map<String, Object> map, Object obj) {  
	        
	        try {  
	            BeanInfo beanInfo = Introspector.getBeanInfo(obj.getClass());  
	            PropertyDescriptor[] propertyDescriptors = beanInfo.getPropertyDescriptors();  
	  
	            for (PropertyDescriptor property : propertyDescriptors) {  
	                String key = property.getName();  
	  
	                if (map.containsKey(key)) {  
	                    Object value = map.get(key);  
	                    // 得到property对应的setter方法  
	                    
	                    Method setter = property.getWriteMethod();  
	                    String type= property.getPropertyType().toString();
	                    
	                    
                            if(type.contains("String")){
                                setter.invoke(obj, value.toString());
                            }else if(type.contains("BigDecimal")){
                                setter.invoke(obj, new BigDecimal(value.toString()));
                            }else if(type.contains("Integer")){
                                setter.invoke(obj, Integer.parseInt(value.toString()));
                            }else if(type.contains("Long")){
                                BigDecimal bd;
                                try
                                {
                                    bd = new BigDecimal(value.toString());
                                    setter.invoke(obj, Long.parseLong(bd.toPlainString()));
                                }
                                catch (Exception e)
                                {
                                    setter.invoke(obj,Long.MIN_VALUE);
                                }
                                
                            }else if(type.contains("int")){
                                if(value instanceof String){
                                    try
                                    {
                                        String val=((String)value).split("\\.")[0];
                                        setter.invoke(obj, Integer.parseInt(val));
                                    }
                                    catch (Exception e)
                                    {
                                        setter.invoke(obj, Long.MIN_VALUE);
                                    }
                                }
                            }else {
                                setter.invoke(obj, value);
                            }
                        }
                        
	  
	            }  
	  
	        } catch (Exception e) {  
	            e.printStackTrace();  
	        }  
	        return;  
	  
	    }
	
	/**
	 * 
	 * 数据转ArrayList
	 * @param strs
	 * @return
	 * @see [类、类#方法、类#成员]
	 */
	@SuppressWarnings("unused")
    private static ArrayList<String> convertStrs2ArrayList(String[] strs){
		if(strs==null){
			return null;
		}
		ArrayList<String> list = new ArrayList<String>();
		list.addAll(Arrays.asList(strs));
	    return list;
	}
	
	private static String[] convertMap2Strs(Map<String, String> map){
	    if(map==null){
	        return new String[]{};
	    }
	    Set<String> set = map.keySet();
        List<String> list = new ArrayList<>();
        for (String key : set) {
            list.add(map.get(key));
        }
        return list.toArray(new String[]{});
    }

}
