package SCode.Scode;

import java.util.ArrayList;
import java.util.List;

import com.SCode.excel.ExcelUtils;
import com.SCode.excel.bean.ExcelBean;
import com.SCode.excel.bean.SheetBean;
import com.SCode.excel.exception.ExcelException;

/**
 * Unit test for simple App.
 */
public class AppTest 
{
    
    
    public void testExcel(){
        ExcelUtils<User> util = new ExcelUtils<>();
        
        ExcelBean<User> bean = new ExcelBean<>();
        bean.setFilePath("D:\\data");
        List<SheetBean<User>> sheets = new ArrayList<>();
        SheetBean<User> sheet = new SheetBean<>();
//        sheet.s
//        bean.setSheets(sheets);
        try {
            util.creatWorkbook(bean);
        }
        catch (ExcelException e) {
            e.printStackTrace();
        }
    }
    
    

}
class User {
    private String name;
    private String age;
    
    public String getName() {
        return name;
    }
    
    public String getAge() {
        return age;
    }
    
    public void setName(String name) {
        this.name = name;
    }
    
    public void setAge(String age) {
        this.age = age;
    }
    
}