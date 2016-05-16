package SCode.Scode;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

import com.SCode.excel.ExcelUtils;
import com.SCode.excel.ExportAnnotation;

/**
 * Unit test for simple App.
 */
public class AppTest 
{
    
    
    public void testExcel() throws IOException{
       User u = new User("shi","12");
       List<User> data = new ArrayList<>();
       data.add(u);
//       Workbook wb = ExcelUtils.creatWorkbook(data);
       Workbook wb = ExcelUtils.creatWorkbook(data, "测试", "2007");
       File file = new File("D:\\data\\test.xlsx");
       FileOutputStream fileOutputStream = new FileOutputStream(file);
       wb.write(fileOutputStream);
       fileOutputStream.flush();
       fileOutputStream.close();
    }
    
    public static void main(String[] args) {
        try {
            new AppTest().testExcel();

        }
        catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

}










