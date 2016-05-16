package SCode.Scode;

import com.SCode.excel.ExportAnnotation;

public class User {
    @ExportAnnotation(name="名称",sort=2)
    public String name;
    @ExportAnnotation(sort=1)
    public String age;
    
    
    
    public User(String name, String age) {
        super();
        this.name = name;
        this.age = age;
    }

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
