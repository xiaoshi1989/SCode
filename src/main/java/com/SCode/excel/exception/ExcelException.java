package com.SCode.excel.exception;


public class ExcelException extends RuntimeException {
    private static final long serialVersionUID = 6145562528625520683L;

    private String code;
    private Object[] params;

    public ExcelException(String code) {
        this(code, null, null);
    }

    public ExcelException(String code, Object param) {
        this(code, new Object[] { param });
    }
    
    public ExcelException(ExcelException.ExcelExceptionEnum num) {
        this(num.code+"",num.message);
    }

    public ExcelException(String code, Object[] params) {
        this(code, params, null);
    }

    public ExcelException(String code, Object[] params, Throwable cause) {
        super(code, cause);
        this.code = code;
        this.params = params;
    }

    public ExcelException(String code, Throwable cause) {
        this(code, null, cause);
    }

    public String getCode() {
        return this.code;
    }

    public Object[] getParams() {
        return this.params;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public void setParams(String[] params) {
        this.params = params;
    }
    
    public static enum ExcelExceptionEnum{
        SECCUSS(1,"导出成功"),
        FAIL(0,"导出失败"),
        PARAMETER_MISSING(100,"参数异常");
        private int code;
        private String message;
        private ExcelExceptionEnum(int code,String message){
            this.code = code;
            this.message =message;
        }
    }
    
    public static void main(String[] args) {
        int n = ExcelException.ExcelExceptionEnum.SECCUSS.code;
        System.out.println(n);
    }
}
