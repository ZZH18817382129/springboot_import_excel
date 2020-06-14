package com.ryan.entry;

public enum ExcelSheetSettingEnum {
    REPORT_TEST("report_test", "Excel文件名称", new String[]{"Sheet名称1", "Sheet名称2"}, new String[]{"标题1", "标题2"}, new String[][]{
            {"字段名称A", "字段名称B", "字段名称C", "字段名称D"},
            {"字段名称A", "字段名称B", "字段名称C", "字段名称D"}
    }),
    REPORT_TEST2("report_test2", "Excel文件名称", new String[]{"标题1", "标题2"}),
    REPORT_TEST3("report_test3", "Excel文件名称");

    private String code;//代码标识(必选)
    private String filename;//代码标识(必选)
    private String[] sheetnames;//Sheet名称(可选)
    private String[] titles;//Sheet标题(可选)
    private String[][] headers;//表头名称(可选)

    ExcelSheetSettingEnum(String code, String filename) {
        this.code = code;
        this.filename = filename;
    }

    ExcelSheetSettingEnum(String code, String filename, String[] titles) {
        this.code = code;
        this.filename = filename;
        this.titles = titles;
    }

    ExcelSheetSettingEnum(String code, String filename, String[] sheetnames, String[] titles, String[][] headers) {
        this.code = code;
        this.filename = filename;
        this.sheetnames = sheetnames;
        this.titles = titles;
        this.headers = headers;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getFilename() {
        return filename;
    }

    public void setFilename(String filename) {
        this.filename = filename;
    }

    public String[] getSheetnames() {
        return sheetnames;
    }

    public void setSheetnames(String[] sheetnames) {
        this.sheetnames = sheetnames;
    }

    public String[] getTitles() {
        return titles;
    }

    public void setTitles(String[] titles) {
        this.titles = titles;
    }

    public String[][] getHeaders() {
        return headers;
    }

    public void setHeaders(String[][] headers) {
        this.headers = headers;
    }
}