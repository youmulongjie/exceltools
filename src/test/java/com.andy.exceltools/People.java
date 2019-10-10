package com.andy.exceltools;

import com.andy.execltools.exports.annotation.ExcelExportCol;
import com.andy.execltools.exports.annotation.ExcelExportConfig;
import com.andy.execltools.imports.annotation.ExcelImportCol;
import com.andy.execltools.imports.annotation.ExcelImportConfig;

import java.util.Date;

/**
 * Description: 导入导出 测试实体类（导入、导出可以不是一个实体，本例只是测试方便）
 * Author: Andy.wang
 * Date: 2019/10/9 17:52
 */
@ExcelImportConfig
@ExcelExportConfig
public class People {

    @ExcelImportCol(col = 0)
    @ExcelExportCol(colHeaderDesc = "ID", cols = 0)
    private int id;

    @ExcelImportCol(col = 1)
    @ExcelExportCol(colHeaderDesc = "姓名", cols = 1)
    private String name;

    @ExcelImportCol(col = 2)
    @ExcelExportCol(colHeaderDesc = "生日", cols = 2)
    private Date birthday;

    @ExcelExportCol(colHeaderDesc = "性别", cols = 3, mapper = "1=男,2=女")
    private int flag;

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    public int getFlag() {
        return flag;
    }

    public void setFlag(int flag) {
        this.flag = flag;
    }

    @Override
    public String toString() {
        return "People{" +
                "id=" + id +
                ", name='" + name + '\'' +
                ", birthday=" + birthday +
                ", flag=" + flag +
                '}';
    }
}
