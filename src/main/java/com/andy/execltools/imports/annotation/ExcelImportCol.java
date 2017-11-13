/**
 * @author ：wanglongjie<br>
 * @createDate ：2015年12月2日上午10:44:28<br>
 */
package com.andy.execltools.imports.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 描述：Excel 导入属性注解类 <br>
 * <br>
 * 1、导入的类必须添加注解类ExcelImportConfig<br>
 * 2、该注解类用在类属性上，获取Excel所在列的记录<br>
 * 
 * @file ：ExcelImportCol.java<br>
 * @author ：wanglongjie<br>
 * @createDate ：2015年12月2日上午10:44:28<br>
 * 
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.FIELD })
public @interface ExcelImportCol {
	/**
	 * 
	 * 描述：Excel 所在列 <br>
	 * 
	 * @method ：col<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日上午10:44:59 <br>
	 * @return Excel 所在列
	 */
	int col();
}
