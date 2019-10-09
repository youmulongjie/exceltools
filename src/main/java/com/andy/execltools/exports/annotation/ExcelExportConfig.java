/**
 * @author ：Andy.wang<br>
 * @createDate ：2015年12月2日下午2:22:43<br>
 */
package com.andy.execltools.exports.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 描述：Excel 导出注解类 <br>
 * <br>
 * 1、导出的类必须添加该注解类 <br>
 * 2、默认导出的Excel第0行为标题行，对象记录行从第1行开始<br>
 * 
 * @file ：ExcelExportConfig.java<br>
 * @author ：Andy.wang<br>
 * @createDate ：2015年12月2日下午2:22:43<br>
 * 
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.TYPE })
public @interface ExcelExportConfig {
	/**
	 * 
	 * 描述： 生成Excel文件的 标题所在的行数，默认为0<br>
	 * 
	 * @method ：headerRow<br>
	 * @author ：Andy.wang<br>
	 * @createDate ：2015年12月2日下午2:23:29 <br>
	 * @return 生成Excel文件的 标题所在的行数
	 */
	int headerRow() default 0;

	/**
	 * 
	 * 描述：生成Excel文件的 对象记录所在的行数，默认从1开始 <br>
	 * 
	 * @method ：lineStartRow<br>
	 * @author ：Andy.wang<br>
	 * @createDate ：2015年12月2日下午2:23:56 <br>
	 * @return 生成Excel文件的 对象记录所在的行数
	 */
	int lineStartRow() default 1;
}
