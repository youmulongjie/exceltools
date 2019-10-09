/**
 * @author ：Andy.wang<br>
 * @createDate ：2015年12月2日下午2:26:29<br>
 */
package com.andy.execltools.exports.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 描述：Excel 导出属性注解类<br>
 * <br>
 * 1、导出的类必须添加注解类ExcelExportConfig<br>
 * 2、该注解类用在类属性上，获取Excel的列标题头和列索引（从0开始）<br>
 * 
 * @file ：ExcelExportCol.java<br>
 * @author ：Andy.wang<br>
 * @createDate ：2015年12月2日下午2:26:29<br>
 * 
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.FIELD })
public @interface ExcelExportCol {
	/**
	 * 
	 * 描述：列标题头说明 <br>
	 * 
	 * @method ：colHeaderDesc<br>
	 * @author ：Andy.wang<br>
	 * @createDate ：2015年12月2日下午2:27:15 <br>
	 * @return 列标题头说明
	 */
	String colHeaderDesc();

	/**
	 * 
	 * 描述：所在列索引，下标从0开始 <br>
	 * 
	 * @method ：cols<br>
	 * @author ：Andy.wang<br>
	 * @createDate ：2015年12月2日下午2:27:38 <br>
	 * @return 所在列索引，下标从0开始
	 */
	int cols();

	/**
	 * 
	 * <p>
	 * 描述：映射配置，默认为空
	 * </p>
	 * <ul>
	 * mapper 参数说明：
	 * <li>1、该参数默认为空</li>
	 * <li>2、该参数采用“key=value”形式，如有多个用“,”隔开；例如“1=有效,0=无效”</li>
	 * </ul>
	 * 
	 * <pre>
	 * 	例如：设置@ExcelExportCol(cols = 5, colHeaderDesc = "状态", mapper = "1=有效,0=无效")，<br>
	 * 	说明 如果该值为1，则返回“有效”；如果值为0，则返回“无效”，其他值直接返回
	 * </pre>
	 * 
	 * @method ：cols<br>
	 * @author ：Andy.wang<br>
	 * @Date 2016年11月2日 上午11:25:04 <br>
	 * @return
	 */
	String mapper() default "";

	/**
	 * 
	 * <p>
	 * 描述: 日期格式，默认为“yyyy-mm-dd”
	 * </p>
	 * 
	 * @Date 2017年11月9日下午12:36:24 <br>
	 * @return
	 */
	String dateFormat() default "yyyy-mm-dd";
}
