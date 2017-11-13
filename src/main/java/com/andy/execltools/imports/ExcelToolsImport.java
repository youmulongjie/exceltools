/**
 * @author ：wanglongjie<br>
 * @createDate ：2015年12月2日上午10:18:29<br>
 */
package com.andy.execltools.imports;

import java.beans.PropertyDescriptor;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.andy.execltools.imports.annotation.ExcelImportCol;
import com.andy.execltools.imports.annotation.ExcelImportConfig;

/**
 * 描述： Excel 导入工具类<br>
 * <br>
 * 方法一：excelImport(InputStream, Class) : 将 文件流 转化为 List对象集合，sheet索引默认为0；<br>
 * 方法二：excelImport(InputStream, Class, int) : 将 文件流 转化为 List对象集合，可设置sheet索引位置<br>
 * 方法三：excelImport(InputStream, Map<Integer, Class<?>> map) : 将 文件流 转化为
 * List对象集合,可同时读取多个sheet<br>
 * 
 * @file ：ExcelToolsImport.java<br>
 * @author ：wanglongjie<br>
 * @createDate ：2015年12月2日上午10:18:29<br>
 */
public class ExcelToolsImport {
	private static Logger logger = LoggerFactory
			.getLogger(ExcelToolsImport.class);

	/**
	 * 缓存每个类中带有 ExcelExportCol 注解的属性List列表，提高效率<br>
	 * 
	 * <pre>
	 *  key：注解 ExcelExportConfig 的Class类的全类名称
	 *  value：该Class类中 带有ExcelExportCol 注解的属性List列表
	 * </pre>
	 */
	private static Map<String, List<Field>> excelImportColAnnoFieldsMap = Collections
			.synchronizedMap(new HashMap<String, List<Field>>());

	/**
	 * 
	 * 描述：获取Excel的载体实体类集合，默认导入Excel的sheet索引值为0 <br>
	 * 
	 * @method ：excelImport<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日上午11:14:59 <br>
	 * @param fileInputStream
	 *            ：导入Excel生成的文件流
	 * @param cla
	 *            ：导入Excel的载体实体类
	 * @return 载体实体类集合
	 * @throws Exception
	 */
	public static <T> List<T> excelImport(InputStream fileInputStream,
			Class<T> cla) throws Exception {
		return excelImport(fileInputStream, cla, 0);
	}

	/**
	 * 
	 * 描述：获取Excel指定的sheet索引的载体实体类集合 <br>
	 * 
	 * @method ：excelImport<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日上午11:08:17 <br>
	 * @param fileInputStream
	 *            ：导入Excel生成的文件流
	 * @param cla
	 *            ：导入Excel的载体实体类
	 * @param sheetIndex
	 *            ：导入Excel的sheet索引值
	 * @return 载体实体类集合
	 * @throws Exception
	 */
	public static <T> List<T> excelImport(InputStream fileInputStream,
			Class<T> cla, int sheetIndex) throws Exception {
		checkValidate(fileInputStream, cla);
		Workbook workbook = WorkbookFactory.create(fileInputStream);

		return excelImport(workbook, cla, sheetIndex);

	}

	/**
	 * 
	 * <p>
	 * 描述：获取Excel（包含多个sheet）的载体实体类集合
	 * </p>
	 * 
	 * @Date 2017年5月24日下午4:21:12 <br>
	 * @param fileInputStream
	 *            导入Excel生成的文件流
	 * @param map
	 *            key为sheet索引，value为实体类Class
	 * @return
	 * @throws Exception
	 */
	public static Map<Integer, List<?>> excelImport(
			InputStream fileInputStream, Map<Integer, Class<?>> map)
			throws Exception {
		Workbook workbook = WorkbookFactory.create(fileInputStream);

		Map<Integer, List<?>> returnMap = new HashMap<Integer, List<?>>();

		Iterator<Entry<Integer, Class<?>>> it = map.entrySet().iterator();
		Entry<Integer, Class<?>> entry = null;

		@SuppressWarnings("rawtypes")
		List<?> list = new ArrayList();
		Integer sheetIndex = 0;
		Class<?> cla = null;
		while (it.hasNext()) {
			entry = it.next();
			sheetIndex = entry.getKey();
			cla = entry.getValue();

			list = excelImport(workbook, cla, sheetIndex);
			returnMap.put(sheetIndex, list);
		}

		return returnMap;
	}

	private static <T> List<T> excelImport(Workbook workbook, Class<T> cla,
			int sheetIndex) throws Exception {
		logger.info("解析 第[" + sheetIndex + "] 个sheet 开始....");
		logger.info("对应的Java类型为：" + cla.getName());

		Sheet sheet = workbook.getSheetAt(sheetIndex);

		// 获取最大行和开始行
		final int rows = sheet.getLastRowNum();
		int startLine = getStartLine(cla);

		List<T> list = new ArrayList<T>();
		Row row = null;
		T t = null;
		for (int i = startLine; i <= rows; i++) {
			row = sheet.getRow(i);
			if (null != row) {
				t = addLine2List(row, cla);
				if (validateNotNull(t)) {
					list.add(t);
				}
			}
		}
		logger.info("解析 第[" + sheetIndex + "] 个sheet 结束....");
		return list;
	}

	/**
	 * 
	 * 描述：读取行，转化为指定的对象 <br>
	 * 
	 * @method ：addLine2List<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日上午11:27:04 <br>
	 * @param row
	 *            ： Excel Row行对象
	 * @param cla
	 *            ：导入Excel的载体实体类
	 * @return
	 * @throws Exception
	 */
	private static <T> T addLine2List(Row row, Class<T> cla) throws Exception {
		T t = cla.newInstance();
		List<Field> list = getExcelImportColAnnoFields(cla);
		for (Field field : list) {
			setCell2Obj(field, row, t);
		}
		return t;
	}

	/**
	 * 
	 * 描述：读取单元格，设置实体属性值 <br>
	 * 
	 * @method ：setCell2Obj<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日下午12:47:32 <br>
	 * @param field
	 *            ：实体对象中带有ExcelImportCol注解的属性
	 * @param row
	 *            ：Excel Row 行对象
	 * @param t
	 *            ：封装的实体对象
	 * @return
	 * @throws Exception
	 */
	private static <T> T setCell2Obj(Field field, Row row, T t)
			throws Exception {
		// 获取列索引、单元格
		int col = field.getAnnotation(ExcelImportCol.class).col();
		Cell cell = row.getCell(col);

		if (null != cell) {
			String typeName = field.getType().getSimpleName();
			// 获取属性的写入方法
			String propertyName = field.getName();
			PropertyDescriptor pd = new PropertyDescriptor(propertyName,
					t.getClass());
			Method m = pd.getWriteMethod();

			try {
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_STRING:
					// 字符串
					String value = cell.getRichStringCellValue().getString();
					m.invoke(t, value);
					break;
				case Cell.CELL_TYPE_NUMERIC:
					// 数字 | 日期
					if (DateUtil.isCellDateFormatted(cell)) {
						Date date = cell.getDateCellValue();
						m.invoke(t, date);
					} else {
						Double d = cell.getNumericCellValue();
						if (BigDecimal.class.getSimpleName().equals(typeName)) {
							BigDecimal bigDecimal = new BigDecimal(d);
							m.invoke(t, bigDecimal);
						}
						if (Double.class.getSimpleName().equals(typeName)
								|| "double".equals(typeName)) {
							Double d1 = new Double(d);
							m.invoke(t, d1);
						}
						if (Float.class.getSimpleName().equals(typeName)
								|| "float".equals(typeName)) {
							Float f = new Float(d);
							m.invoke(t, f);
						}
						if (Integer.class.getSimpleName().equals(typeName)
								|| "int".equals(typeName)) {
							Integer i = new BigDecimal(d).intValue();
							m.invoke(t, i);
						}
						if (Long.class.getSimpleName().equals(typeName)
								|| "long".equals(typeName)) {
							Long l = new BigDecimal(String.valueOf(d))
									.longValue();
							m.invoke(t, l);
						}
						if (String.class.getSimpleName().equals(typeName)
								|| "String".equals(typeName)) {
							DecimalFormat df = new DecimalFormat("0");
							String s = df.format(d);

							// String s = String.valueOf(d);
							// 如果小数点后为0，则将后面的值去掉
							if (s.endsWith(".0")) {
								s = s.substring(0, s.length() - 2);
							} else if (s.endsWith(".00")) {
								s = s.substring(0, s.length() - 3);
							}
							m.invoke(t, s);
						}
					}
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					// boolean 类型
					boolean b = cell.getBooleanCellValue();
					m.invoke(t, b);
					break;
				default:
					// 字符串
					String defaultValue = cell.getRichStringCellValue()
							.getString();
					m.invoke(t, defaultValue);
					break;
				}
			} catch (Exception e) {
				logger.error("方法【" + m.getName() + "】参数类型出错：" + e.getMessage());

				StringBuffer sb = new StringBuffer(16);
				sb.append("第").append(row.getRowNum()).append("]行，第[")
						.append(col).append("]列 数据格式不正确！");
				logger.error(sb.toString());

				throw new Exception(sb.toString());
			}

		}
		return t;
	}

	/**
	 * 
	 * 描述：获取开始行
	 * 
	 * <br>
	 * 
	 * @method ：getStartLine<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日上午10:54:34 <br>
	 * @param cla
	 *            ：导入Excel的载体实体类
	 * @return 获取开始行
	 */
	private static <T> int getStartLine(Class<T> cla) {
		return cla.getAnnotation(ExcelImportConfig.class).startLine();
	}

	/**
	 * 
	 * 描述：获取Excel的载体实体类中添加ExcelImportCol注解的属性集合 <br>
	 * 
	 * @method ：getExcelImportColAnnoFields<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日上午10:59:24 <br>
	 * @param cla
	 *            ：导入Excel的载体实体类
	 * @return
	 * @throws Exception
	 */
	private static <T> List<Field> getExcelImportColAnnoFields(Class<T> cla)
			throws Exception {
		if (!excelImportColAnnoFieldsMap.containsKey(cla.getName())) {
			List<Field> list = new ArrayList<Field>();
			Field[] fields = cla.getDeclaredFields();
			for (Field f : fields) {
				if (f.isAnnotationPresent(ExcelImportCol.class)) {
					list.add(f);
				}
			}
			excelImportColAnnoFieldsMap.put(cla.getName(), list);
		}

		return excelImportColAnnoFieldsMap.get(cla.getName());
	}

	/**
	 * 
	 * 描述：验证导入Excel的载体实体类是否合法
	 * 
	 * <br>
	 * 
	 * @method ：checkValidate<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日上午11:01:59 <br>
	 * @param fileInputStream
	 *            : 导入Excel生成的文件流
	 * @param cla
	 *            :导入Excel的载体实体类
	 * @return 验证通过返回 true；验证失败返回 false
	 * @throws Exception
	 */
	private static <T> boolean checkValidate(InputStream fileInputStream,
			Class<T> cla) throws Exception {
		if (null == fileInputStream) {
			throw new Exception("导入Excel生成的文件流为空！");
		}

		if (!cla.isAnnotationPresent(ExcelImportConfig.class)) {
			throw new Exception("指定的实体类" + cla.getName()
					+ " 缺少ExcelImportConfig注解！");
		}

		if (getExcelImportColAnnoFields(cla).size() == 0) {
			throw new Exception("指定的实体类" + cla.getName()
					+ " 属性缺少ExcelImportCol注解！");
		}

		return true;
	}

	/**
	 * 
	 * 描述：判断实体对象是否为空（通过 notNullCols() 判断） <br>
	 * 
	 * @method ：validateNotNull<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日下午12:50:25 <br>
	 * @param t
	 *            : 实体对象
	 * @return
	 * @throws Exception
	 */
	private static <T> boolean validateNotNull(T t) throws Exception {
		boolean validate = false;
		int[] notNullCols = t.getClass().getAnnotation(ExcelImportConfig.class)
				.notNullCols();
		if (null == notNullCols || notNullCols.length == 0) {
			validate = true;
		} else {
			boolean[] b = new boolean[notNullCols.length];
			List<Field> list = getExcelImportColAnnoFields(t.getClass());
			PropertyDescriptor pd = null;
			Method m = null;
			Object fieldValue = null;
			int col = 0;
			for (int i = 0; i < notNullCols.length; i++) {
				for (Field f : list) {
					col = f.getAnnotation(ExcelImportCol.class).col();
					// 判断 该列值是否为空
					if (notNullCols[i] == col) {
						pd = new PropertyDescriptor(f.getName(), t.getClass());
						m = pd.getReadMethod();
						fieldValue = m.invoke(t);
						if (null == fieldValue
								|| fieldValue.toString().equals("")) {
							b[i] = false;
						} else {
							b[i] = true;
						}
						break;
					}
				}
			}

			for (int i = 0; i < b.length; i++) {
				validate = validate || b[i];
			}
		}
		return validate;
	}
}
