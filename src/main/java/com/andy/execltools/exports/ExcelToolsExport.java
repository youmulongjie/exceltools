/**
 * Excel 导出 包 <br>
 * @author ：wanglongjie<br>
 * @createDate ：2015年12月2日下午2:20:18<br>
 */
package com.andy.execltools.exports;

import java.beans.PropertyDescriptor;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.andy.execltools.exports.annotation.ExcelExportCol;
import com.andy.execltools.exports.annotation.ExcelExportConfig;

/**
 * 描述：Excel 导出工具类 <br>
 * <br>
 * 方法一：createExcelExport(List, String) : 将List集合 转化为 byte 数组，生成一张sheet的byte数组源<br>
 * 方法二：createExcelExport(Map) : 将Map（Map的key为sheet的名称，value为填充的Bean对象集合）集合 转化为
 * byte 数组，同一Excel生成多张sheet的byte数组源<br>
 * 
 * <br>
 * <p>
 * <blockquote>
 * 
 * <pre>
 * 更新日志：
 * 1、2017-11-09 Andy.wang 更新：
 *  1）ExcelExportCol注解类添加dateFormat()方法，默认值为“yyyy-mm-dd”；
 *  2）优化内存方式，将内存数据刷新到硬盘，解决大数据量时造成的内存溢出问题。
 * </pre>
 * 
 * </blockquote>
 * <p>
 * 
 * @file ：ExcelToolsExport.java<br>
 * @author ：wanglongjie<br>
 * @createDate ：2015年12月2日下午2:20:18<br>
 * 
 */
public class ExcelToolsExport {
	/**
	 * 数字格式化
	 */
	private static final String NUMERIC_FORMAT = "#############0.00######";
	/**
	 * 输出流的缓冲大小
	 */
	private static final int BUFFER_SIZE = 4096 * 10;
	/**
	 * POI 内存中缓存记录行数（此版本重大更新，解决大数据量时内存溢出问题）
	 */
	private static final int ROWACCESS = 100;

	/**
	 * 缓存每个类中带有 ExcelExportCol 注解的属性List列表，提高效率<br>
	 * 
	 * <pre>
	 *  key：注解 ExcelExportConfig 的Class类的全类名称
	 *  value：该Class类中 带有ExcelExportCol 注解的属性List列表
	 * </pre>
	 */
	private static Map<String, List<Field>> excelExportColAnnoFieldsMap = Collections
			.synchronizedMap(new HashMap<String, List<Field>>());

	/**
	 * 
	 * 描述：将List对象集合 转化为 byte 数组，生成一张sheet的byte数组源 <br>
	 * 
	 * @method ：createExcelExport<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日下午4:02:35 <br>
	 * @param list
	 *            : 对象集合数据源
	 * @param sheetName
	 *            ：要生成的 sheet 的名称
	 * @return byte数组
	 * @throws Exception
	 */
	public static <T> byte[] createExcelExport(List<T> list, String sheetName)
			throws Exception {
		checkValidate(list);

		Workbook wb = createWorkbook();
		sheetName = (null == sheetName || sheetName.equals("")) ? "sheet1"
				: sheetName;
		SXSSFSheet sheet = (SXSSFSheet) wb.createSheet(sheetName);
		setExcelHeader(sheet, list.get(0));
		setExcelLines(sheet, list, wb);

		return getByteFormWb(wb);
	}

	/**
	 * 
	 * 描述：将Map对象集合 转化为 byte 数组，同一Excel生成多张sheet的byte数组源 <br>
	 * 
	 * @method ：createExcelExport<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日下午4:12:03 <br>
	 * @param map
	 *            : 封装的sheet的数据源集合（Map的key为sheet的名称，value为填充的Bean对象集合）
	 * @return byte数组
	 * @throws Exception
	 */
	public static byte[] createExcelExport(Map<String, List<?>> map)
			throws Exception {
		if (null != map && map.size() > 0) {
			Workbook wb = createWorkbook();

			Iterator<String> it = map.keySet().iterator();
			SXSSFSheet sheet = null;// 生成的 sheet 对象
			String sheetName = null;// sheet 名称
			List<?> list = null; // sheet 数据源
			while (it.hasNext()) {
				sheetName = (String) it.next();
				list = map.get(sheetName);

				checkValidate(list);

				sheet = (SXSSFSheet) wb.createSheet(sheetName);
				setExcelHeader(sheet, list.get(0));
				setExcelLines(sheet, list, wb);
			}

			return getByteFormWb(wb);
		}
		return null;
	}

	/**
	 * 
	 * 描述：设置 生成Excel的内容记录行 <br>
	 * 
	 * @method ：setExcelLines<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日下午3:23:25 <br>
	 * @param sheet
	 *            ：创建的 Sheet对象
	 * @param list
	 *            ：对象集合数据源
	 * @param wb
	 *            ：创建的WorkBook 工作薄对象
	 * @throws Exception
	 */
	private static <T> void setExcelLines(SXSSFSheet sheet, List<T> list,
			Workbook wb) throws Exception {
		int lineStartRow = getLineStartRow(list.get(0).getClass());
		Row row = null;
		for (int i = 0; i < list.size(); i++) {
			row = sheet.createRow(lineStartRow);
			obj2Cell(row, list.get(i), wb);
			lineStartRow++;
			// 每当行数达到设置的ROWACCESS值 就刷新数据到硬盘,以清理内存
			if (i % ROWACCESS == 0) {
				sheet.flushRows();
			}
		}
	}

	/**
	 * 
	 * 描述：设置 生成Excel的标题头行 <br>
	 * 
	 * @method ：setExcelHeader<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日下午3:20:11 <br>
	 * @param sheet
	 *            ：Sheet 对象
	 * @param t
	 *            ：Excel 的载体实体类对象
	 * @throws Exception
	 */
	private static <T> void setExcelHeader(Sheet sheet, T t) throws Exception {
		int headRow = getHeaderRow(t.getClass());
		Row row = sheet.createRow(headRow);

		List<Field> list = getExcelExportColAnnoFields(t.getClass());
		ExcelExportCol excelExportCol = null;
		int cols = 0;// 标题列索引
		String colHeaderDesc = null;// 标题头说明
		for (Field f : list) {
			excelExportCol = f.getAnnotation(ExcelExportCol.class);
			cols = excelExportCol.cols();
			colHeaderDesc = excelExportCol.colHeaderDesc();
			row.createCell(cols).setCellValue(colHeaderDesc);
		}
	}

	/**
	 * 
	 * 描述：填充 Excel 数据行 <br>
	 * 
	 * @method ：obj2Cell<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日下午3:59:57 <br>
	 * @param row
	 *            : 创建的Row 行对象
	 * @param t
	 *            ：对象数据源
	 * @param wb
	 *            ：创建的Workbook 工作薄对象
	 * @throws Exception
	 */
	private static <T> void obj2Cell(Row row, T t, Workbook wb)
			throws Exception {
		List<Field> list = getExcelExportColAnnoFields(t.getClass());

		ExcelExportCol excelExportCol = null;
		Cell cell = null; // 单元格
		int cols = 0;// 单元格 列索引
		Object value = null;// 单元格内容（反射获取的属性值）

		PropertyDescriptor pd = null;
		Method m = null;
		for (Field f : list) {
			excelExportCol = f.getAnnotation(ExcelExportCol.class);
			cols = excelExportCol.cols();

			// 反射获取对象的属性值
			pd = new PropertyDescriptor(f.getName(), t.getClass());
			m = pd.getReadMethod();
			value = m.invoke(t);

			// 处理 映射配置
			String mapper = excelExportCol.mapper();
			// 映射配置后的结果
			String mapperResult = null;
			// 日期格式
			String dateFormat = excelExportCol.dateFormat();

			if (null != value) {
				if (null != mapper && !mapper.trim().equals("")) {
					mapper = mapper.replaceAll(" ", "").replaceAll("，", ",");
					if (mapper.contains("=")) {
						if (mapper.contains(",") || mapper.contains("，")) {
							String[] groups = mapper.split(",");
							for (int i = 0; i < groups.length; i++) {
								String group[] = groups[i].split("=");
								if (group[0].equals(value.toString())) {
									mapperResult = group[1];
									break;
								}
							}
						} else {
							String keyValue[] = mapper.split("=");
							if (keyValue[0].trim().equals(value.toString())) {
								mapperResult = keyValue[1].trim();
							}
						}
					}
				}
			}

			if (null != mapperResult) {
				cell = row.createCell(cols);
				fillCell(mapperResult, cell, wb, dateFormat);
				mapperResult = null;
			} else {
				cell = row.createCell(cols);
				fillCell(value, cell, wb, dateFormat);
			}

		}
	}

	/**
	 * 
	 * 描述：填充 单元格 <br>
	 * 
	 * @method ：fillCell<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日下午3:57:40 <br>
	 * @param value
	 *            ：单元格 内容
	 * @param cell
	 *            ：要填充的单元格
	 * @param wb
	 *            ： 创建的WorkBook 工作薄对象
	 * @param dateFormat
	 *            ：日期格式
	 */
	private static void fillCell(Object value, Cell cell, Workbook wb,
			String dateFormat) {
		if (null == value) {
			cell.setCellValue("");
			return;
		}

		if (value instanceof java.util.Date) {
			java.util.Date d = (java.util.Date) value;
			DataFormat format = wb.createDataFormat();
			// 日期格式化
			CellStyle cellStyle = wb.createCellStyle();
			cellStyle.setDataFormat(format.getFormat(dateFormat));

			cell.setCellStyle(cellStyle);
			cell.setCellValue(d);
			return;
		}

		if (value instanceof java.sql.Date) {
			java.sql.Date d = (java.sql.Date) value;
			cell.setCellValue(d);
			return;
		}

		if (value instanceof Timestamp) {
			Timestamp ts = (Timestamp) value;
			cell.setCellValue(ts);
			return;
		}

		if (value instanceof BigDecimal) {
			BigDecimal b = (BigDecimal) value;
			cell.setCellValue(b.doubleValue());
			return;
		}

		if (value instanceof Double) {
			Double d = (Double) value;
			DataFormat format = wb.createDataFormat();
			// 数字格式化
			CellStyle cellStyle = wb.createCellStyle();
			cellStyle.setDataFormat(format.getFormat(NUMERIC_FORMAT));

			cell.setCellStyle(cellStyle);
			cell.setCellValue(d);
			return;
		}

		if (value instanceof Float) {
			Float f = (Float) value;
			DataFormat format = wb.createDataFormat();
			// 数字格式化
			CellStyle cellStyle = wb.createCellStyle();
			cellStyle.setDataFormat(format.getFormat(NUMERIC_FORMAT));

			cell.setCellStyle(cellStyle);
			cell.setCellValue(f);
			return;
		}

		if (value instanceof Long) {
			Long l = (Long) value;
			cell.setCellValue(l);
			return;
		}

		if (value instanceof Integer) {
			Integer i = (Integer) value;
			cell.setCellValue(i);
			return;
		}

		if (value instanceof Boolean) {
			Boolean b = (Boolean) value;
			cell.setCellValue(b);
			return;
		}

		if (value instanceof String) {
			String s = (String) value;
			cell.setCellValue(s);
			return;
		}

	}

	/**
	 * 
	 * 描述：创建 WorkBook 工作薄对象 <br>
	 * 
	 * @method ：createWorkbook<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日下午3:01:03 <br>
	 * @param flag
	 *            ：true:xls(1997-2007) false:xlsx(2007以上)
	 * @return WorkBook 工作薄对象
	 */
	private static Workbook createWorkbook() {
		Workbook wb = new SXSSFWorkbook(ROWACCESS);
		return wb;
	}

	/**
	 * 
	 * 描述：将创建的 Workbook 工作薄对象转化为byte输出流 <br>
	 * 
	 * @method ：getByteFormWb<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日下午2:59:10 <br>
	 * @param wb
	 *            : 创建的 Workbook 工作薄对象
	 * @return byte输出流
	 * @throws Exception
	 */
	private static byte[] getByteFormWb(Workbook wb) throws Exception {
		if (null != wb) {
			ByteArrayOutputStream byStream = new ByteArrayOutputStream(
					BUFFER_SIZE);
			wb.write(byStream);
			return byStream.toByteArray();
		}
		return null;
	}

	/**
	 * 
	 * 描述： 读取生成Excel文件的 标题所在的行数 <br>
	 * 
	 * @method ：getHeaderRow<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日下午2:35:20 <br>
	 * @param cla
	 *            ：Excel 的载体实体类
	 * @return
	 * @throws Exception
	 */
	private static <T> int getHeaderRow(Class<T> cla) throws Exception {
		return cla.getAnnotation(ExcelExportConfig.class).headerRow();
	}

	/**
	 * 
	 * 描述：读取生成Excel文件的 对象记录所在的行数 <br>
	 * 
	 * @method ：getLineStartRow<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日下午2:36:15 <br>
	 * @param cla
	 *            ：Excel 的载体实体类
	 * @return
	 * @throws Exception
	 */
	private static <T> int getLineStartRow(Class<T> cla) throws Exception {
		return cla.getAnnotation(ExcelExportConfig.class).lineStartRow();
	}

	/**
	 * 
	 * 描述： 获取Excel的载体实体类中添加ExcelExportCol注解的属性集合<br>
	 * 
	 * @method ：getExcelExportColAnnoFields<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日下午2:39:16 <br>
	 * @param cla
	 *            ：Excel 的载体实体类
	 * @return
	 * @throws java.lang.Exception
	 */
	private final static <T> List<Field> getExcelExportColAnnoFields(
			Class<T> cla) throws Exception {
		if (!excelExportColAnnoFieldsMap.containsKey(cla.getName())) {
			List<Field> list = new ArrayList<Field>();
			Field[] fields = cla.getDeclaredFields();
			for (Field f : fields) {
				if (f.isAnnotationPresent(ExcelExportCol.class)) {
					list.add(f);
				}
			}
			excelExportColAnnoFieldsMap.put(cla.getName(), list);
		}
		return excelExportColAnnoFieldsMap.get(cla.getName());
	}

	/**
	 * 
	 * 描述：验证导出Excel的载体实体类是否合法 <br>
	 * 
	 * @method ：checkValidate<br>
	 * @author ：wanglongjie<br>
	 * @createDate ：2015年12月2日下午2:54:59 <br>
	 * @param list
	 *            ：对象集合数据源
	 * @return
	 * @throws Exception
	 */
	private static boolean checkValidate(List<?> list) throws Exception {
		if (null == list || list.size() == 0) {
			throw new Exception("指定的对象集合数据源为null，或者长度等于0！");
		}
		Class<?> cla = list.get(0).getClass();

		if (!cla.isAnnotationPresent(ExcelExportConfig.class)) {
			throw new Exception("指定的实体类" + list.get(0).getClass().getName()
					+ " 缺少ExcelExportConfig注解！");
		}

		int headerRow = getHeaderRow(cla);
		int lineStartRow = getLineStartRow(cla);

		if (headerRow >= lineStartRow) {
			throw new Exception("指定的实体类" + cla.getName()
					+ " 设置的标题头行应该小于内容记录开始行！");
		}

		if (getExcelExportColAnnoFields(cla).size() == 0) {
			throw new Exception("指定的实体类" + cla.getName()
					+ " 属性缺少ExcelExportCol注解！");
		}

		return true;
	}
}
