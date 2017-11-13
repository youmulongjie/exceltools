exceltools
================================
POI 版本 3.12

## Excel 数据导入Bean

- 配置导入注解类 ExcelImportConfig（导入的Bean必须添加该注解）
- 配置导入属性注解类 ExcelImportConfig（导入的Bean 至少有一个属性添加该注解）
- 导入工具类 ExcelToolsImport
	- 调用方法一：获取Excel的载体实体类集合，默认导入Excel的sheet索引值为0：
	```java
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
			Class<T> cla) throws Exception
	```
	- 调用方法二：获取Excel指定的sheet索引的载体实体类集合：
	```java
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
			Class<T> cla, int sheetIndex) throws Exception
	```
	- 调用方法三：获取Excel（包含多个sheet）的载体实体类集合：
	```java
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
			throws Exception
	```

## Bean 数据导出Excel

- 配置Excel 导出注解类 ExcelExportConfig（导出的Bean必须添加该注解）
- 配置Excel 导出属性注解类 ExcelExportCol（导出的Bean 至少有一个属性添加该注解）
- 导出工具类 ExcelToolsExport
	- 调用方法一：将List对象集合 转化为 byte 数组，生成一张sheet的byte数组源：
	```java
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
			throws Exception
	```
	- 调用方法二：将Map对象集合 转化为 byte 数组，同一Excel生成多张sheet的byte数组：
	```java
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
			throws Exception
	```

## Andy.wang

<img src="doc/594580820.jpg" width="15%" alt="Andy.wang的QQ"/>


