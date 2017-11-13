exceltools
================================
POI 版本 3.12

## 快速开始

- Excel 导入：

```java
//初始化client,apikey作为所有请求的默认值(可以为空)
YunpianClient clnt = new YunpianClient("apikey").init();

//修改账户信息API
Map<String, String> param = clnt.newParam(2);
param.put(YunpianClient.MOBILE, "18616020***");
param.put(YunpianClient.TEXT, "【云片网】您的验证码是1234");
Result<SmsSingleSend> r = clnt.sms().single_send(param);
//获取返回结果，返回码:r.getCode(),返回码描述:r.getMsg(),API结果:r.getData(),其他说明:r.getDetail(),调用异常:r.getThrowable()

//账户:clnt.user().* 签名:clnt.sign().* 模版:clnt.tpl().* 短信:clnt.sms().* 语音:clnt.voice().* 流量:clnt.flow().* 隐私通话:clnt.call().*

//最后释放client
clnt.close() 
```

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

SDK开源QQ群

<img src="doc/sdk_qq.jpeg" width="15%" alt="SDK开源QQ群"/>

## 文档链接
- [api文档](https://www.yunpian.com/api2.0/guide.html)

