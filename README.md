exceltools
================================
POI 版本 3.12，Web项目 导入导出Excel工具类，简单调用方法即可。

## Excel 数据导入Bean

- （前提：根据上传的文件 获取输入流 InputStream）
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
- （后话：获取 byte[] 数组后，将byte数组，转化为 文件响应对象ResponseEntity<byte[]>，返回给前端即可）

## 使用可参考测试类 MyTest.java
```java
package com.andy.exceltools;

import com.andy.execltools.exports.ExcelToolsExport;
import com.andy.execltools.imports.ExcelToolsImport;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Description: 测试类
 * Author: Andy.wang
 * Date: 2019/10/9 17:44
 */
public class MyTest {
    public static void main(String[] args) {
        importTest();

        exportTest();
    }

    // 测试 导入
    private static void importTest() {
        InputStream in = Thread.currentThread().getContextClassLoader().getResourceAsStream("import1.xlsx");

        try {
            List<People> list = ExcelToolsImport.excelImport(in, People.class);

            for (int i = 0; i < list.size(); i++) {
                System.out.println(list.get(i));
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // 测试 导出
    private static void exportTest() {
        List<People> list = new ArrayList<People>(3);

        People people = new People();
        people.setId(1001);
        people.setName("Andy");
        people.setBirthday(new Date());
        people.setFlag(1);
        list.add(people);

        people = new People();
        people.setId(1002);
        people.setName("星星");
        people.setBirthday(new Date());
        people.setFlag(2);
        list.add(people);

        people = new People();
        people.setId(1003);
        people.setName("Sun");
        people.setBirthday(new Date());
        people.setFlag(3);
        list.add(people);

        try {
            byte[] bytes = ExcelToolsExport.createExcelExport(list, "学生名单");

            // 1、获取 byte 数组后，将其 转化为 文件响应对象 ResponseEntity<byte[]> （spring 环境下的项目）
            // 2、将 文件响应对象 download 返回
            // 3、需要在 spring 配置文件中配置
            /**
             *     <!--下载 ByteArray处理器 -->
             *     <bean id="byteArrayConverter"
             *           class="org.springframework.http.converter.ByteArrayHttpMessageConverter">
             *         <property name="supportedMediaTypes">
             *             <list>
             *                 <value>text/plain;charset=UTF-8</value>
             *             </list>
             *         </property>
             *     </bean>
             */

            // ResponseEntity<byte[]> download = download(bytes, "学生名单");
            // return download;
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    /**
     * <p>
     * 描述：单文件下载
     * </p>
     *
     * @return
     * @throws IOException
     */
    /*public static ResponseEntity<byte[]> download(byte[] context, String fileName) throws IOException {
        HttpServletRequest request = ThreadContextHolder.getHttpRequest();
        String agent = request.getHeader("USER-AGENT");
        String codedfilename = null;

        if (null != agent && -1 != agent.indexOf("MSIE") || null != agent
                && -1 != agent.indexOf("Trident")) {
            // ie
            codedfilename = java.net.URLEncoder.encode(fileName, "UTF8");
        } else if (null != agent && -1 != agent.indexOf("Mozilla")) {
            // 火狐,chrome等
            codedfilename = new String(fileName.getBytes("UTF-8"), "iso-8859-1");
        }
        HttpHeaders headers = new HttpHeaders();
        headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename="
                + codedfilename);
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        return new ResponseEntity<>(context, headers, HttpStatus.OK);
    }*/
}

```

## Andy.wang

<img src="doc/594580820.jpg" width="15%" alt="Andy.wang的QQ"/>


