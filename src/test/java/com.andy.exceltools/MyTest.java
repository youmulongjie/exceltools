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
