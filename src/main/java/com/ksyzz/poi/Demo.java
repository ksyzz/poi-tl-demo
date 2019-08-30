package com.ksyzz.poi;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.render.Render;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * @author fengqian
 * @since <pre>2019/08/30</pre>
 */
public class Demo {
    public static void main(String[] args) throws Exception{

        String path = "1.docx";
        InputStream templateFile = Demo.class.getClassLoader().getResourceAsStream(path);
        Map map = new HashMap();
        map.put("pic", new PictureRenderData(120, 80, ".png", Demo.class.getClassLoader().getResourceAsStream("1.png")));


        // 将数据整合到模板中去
        Configure.ConfigureBuilder builder = Configure.newBuilder();
        builder.supportGrammerRegexForAll();
        builder.addPlugin('@', new MyPictureRenderPolicy());
        XWPFTemplate template = XWPFTemplate.compile(templateFile, builder.build()).render(map);

        String docPath = "C:\\Users\\csdc01\\Desktop\\out.docx";
        FileOutputStream outputStream1 = new FileOutputStream(docPath);
        template.write(outputStream1);
        outputStream1.flush();
        outputStream1.close();
    }
}
