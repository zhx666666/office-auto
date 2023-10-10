package io.github.zhx666666.word;

import org.apache.http.client.config.RequestConfig;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.utils.URIBuilder;
import org.apache.http.impl.client.CloseableHttpClient;

import java.io.*;
import java.util.UUID;

public class ImageUtil {

    public static File getImgFile(String imgUrl) {
        // 创建图片对象
        File imgFile = FileUtil.createTempFile(UUID.randomUUID().toString().replaceAll("-", "") + ".jpg");
        // 创建client对象
        CloseableHttpClient client = null;
        try {
            client = new DefaultSSLUtils();
        } catch (Exception e) {
            e.printStackTrace();
        }
        // 创建response对象
        CloseableHttpResponse response = null;
        // 获取输入流
        InputStream inputStream = null;
        // 文件输出流
        FileOutputStream out = null;
        try {
            // 构造一个URL对象
            URIBuilder uriBuilder = new URIBuilder(imgUrl);
            // 创建http对象
            HttpGet httpGet = new HttpGet(uriBuilder.build());
            // 处理config设置
            RequestConfig requestConfig = RequestConfig.custom().setConnectTimeout(10000).setConnectionRequestTimeout(10000).setSocketTimeout(10000).build();
            httpGet.setConfig(requestConfig);
            // 执行请求
            response = client.execute(httpGet);
            // 获取输入流
            inputStream = response.getEntity().getContent();
            // 以流的方式输出图片
            out = new FileOutputStream(imgFile);
            byte[] arr = new byte[1024];
            int len = 0;
            while ((len = inputStream.read(arr)) != -1) {
                out.write(arr, 0, len);
            }
            out.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // 回收资源
            close(client, response, inputStream, out);
        }
        return imgFile;
    }

    /**
     * 关闭资源
     *
     * @param closeables 资源列表
     **/
    private static void close(Closeable... closeables) {
        for (Closeable closeable : closeables) {
            if (closeable != null) {
                try {
                    closeable.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}

