package io.github.zhx666666.word;

import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.TypeReference;
import com.alibaba.fastjson.parser.Feature;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.Serializable;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @Author: zhaohaoxin
 * @Date: 2023-10-10-15:22
 */
public class WordCreateUtil {



    public static String createWord(String str,String path,String fileName,String parentId) {




        List<TreeView> viewList = JSONObject.parseObject(str, new TypeReference<List<TreeView>>() {
        }, Feature.OrderedField);

        // 组装目录树，每个目录里面放置的是富文本，其中富文本存在于directoryBody字段中，目录名称存在于directoryName字段中，然后组装出来的目录结构树如下图
        List<TreeView> treeList = createTreeViewTree(parentId, viewList);


        try {

            // 创建空的word文档作为导出位置，其中word文档名称是“坦克.docx”

            // 参数判断
            if (StringUtils.isBlank(fileName)) {
                return null;
            }
            // 创建临时目录
            //String tmpdirPath = System.getProperty("java.io.tmpdir");
            String dateStr = new SimpleDateFormat("yyyyMMdd").format(new Date());
            String dateStr2 = new SimpleDateFormat("HHmmss").format(new Date());
            fileName = dateStr2+"_"+fileName + ".docx";
            //String uuId = UUID.randomUUID().toString().replaceAll("-", "");
            String[] params = {path, dateStr};
            // 结果例如：C:\Users\Administrator\AppData\Local\Temp\20220630\0f774122c38b423793cc1c121611c142
            String fileUrl = StringUtils.join( params, File.separator);
            if (!new File(fileUrl).exists()) {
                new File(fileUrl).mkdirs();
            }
            // 创建临时文件
            File file = new File(fileUrl, fileName);
            file.createNewFile();


            // 第二个参数为true代表开启公文格式
            XWPFDocument document = getXWPFDocument(treeList, true);
            WordUtil.generateDocxFile(document, file);
            return path+dateStr+"/"+fileName;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;

    }

    /**
     * 创建目录树
     *
     * @param
     * @return
     * @author 明快de玄米61
     * @date 2022/10/12 0:02
     **/
    private static List<TreeView> createTreeViewTree(String parentId, List<TreeView> viewList) {
        List<TreeView> treeList = new ArrayList<TreeView>();

        for (TreeView view : viewList) {
            if (parentId.equals(view.getParentId())) {
                view.setChildren(createTreeViewTree(view.getId(), viewList));
                treeList.add(view);
            }
        }
        return treeList;
    }

    /**
     * 获取XWPFDocument
     *
     * @param treeList    目录树
     * @param odfDealFlag 是否开启公文格式
     * @return
     * @author 明快de玄米61
     * @date 2022/10/12 0:01
     **/
    public static XWPFDocument getXWPFDocument(List<TreeView> treeList, boolean odfDealFlag) {
        // 初始化XWPFDocument
        XWPFDocument document = WordUtil.initXWPFDocument();

        // 组装模板目录列表
        dealDirectoryViewTree(document, treeList, odfDealFlag);

        return document;
    }

    private static void dealDirectoryViewTree(XWPFDocument document, List<TreeView> treeList, boolean odfDealFlag) {
        for (int i = 0; i < treeList.size(); i++) {
            TreeView view = treeList.get(i);
            String directoryName = view.getDirectoryName();
            Integer type = view.getType();
            if (StringUtils.isEmpty(directoryName)) {
                directoryName = " ";
            }

            // 处理中间大标题
            if (type == 0) {
                WordUtil.dealDocxTitle(document, directoryName);
            }
            // 处理正文标题
            else {
                WordUtil.dealHeading(document, type, i + 1, directoryName, odfDealFlag);
            }

            // 处理正文
            WordUtil.dealHtmlContent(document, view.getDirectoryBody());

            // 处理子级列表
            if (view.getChildren() != null && view.getChildren().size() > 0) {
                dealDirectoryViewTree(document, view.getChildren(), odfDealFlag);
            }
        }
    }
}

class TreeView implements Serializable {

    private static final long serialVersionUID = 1L;

    // 分类id
    private String id;

    // 目录名称
    private String directoryName;

    // 目录详情
    private String directoryBody;

    // 目录详情数据
    private String bodyData;

    // 父节点id
    private String parentId;


    // 节点类型
    private Integer type;


    // 排序号
    private Integer seq;

    // 句子
    private String sentences;

    // 关系图数据
    private String statisticalBody;

    // 子级集合
    private List<TreeView> children;


    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getDirectoryName() {
        return directoryName;
    }

    public void setDirectoryName(String directoryName) {
        this.directoryName = directoryName;
    }

    public String getDirectoryBody() {
        return directoryBody;
    }

    public void setDirectoryBody(String directoryBody) {
        this.directoryBody = directoryBody;
    }

    public String getBodyData() {
        return bodyData;
    }

    public void setBodyData(String bodyData) {
        this.bodyData = bodyData;
    }

    public String getParentId() {
        return parentId;
    }

    public void setParentId(String parentId) {
        this.parentId = parentId;
    }

    public Integer getType() {
        return type;
    }

    public void setType(Integer type) {
        this.type = type;
    }

    public Integer getSeq() {
        return seq;
    }

    public void setSeq(Integer seq) {
        this.seq = seq;
    }

    public String getSentences() {
        return sentences;
    }

    public void setSentences(String sentences) {
        this.sentences = sentences;
    }

    public String getStatisticalBody() {
        return statisticalBody;
    }

    public void setStatisticalBody(String statisticalBody) {
        this.statisticalBody = statisticalBody;
    }

    public List<TreeView> getChildren() {
        return children;
    }

    public void setChildren(List<TreeView> children) {
        this.children = children;
    }
}
