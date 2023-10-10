package io.github.zhx666666.word;


/**
 * @Author: zhaohaoxin
 * @Date: 2023-10-10-11:15
 */
public class WordTextParam {
    private String directoryBody;//内容
    private String directoryName;//标题
    private String id;//id
    private String parentId;//父级id
    private String rId;//
    private String seq;//
    private String type;//

    public String getDirectoryBody() {
        return directoryBody;
    }

    public void setDirectoryBody(String directoryBody) {
        this.directoryBody = directoryBody;
    }

    public String getDirectoryName() {
        return directoryName;
    }

    public void setDirectoryName(String directoryName) {
        this.directoryName = directoryName;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getParentId() {
        return parentId;
    }

    public void setParentId(String parentId) {
        this.parentId = parentId;
    }

    public String getrId() {
        return rId;
    }

    public void setrId(String rId) {
        this.rId = rId;
    }

    public String getSeq() {
        return seq;
    }

    public void setSeq(String seq) {
        this.seq = seq;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }
}
