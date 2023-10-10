package io.github.zhx666666.word;


import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.select.Elements;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class WordUtil {

    // 单位 start
    private static final int PER_LINE = 100;
    //每个字符的单位长度
    private static final int PER_CHART = 100;
    // 厘米转磅
    private static final double CM_2_POUND = 28.35;
    //每一磅的单位长度
    private static final int PER_POUND = 20;
    //行距单位长度
    private static final int ONE_LINE = 240;
    // 单位 end

    // 纸张 start
    // A4纸的大小是 宽—21（厘米）、高—29.7（厘米），然后厘米转磅数在乘以20，然后计算出来的大约结果
    // 纸张宽度
    private static final double PAGE_WIDTH = 21;

    // 纸张高度
    private static final double PAGE_HEIGHT = 29.7;
    // 纸张 end

    // 页边距 start
    // 上页边距(单位：cm)
    private static final double TOP = 3.7;
    // 上页边距(单位：cm)
    private static final double BOTTOM = 3.5;
    // 上页边距(单位：cm)
    private static final double LEFT = 2.8;
    // 上页边距(单位：cm)
    private static final double RIGHT = 2.6;
    // 页边距 end

    // 页码样式 start
    // 字体
    private static final String PAGE_FONT_FAMILY = "宋体";
    // 字号
    private static final Integer PAGE_FONT_SIZE = 14;
    // 页码样式 end

    // 正文样式 start
    // 字体
    private static final String PARA_FONT_FAMILY = "仿宋_GB2312";
    // 字号
    private static final Integer PARA_FONT_SIZE = 16;
    // 行距(单位：磅)
    private static final double PARA_ROW_SPACE = 28.95;
    // 正文样式 end

    // 标题最大级别
    private static final Integer MAX_HEADING_LEVEL = 9;

    // 最大正文图片宽度
    private static final Integer MAX_PAGE_IMG_WIDTH = 350;

    // 最大表格图片宽度
    private static final Integer MAX_TABLE_IMG_WIDTH = 200;

    // 表格最大宽度（单位：cm）
    private static final double TABLE_WIDTH = 16.19;

    // 单元格边距（单位：磅）
    private static final double CELL_MARGIN = 5.67;

    /** 标题样式集合 **/
    private static List<HeadingStyle> headingStyleList = new ArrayList<>();

    /** 定义标题格式 **/
    static {
        // 处理前四级标题样式
        HeadingStyle one = new HeadingStyle(16, "黑体", true);
        HeadingStyle two = new HeadingStyle(16, "楷体", true);
        HeadingStyle three = new HeadingStyle(16, "仿宋", true);
        HeadingStyle four = new HeadingStyle(16, "仿宋", true);
        headingStyleList.add(one);
        headingStyleList.add(two);
        headingStyleList.add(three);
        headingStyleList.add(four);

        // 处理四级以下的标题样式（注意：由于公文格式中未指定样式，所以采用默认格式）
        for (Integer i = 5; i <= MAX_HEADING_LEVEL; i++) {
            HeadingStyle headingStyle = new HeadingStyle(16, "仿宋", true);
            headingStyleList.add(headingStyle);
        }
    }

    /**
     * 生成docx文档文件
     * @author 明快de玄米61
     * @date   2022/9/13 16:37
     * @param  document XWPFDocument对象
     * @param  docxFile docx文档文件
     * @return docx文档文件
     **/
    public static void generateDocxFile(XWPFDocument document, File docxFile) {
        try {
            document.write(new FileOutputStream(docxFile));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 初始化XWPFDocument
     * @author 明快de玄米61
     * @date   2022/9/13 10:40
     **/
    public static XWPFDocument initXWPFDocument() {
        XWPFDocument document = new XWPFDocument();

        // 多处使用，必须统一设置，不然会出现word和wps效果不一致的情况
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();

        // 初始化纸张大小，必须放在页边距设置前面，不然页边距就会出现问题
        initPageSize(sectPr, PAGE_WIDTH, PAGE_HEIGHT);

        // 初始化页边距
        initPageMargin(sectPr);

        // 初始化页脚
        initFooter(document, sectPr, PAGE_FONT_FAMILY, PAGE_FONT_SIZE, "000000", null, " ");

        // 初始化标题级别
        initHeadingStyle(document);

        return document;
    }

    /**
     * 初始化页面大小
     *
     * @author 明快de玄米61
     * @date   2023/7/12 9:57
     * @param  sectPr CTSectPr对象
     * @param  pageWidth 页面宽度
     * @param  pageHeight 页面高度
     **/
    private static void initPageSize(CTSectPr sectPr, double pageWidth, double pageHeight) {
        CTPageSz pgsz = sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz();
        BigInteger w = BigInteger.valueOf(Math.round(pageWidth * CM_2_POUND * PER_POUND));
        pgsz.setW(w);
        BigInteger h = BigInteger.valueOf(Math.round(pageHeight * CM_2_POUND * PER_POUND));
        pgsz.setH(h);
    }

    /**
     * 处理文章中间标题
     * @author 明快de玄米61
     * @date   2022/9/13 15:23
     * @param  document XWPFDocument对象
     * @param  text 标题文本
     **/
    public static void dealDocxTitle(XWPFDocument document, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setSpacingAfterLines((int) (100));
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setVerticalAlignment(TextAlignment.TOP);
        XWPFRun xwpfRun = paragraph.createRun();
        xwpfRun.setBold(true);
        xwpfRun.setFontSize((int) (22));
        xwpfRun.setFontFamily("方正小标宋简体");
        xwpfRun.setText(text);
    }

    /**
     * 处理标题
     * @author 明快de玄米61
     * @date   2022/9/13 15:18
     * @param  document XWPFDocument对象
     * @param  level 级别
     * @param  sort 同级排序
     * @param  text 标题文本
     * @param  odfDealFlag 公文格式转换标志
     **/
    public static void dealHeading(XWPFDocument document, int level, int sort, String text, boolean odfDealFlag) {
        // 如果标题级别超标，那么使用最后一个标题的样式
        if (level > headingStyleList.size()) {
            level = headingStyleList.size();
        }
        HeadingStyle headingStyle = headingStyleList.get(level - 1);
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setStyle(getHeadingStyle(level));
        // 设置间距
        setPSpacing(paragraph, 1, 1);
        XWPFRun xwpfRun = paragraph.createRun();
        xwpfRun.setBold(headingStyle.isBold());
        xwpfRun.setFontSize((int) (headingStyle.getFontSize()));
        xwpfRun.setFontFamily(headingStyle.getFontFamily());
        xwpfRun.setText(getHeadingTextByODF(level, sort, text, odfDealFlag));
    }

    /**
     * 根据公文格式获取标题（注意：ODF：公文格式Official document format）
     * @author 明快de玄米61
     * @date   2022/9/13 15:59
     * @param  level 标题级别
     * @param  sort 同级序号
     * @param  text 原始文本
     * @return 按照公文格式处理之后的文本
     * @param  odfDealFlag 是否进行公文格式转换
     **/
    private static String getHeadingTextByODF(int level, int sort, String text, boolean odfDealFlag) {
        if (!odfDealFlag) {
            return text;
        }

        // 处理前面四级标题
        String prefix = null;
        switch (level) {
            case 1:
                prefix = ChineseNumToArabicNumUtil.arabicNumToChineseNum(sort) + "、";
                break;
            case 2:
                prefix = "（" + ChineseNumToArabicNumUtil.arabicNumToChineseNum(sort) + "）";
                break;
            case 3:
                prefix = sort + ".";
                break;
            case 4:
                prefix = "（" + sort + "）";
                break;
        }

        // 判断前四级标题是否已经包含前缀
        if (prefix != null) {
            text = text.startsWith(prefix) ? text : prefix + text;
        }

        // 四级以下标题不在处理
        return text;
    }

    /**
     * 获取标题样式标志
     * @author 明快de玄米61
     * @date   2022/9/13 15:18
     * @param  level 级别
     * @return 标题样式标志
     **/
    private static String getHeadingStyle(int level) {
        return String.valueOf(level);
    }

    /**
     * 处理正文内容
     * @author 明快de玄米61
     * @date   2022/9/13 11:45
     * @param  document
     * @return
     **/
    public static void dealHtmlContent(XWPFDocument document, String html) {
        // 判断html是否为空
        if (StringUtils.isEmpty(html)) {
            return;
        }

        // 去除特殊字符
        html = dealSpecialCharacters(html);

        // 处理不存在p标签的情况（说明：如果直接复制一句话到另外一个输入框，那么就不会生成p标签）
        List<String> extractResultList = getExtractResultList(html);

        // 处理html
        for (String content : extractResultList) {
            Pattern tablePattern = Pattern.compile("<table.*?</table>");
            Pattern hPattern = Pattern.compile("<h.*?</h[1-9]{1}>");
            Pattern imgPattern = Pattern.compile("<img.*?/>");
            Pattern aPattern = Pattern.compile("<a.*?</a>");

            // 表格(说明：复用采集1.0中处理表格的代码)
            if (tablePattern.matcher(content).find()) {
                // 处理tbody，适配采集1.0中导出表格的代码
                Pattern pattern = Pattern.compile("<tbody.*?</tbody>");
                Matcher matcher = pattern.matcher(content);
                String tbody = null;
                while (matcher.find()) {
                    tbody = matcher.group();
                }
                // 按照采集1.0中要求进行数据封装
                String table = "<table style=\"margin:0 auto; text-align:center\">" + tbody + "</table>";
                table = table.replace("<tbody", "<tbody ").replace("<td", "<td ").replace("<tr", "<tr ").replace("<th", "<th ");
                table = table.replace("<th", "<td").replace("</th", "</td");
                Document word = Jsoup.parse(table);
                // 使用采集1.0中处理表格的工具类代码
                parseTableElement(word, document);
            }
            // 标题
            else if (hPattern.matcher(content).find()) {
                XWPFParagraph paragraph = createP(document);
                // 设置对齐、缩进
                setHAttr(paragraph, content);
                // 设置行距
                setPRowSpacing(paragraph, PARA_ROW_SPACE);
                dealHText(paragraph, content);
            }
            // 图片
            else if (imgPattern.matcher(content).find()) {
                XWPFParagraph paragraph = createP(document);
                // 设置对齐、缩进
                setPAttr(paragraph, content);
                // 设置段前段后间距
                setPSpacing(paragraph, 1, 1);
                dealImg(paragraph, content, PARA_FONT_FAMILY, PARA_FONT_SIZE);
            }
            // 超链接
            else if (aPattern.matcher(content).find()) {
                XWPFParagraph paragraph = createP(document);
                // 设置对齐、缩进
                setPAttr(paragraph, content);
                // 设置行距
                setPRowSpacing(paragraph, PARA_ROW_SPACE);
                dealLink(paragraph, content, PARA_FONT_FAMILY, PARA_FONT_SIZE);
            }
            // 纯文本
            else {
                XWPFParagraph paragraph = createP(document);
                // 设置对齐、缩进
                setPAttr(paragraph, content);
                // 设置行距
                setPRowSpacing(paragraph, PARA_ROW_SPACE);
                String text = Jsoup.parse(content).text();
                dealPText(paragraph, text, PARA_FONT_FAMILY, PARA_FONT_SIZE);
            }
        }
    }

    private static String dealSpecialCharacters(String html) {
        return html.replaceAll("[\r|\n|\b]", "");
    }

    private static void dealHText(XWPFParagraph paragraph, String content) {
        // 处理字体大小
        Integer hNum = getHNum(content);
        Map<Integer, String> familyMap = new HashMap<>();
        familyMap.put(1, "黑体");
        familyMap.put(2, "楷体");
        familyMap.put(3, "仿宋");
        familyMap.put(4, "仿宋");
        familyMap.put(5, "仿宋");
        familyMap.put(6, "仿宋");
        String family = StringUtil.toStr(familyMap.get(hNum), PARA_FONT_FAMILY);
        Map<Integer, Integer> fontSizeMap = new HashMap<>();
        // delete start by 明快de玄米61 time 2022/9/13 reason 标题太大，暂时删除
//        fontSizeMap.put(1, 32);
//        fontSizeMap.put(2, 24);
//        fontSizeMap.put(3, 19);
//        fontSizeMap.put(4, 16);
//        fontSizeMap.put(5, 14);
//        fontSizeMap.put(6, 13);
        // delete end by 明快de玄米61 time 2022/9/13 reason 标题太大，暂时删除
        // add start by 明快de玄米61 time 2022/9/13 reason 标题太大，字号暂时使用16
        fontSizeMap.put(1, 16);
        fontSizeMap.put(2, 16);
        fontSizeMap.put(3, 16);
        fontSizeMap.put(4, 16);
        fontSizeMap.put(5, 16);
        fontSizeMap.put(6, 16);
        // add end by 明快de玄米61 time 2022/9/13 reason 标题太大，字号暂时使用16
        Integer fontSize = StringUtil.toInt(fontSizeMap.get(hNum), PARA_FONT_SIZE);
        // 按照链接形式处理
        dealLink(paragraph, content,  family ,fontSize);
    }

    private static Map<String, String> getFirstLabelStyle(String content, String labelName) {
        Pattern pattern = Pattern.compile(String.format("<%s.*?style=\"(.*?)\".*?>", labelName));
        Matcher matcher = pattern.matcher(content);
        String style = null;
        while (matcher.find()) {
            style = matcher.group(1).trim();
        }
        Map<String, String> attrMap = new HashMap<>();
        if (style != null) {
            String[] attrArr = style.split(";");
            for (String attr : attrArr) {
                String[] keyAndValue = attr.split(":");
                attrMap.put(keyAndValue[0].trim(), keyAndValue[1].trim());
            }
        }
        return attrMap;
    }

    private static void setPAttr(XWPFParagraph paragraph, String content) {
        Map<String, String> attrMap = getFirstLabelStyle(content, "p");
        setPStyle(paragraph, attrMap);
    }

    private static void setHAttr(XWPFParagraph paragraph, String content) {
        Map<String, String> attrMap = getFirstLabelStyle(content, "h[1-9]{1}");
        // 处理style
        setPStyle(paragraph, attrMap);
    }

    private static Integer getHNum(String content) {
        Pattern pattern = Pattern.compile("<h([1-9]{1}).*?");
        Matcher matcher = pattern.matcher(content);
        Integer num = null;
        while (matcher.find()) {
            num = Integer.valueOf(matcher.group(1));
            break;
        }
        return num;
    }

    private static void setPStyle(XWPFParagraph paragraph, Map<String, String> attrMap) {
        String align = attrMap.get("text-align");
        String indent = attrMap.get("text-indent");
        // 对齐方式
        if (align != null) {
            switch (align.toLowerCase()) {
                case "left":
                    paragraph.setAlignment(ParagraphAlignment.LEFT);
                    break;
                case "center":
                    paragraph.setAlignment(ParagraphAlignment.CENTER);
                    break;
                case "right":
                    paragraph.setAlignment(ParagraphAlignment.RIGHT);
                    break;
                case "justify":
                    paragraph.setAlignment(ParagraphAlignment.BOTH);
            }
        } else {
            paragraph.setAlignment(ParagraphAlignment.LEFT);
        }
        // 缩进
        if (indent != null) {
            if (indent.contains("em")) {
                setTextIndent(paragraph, Integer.valueOf(indent.replaceAll("em", "")));
            }
        }
    }

    private static XWPFParagraph createP(XWPFDocument document) {
        // 正文1
        XWPFParagraph paragraph = document.createParagraph();
        return paragraph;
    }

    private static void dealPImg(XWPFParagraph paragraph, String src, Integer width) {
        File imgFile = ImageUtil.getImgFile(src);
        try {
            writeImage(paragraph, imgFile.getAbsolutePath(), width);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            FileUtil.deleteParentFile(imgFile);
        }
    }

    private static void dealPText(XWPFParagraph paragraph, String text, String family, Integer fontSize) {
        XWPFRun firstRun = paragraph.createRun();
        // 设置字体和字号
        setTextFontFamilyAndFontSize(firstRun, StringUtil.toStr(family, PARA_FONT_FAMILY), StringUtil.toInt(fontSize, 16));
        // 设置文本
        firstRun.setText(reverseEscapeChar(text));
    }

    private static void dealPLink(XWPFParagraph paragraph, String html, String href, String family, Integer fontSize) {
        String name = Jsoup.parse(html).text();
        String id = paragraph
                .getDocument()
                .getPackagePart()
                .addExternalRelationship(href,
                        XWPFRelation.HYPERLINK.getRelation()).getId();
        CTHyperlink cLink = paragraph.getCTP().addNewHyperlink();
        cLink.setId(id);
        // 创建链接文本
        CTText ctText1 = CTText.Factory.newInstance();
        ctText1.setStringValue(name);
        CTR ctr = CTR.Factory.newInstance();
        CTRPr rpr = ctr.addNewRPr();
        //设置超链接样式
        CTColor color = CTColor.Factory.newInstance();
        color.setVal("0000FF");
        rpr.setColor(color);
        rpr.addNewU().setVal(STUnderline.SINGLE);
        //设置字体
        CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
        fonts.setAscii(StringUtil.toStr(family, PARA_FONT_FAMILY));
        fonts.setEastAsia(StringUtil.toStr(family, PARA_FONT_FAMILY));
        fonts.setHAnsi(StringUtil.toStr(family, PARA_FONT_FAMILY));
        //设置字体大小
        CTHpsMeasure sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
        sz.setVal(new BigInteger(StringUtil.toStr(fontSize * 2, "32")));
        ctr.setTArray(new CTText[] { ctText1 });
        // Insert the linked html into the link
        cLink.setRArray(new CTR[] { ctr });

    }

    private static void dealLink(XWPFParagraph paragraph, String html, String family, Integer fontSize) {
        List<LinkInfo> linkList=new ArrayList<>();
        Pattern pattern = Pattern.compile("(<a.*?)href=\"(.*?)\".*?>(.*)</a>");
        Matcher matcher=pattern.matcher(html);
        while(matcher.find()) {
            linkList.add(new LinkInfo(matcher.start(), matcher.end(), matcher.group(2), matcher.group(3)));
        }
        if (linkList.size() > 0) {
            for (int i = 0; i < linkList.size(); i++) {
                // 当前
                LinkInfo current = linkList.get(i);

                // 处理头部
                if (i == 0 && current.getStart() > 0) {
                    String text = Jsoup.parse(html.substring(0, current.getStart())).text();
                    dealPText(paragraph, text, family, fontSize);
                }

                // 处理自身
                dealPLink(paragraph, current.getHtml(), current.getHref(),  family, fontSize);

                // 处理中间
                if (i > 0 && i < linkList.size() - 1) {
                    // 下一个
                    LinkInfo next = linkList.get(i+1);
                    if (current.getEnd() < next.getStart()) {
                        String text = Jsoup.parse(html.substring(current.getEnd() + 1, next.getStart())).text();
                        dealPText(paragraph, text, family, fontSize);
                    }
                }

                // 处理尾部
                if (i == linkList.size() - 1 && current.getEnd() < html.length()) {
                    String text = Jsoup.parse(html.substring(current.getEnd())).text();
                    dealPText(paragraph, text, family, fontSize);
                }
            }
        } else {
            // TODO 处理段落
            String text = Jsoup.parse(html).text();
            dealPText(paragraph, text, family, fontSize);
        }
    }

    private static void dealImg(XWPFParagraph paragraph, String html, String fontFamily, int fontSize) {
        Pattern pattern = Pattern.compile("<p.*?>(.*?)(<img.*?/>)(.*?)</p>");
        Matcher matcher = pattern.matcher(html);
        while (matcher.find()) {
            dealLink(paragraph, matcher.group(1), fontFamily, fontSize);
            String src = getAttrByImg(matcher.group(2), "src");
            String width = getAttrByImg(matcher.group(2), "width");
            dealPImg(paragraph, src, StringUtils.isEmpty(width) ? null : Integer.valueOf(width));
            dealLink(paragraph, matcher.group(3), fontFamily, fontSize);
        }
    }

    private static String getAttrByImg(String html, String attrName) {
        if (html == null) {
            return null;
        }
        Document document = Jsoup.parse(html);
        Elements img = document.getElementsByTag("img");
        return img.get(0).attr(attrName);
    }

    /**
     * 将表格、标题、文本抽取出来
     * @author 明快de玄米61
     * @date   2022/9/9 9:45
     * @param  html html代码
     * @return 抽取结果列表
     **/
    private static List<String> getExtractResultList(String html) {
        List<String> result = new ArrayList<>();
        // 抽取表格
        List<TableInfo> tableList=new ArrayList<TableInfo>();
        Pattern pt=Pattern.compile("<table.*?</table>");
        Matcher mt=pt.matcher(html);
        while(mt.find()) {
            tableList.add(new TableInfo(mt.start(), mt.end(), mt.group()));
        }
        if (tableList.size() > 0) {
            for (int i = 0; i < tableList.size(); i++) {
                // 当前
                TableInfo current = tableList.get(i);

                // 处理第一次表格之前的内容
                if (i == 0 && current.getStart() > 0) {
                    dealPAndHLabel(html.substring(0, current.getStart()), result);
                }

                // 处理两个表格相连的情况
                dealTableConnect(result);

                // 处理表格中单元格空白导致转换pdf报错问题
                String currentHtml = dealTableTdBlank(current.getHtml());

                // 处理表格内容
                result.add(currentHtml);

                // 处理表格后面的内容（注意：不处理最后一个表格之后的内容）
                if (i < tableList.size() - 1) {
                    // 下一个
                    TableInfo next = tableList.get(i+1);
                    if (current.getEnd() < next.getStart()) {
                        dealPAndHLabel(html.substring(current.getEnd(), next.getStart()), result);
                    }
                }

                // 处理表格后面的内容（注意：只处理最后一个表格之后的内容）
                if (i == tableList.size() - 1 && current.getEnd() < html.length()) {
                    dealPAndHLabel(html.substring(current.getEnd()), result);
                }
            }
        } else {
            dealPAndHLabel(html, result);
        }
        return result;
    }

    private static String dealImgNotWrapByP(String html) {
        List<LabelInfo> labelInfos = new ArrayList<>();
        String regex = "<p.*?</p>";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(html);
        while (matcher.find()) {
            labelInfos.add(new LabelInfo(matcher.start(), matcher.end(), matcher.group()));
        }
        if (labelInfos.size() > 0) {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < labelInfos.size(); i++) {
                // 当前
                LabelInfo current = labelInfos.get(i);

                // 处理第一次表格之前的内容
                if (i == 0 && current.getStart() > 0) {
                    usePWrapImg(sb, html.substring(0, current.getStart()));
                }

                // 处理表格内容
                sb.append(current.getHtml());

                // 处理表格后面的内容（注意：不处理最后一个表格之后的内容）
                if (i < labelInfos.size() - 1) {
                    // 下一个
                    LabelInfo next = labelInfos.get(i + 1);
                    if (current.getEnd() < next.getStart()) {
                        usePWrapImg(sb, html.substring(current.getEnd(), next.getStart()));
                    }
                }

                // 处理表格后面的内容（注意：只处理最后一个表格之后的内容）
                if (i == labelInfos.size() - 1 && current.getEnd() < html.length()) {
                    usePWrapImg(sb, html.substring(current.getEnd()));
                }
            }
            html = sb.toString();
        } else {
            StringBuilder sb = new StringBuilder();
            usePWrapImg(sb, html);
            html = sb.toString();
        }
        return html;
    }

    private static void usePWrapImg(StringBuilder sb, String html) {
        Pattern pattern = Pattern.compile("<img.*?/>");
        String s = pattern.matcher(html).replaceAll("<p style=\"text-align: center;\">$0</p>");
        sb.append(s);
    }

    /**
     * 使用p标签包括其他标签
     * @author 明快de玄米61
     * @date   2022/9/16 1:13
     * @param  html html标签
     * @return 处理之后的html标签
     **/
    private static String usePWrapLabel(String html) {
        boolean empty = StringUtils.isEmpty(Jsoup.parse(html).text());
        if (!empty) {
            return "<p style=\"text-indent: 2em\">" + html + "</p>";
        }
        return null;
    }

    private static String dealTableTdBlank(String html) {
        Pattern pattern = Pattern.compile("(<td.*?>).*?</td>");
        Matcher matcher = pattern.matcher(html);
        Map<String, String> map = new HashMap<>();
        while (matcher.find()) {
            String oldTd = matcher.group();
            if (StringUtils.isEmpty(Jsoup.parse(oldTd).text())) {
                String group1 = matcher.group(1);
                String newTd = oldTd.replaceFirst(group1, group1 + "&nbsp;");
                map.put(oldTd, newTd);
            }
        }
        for (Map.Entry<String, String> entry : map.entrySet()) {
            String key = entry.getKey();
            String value = entry.getValue();
            html = html.replaceAll(key, value);
        }
        return html;
    }

    private static void dealTableConnect(List<String> result) {
        if (result.size() > 0) {
            String last = result.get(result.size() - 1);
            boolean matches = last.matches("<table.*?</table>");
            if (matches) {
                result.add("<p></p>");
            }
        }
    }

    /**
     * 处理p标签和h标签
     * @author 明快de玄米61
     * @date   2022/9/9 1:11
     * @param  html html标签
     * @param  result p标签和h标签集合
     * @return
     **/
    private static void dealPAndHLabel(String html, List<String> result) {
        // 处理img标签没有被p标签包裹的情况
        html = dealImgNotWrapByP(html);
        // 处理p标签和h标签
        List<LabelInfo> labelInfos = new ArrayList<>();
        Pattern pattern = Pattern.compile("(<p|<h[1-9]{1}).*?(</p>|</h[1-9]{1}>)", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(html);
        while (matcher.find()) {
            labelInfos.add(new LabelInfo(matcher.start(), matcher.end(), matcher.group()));
        }
        if (labelInfos.size() > 0) {
            for (int i = 0; i < labelInfos.size(); i++) {
                // 当前
                LabelInfo current = labelInfos.get(i);

                // 处理第一次之前的内容
                if (i == 0 && current.getStart() > 0) {
                    String s = usePWrapLabel(html.substring(0, current.getStart()));
                    if (s != null) {
                        result.add(s);
                    }
                }

                // 处理内容
                result.add(current.getHtml());

                // 处理后面的内容（注意：不处理最后一个之后的内容）
                if (i < labelInfos.size() - 1) {
                    // 下一个
                    LabelInfo next = labelInfos.get(i + 1);
                    if (current.getEnd() < next.getStart()) {
                        String s = usePWrapLabel(html.substring(current.getEnd(), next.getStart()));
                        if (s != null) {
                            result.add(s);
                        }
                    }
                }

                // 处理后面的内容（注意：只处理最后一个之后的内容）
                if (i == labelInfos.size() - 1 && current.getEnd() < html.length()) {
                    String s = usePWrapLabel(html.substring(current.getEnd()));
                    if (s != null) {
                        result.add(s);
                    }
                }
            }
        } else {
            String s = usePWrapLabel(html);
            if (s != null) {
                result.add(s);
            }
        }
    }

    /**
     * 反转义特殊字符
     * @author 明快de玄米61
     * @date   2022/9/9 1:11
     * @param
     * @return
     **/
    private static String reverseEscapeChar(String text) {
        return text.replaceAll("&nbsp;", " ").replaceAll("&lt;", "<").replaceAll("&gt;", ">").replaceAll("&amp;", "&");
    }

    /**
     * 设置首行缩进
     * @author 明快de玄米61
     * @date   2022/9/9 9:47
     * @param
     * @return
     **/
    private static void setTextIndent(XWPFParagraph paragraph, int num) {
        CTPPr pPr = getPrOfParagraph(paragraph);
        CTInd pInd = pPr.getInd() != null ? pPr.getInd() : pPr.addNewInd();
        pInd.setFirstLineChars(BigInteger.valueOf((long) (num * PER_CHART)));
    }
    private static CTPPr getPrOfParagraph(XWPFParagraph paragraph) {
        CTP ctp = paragraph.getCTP();
        CTPPr pPr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        return pPr;
    }

    //设置行距
    private static void setPRowSpacing(XWPFParagraph titleParagraph, double rowSpace) {
        CTP ctp = titleParagraph.getCTP();
        CTPPr ppr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
        spacing.setAfter(BigInteger.valueOf(0));
        spacing.setBefore(BigInteger.valueOf(0));
        //设置行距类型为 EXACT
        spacing.setLineRule(STLineSpacingRule.EXACT);
        //1磅数是20
        spacing.setLine(BigInteger.valueOf((int) (rowSpace * PER_POUND)));
    }

    //设置间距
    private static void setPSpacing(XWPFParagraph paragraph, Integer before, Integer after) {
        CTP ctp = paragraph.getCTP();
        CTPPr ppr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
        spacing.setAfter(BigInteger.valueOf(before * ONE_LINE));
        spacing.setBefore(BigInteger.valueOf(after * ONE_LINE));
        //设置行距类型为 EXACT
        spacing.setLineRule(STLineSpacingRule.EXACT);
    }

    //设置字体字号
    private static void setTextFontFamilyAndFontSize(XWPFRun xwpfRun, String fontFamily, int fontSize) {
        xwpfRun.setFontFamily(fontFamily);
        xwpfRun.setFontSize(fontSize);
    }

    /**
     * 初始化页码
     * @author 明快de玄米61
     * @date   2022/9/13 11:30
     * @param  document XWPFDocument对象
     **/
    private static void initFooter(XWPFDocument document, CTSectPr sectPr, String fontFamily, int fontSize, String color, String prefix, String suffix) {
        // delete start by 明快de玄米61 time 2022/9/30 reason 使用其他方案解决页码字体不是宋体的问题
//        // 创建页脚对象
//        XWPFFooter footer = document.createFooter(HeaderFooterType.DEFAULT);
//        XWPFParagraph paragraph = footer.createParagraph();
//        // 水平右对齐
//        paragraph.setAlignment(ParagraphAlignment.RIGHT);
//        // 垂直居中
//        paragraph.setVerticalAlignment(TextAlignment.CENTER);
//
//        // 处理前缀
//        setStyle(paragraph.createRun(), fontFamily, fontSize, false, color, prefix);
//
//        // 处理页码
//        CTFldChar fldChar = paragraph.createRun().getCTR().addNewFldChar();
//        fldChar.setFldCharType(STFldCharType.Enum.forString("begin"));
//        XWPFRun numberRun = paragraph.createRun();
//        CTText ctText = numberRun.getCTR().addNewInstrText();
//        ctText.setStringValue("PAGE  \\* MERGEFORMAT");
//        ctText.setSpace(SpaceAttribute.Space.Enum.forString("preserve"));
//        setStyle(numberRun, fontFamily, fontSize, false, color, null);
//        fldChar = paragraph.createRun().getCTR().addNewFldChar();
//        fldChar.setFldCharType(STFldCharType.Enum.forString("end"));
//
//        // 处理后缀
//        setStyle(paragraph.createRun(), fontFamily, fontSize, false, color, suffix);
        // delete end by 明快de玄米61 time 2022/9/30 reason 使用其他方案解决页码字体不是宋体的问题

        // add start by 明快de玄米61 time 2022/9/30 reason 使用下面方案解决页码字体不是宋体的问题
        // 创建页码对象
        CTP pageNo = CTP.Factory.newInstance();

        // 创建段落对象
        XWPFParagraph paragraph = new XWPFParagraph(pageNo, document);
        // 水平右对齐
        paragraph.setAlignment(ParagraphAlignment.RIGHT);
        // 垂直居中
        paragraph.setVerticalAlignment(TextAlignment.CENTER);

        // 处理前缀
        setStyle(paragraph.createRun(), fontFamily, fontSize, false, color, prefix);

        // 处理页码
        int doubleFontSize = fontSize * 2;
        CTPPr begin = pageNo.addNewPPr();
        begin.addNewPStyle().setVal("style21");
        begin.addNewJc().setVal(STJc.RIGHT);
        CTR  pageBegin=pageNo.addNewR();
        pageBegin.addNewRPr().addNewRFonts().setAscii(fontFamily);
        pageBegin.addNewRPr().addNewRFonts().setCs(fontFamily);
        pageBegin.addNewRPr().addNewRFonts().setEastAsia(fontFamily);
        pageBegin.addNewRPr().addNewRFonts().setHAnsi(fontFamily);
        pageBegin.addNewRPr().addNewSz().setVal(BigInteger.valueOf(doubleFontSize));
        pageBegin.addNewRPr().addNewSzCs().setVal(BigInteger.valueOf(doubleFontSize));
        pageBegin.addNewFldChar().setFldCharType(STFldCharType.BEGIN);
        CTR  page=pageNo.addNewR();
        page.addNewRPr().addNewRFonts().setAscii(fontFamily);
        page.addNewRPr().addNewRFonts().setCs(fontFamily);
        page.addNewRPr().addNewRFonts().setEastAsia(fontFamily);
        page.addNewRPr().addNewRFonts().setHAnsi(fontFamily);
        page.addNewRPr().addNewSz().setVal(BigInteger.valueOf(doubleFontSize));
        page.addNewRPr().addNewSzCs().setVal(BigInteger.valueOf(doubleFontSize));
        page.addNewInstrText().setStringValue("PAGE   \\* MERGEFORMAT");
        CTR  pageSep=pageNo.addNewR();
        pageSep.addNewRPr().addNewRFonts().setAscii(fontFamily);
        pageSep.addNewRPr().addNewRFonts().setCs(fontFamily);
        pageSep.addNewRPr().addNewRFonts().setEastAsia(fontFamily);
        pageSep.addNewRPr().addNewRFonts().setHAnsi(fontFamily);
        pageSep.addNewRPr().addNewSz().setVal(BigInteger.valueOf(doubleFontSize));
        pageSep.addNewRPr().addNewSzCs().setVal(BigInteger.valueOf(doubleFontSize));
        pageSep.addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
        CTR end = pageNo.addNewR();
        CTRPr endRPr = end.addNewRPr();
        endRPr.addNewNoProof();
        endRPr.addNewLang().setVal("zh-CN");
        end.addNewRPr().addNewRFonts().setAscii(fontFamily);
        end.addNewRPr().addNewRFonts().setCs(fontFamily);
        end.addNewRPr().addNewRFonts().setEastAsia(fontFamily);
        end.addNewRPr().addNewRFonts().setHAnsi(fontFamily);
        end.addNewRPr().addNewSz().setVal(BigInteger.valueOf(doubleFontSize));
        end.addNewRPr().addNewSzCs().setVal(BigInteger.valueOf(doubleFontSize));
        end.addNewFldChar().setFldCharType(STFldCharType.END);

        // 处理后缀
        setStyle(paragraph.createRun(), fontFamily, fontSize, false, color, suffix);

        // 不太明白含义，但是不添加就无法生成页码
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);
        policy.createFooter(STHdrFtr.DEFAULT, new XWPFParagraph[] { paragraph });
        // add end by 明快de玄米61 time 2022/9/30 reason 使用下面方案解决页码字体不是宋体的问题
    }

    private static void setStyle(XWPFRun run, String fontFamily, int fontSize, boolean bold, String color, String text) {
        run.setBold(bold);
        run.setFontFamily(fontFamily);
        run.setFontSize(fontSize);
        if(!StringUtils.isEmpty(text)){
            run.setText(text);
        }
        run.setColor(StringUtils.isEmpty(color) ? "000000" : color);
    }

    /**
     * 初始化页边距
     * @author 明快de玄米61
     * @date   2022/9/13 11:01
     * @param  document XWPFDocument对象
     **/
    private static void initHeadingStyle(XWPFDocument document) {
        for (int i = 1; i <= MAX_HEADING_LEVEL; i++) {
            String heading = getHeadingStyle(i);
            createHeadingStyle(document, heading, i);
        }
    }

    /**
     * 初始化页边距
     * @author 明快de玄米61
     * @date   2022/9/13 11:01
     * @param  sectPr CTSectPr对象
     **/
    private static void initPageMargin(CTSectPr sectPr) {
        CTPageMar ctpagemar = sectPr.addNewPgMar();
        ctpagemar.setTop(new BigInteger(String.valueOf(Math.round(TOP * CM_2_POUND * PER_POUND))));
        ctpagemar.setBottom(new BigInteger(String.valueOf(Math.round(BOTTOM * CM_2_POUND * PER_POUND))));
        ctpagemar.setLeft(new BigInteger(String.valueOf(Math.round(LEFT * CM_2_POUND * PER_POUND))));
        ctpagemar.setRight(new BigInteger(String.valueOf(Math.round(RIGHT * CM_2_POUND * PER_POUND))));
    }

    private static void createHeadingStyle(XWPFDocument doc, String strStyleId, int headingLevel) {
        //创建样式
        CTStyle ctStyle = CTStyle.Factory.newInstance();
        //设置id
        ctStyle.setStyleId(strStyleId);

        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);

        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        // 数字越低在格式栏中越突出
        ctStyle.setUiPriority(indentNumber);

        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);

        // 样式将显示在“格式”栏中
        ctStyle.setQFormat(onoffnull);

        // 样式定义给定级别的标题
        CTPPr ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);

        XWPFStyle style = new XWPFStyle(ctStyle);

        // 获取新建文档对象的样式
        style.setType(STStyleType.PARAGRAPH);
        XWPFStyles styles = doc.createStyles();
        styles.addStyle(style);
    }

    private static void parseTableElement(Element child, XWPFDocument document){
        //先将合并的行列补齐，再对补齐后的表格进行数据处理
        child = simplifyTable(child);
        Elements trList = child.select("tr");
        Elements thList=trList.first().select("th");
        Elements tdList = trList.get(0).getElementsByTag("td");
        XWPFTable table;
        Map<String,Boolean>[][] array;
        if(tdList.isEmpty()) {
//        	 String colspan = thList.attr("colspan");
//        	 if(!StringUtils.isEmpty(colspan)){
//        		 table = document.createTable(trList.size(), Integer.valueOf(colspan));
//        		 array = new Map[trList.size()][Integer.valueOf(colspan)];
//        	 }else {
            table = document.createTable(trList.size(), thList.size());
            array = new Map[trList.size()][thList.size()];
//        	 }

        }else{
            table = document.createTable(trList.size(), tdList.size());
            array = new Map[trList.size()][tdList.size()];
        }
//      Map<String,Boolean>[][] array = new Map[trList.size()][tdList.size()];

        //表格属性
        CTTblPr tablePr = table.getCTTbl().addNewTblPr();
        //表格宽度
        CTTblWidth width = tablePr.addNewTblW();
        width.setW(BigInteger.valueOf((int)(TABLE_WIDTH * CM_2_POUND * PER_POUND)));
        //设置表格宽度为非自动
        width.setType(STTblWidth.DXA);

        for (int row = 0; row < trList.size(); row++) {
            Element trElement = trList.get(row);
            Elements tds = trElement.getElementsByTag("td");
            if(tds.isEmpty()) {
                tds=trElement.getElementsByTag("th");
            }

            for(int col = 0; col < tds.size(); col++) {
                Element colElement = tds.get(col);
                String colspan = colElement.attr("colspan");
                String rowspan = colElement.attr("rowspan");
//                String style = colElement.attr("style");
                StringBuilder styleSB = new StringBuilder();
                if(!StringUtils.isEmpty(colspan)){
                    int colCount = Integer.parseInt(colspan);
                    for(int i=0;i<colCount-1;i++){
                        try {
                            array[row][col+i+1] = new HashMap<String, Boolean>();
                            array[row][col+i+1].put("mergeCol", true);
                        }catch(Exception e) {
                            e.printStackTrace();
                        }
                    }
                }
                if(!StringUtils.isEmpty(rowspan)){
                    int rowCount = Integer.parseInt(rowspan);
                    for(int i=0;i<rowCount-1;i++){
                        array[row+i+1][col] = new HashMap<String, Boolean>();
                        array[row+i+1][col].put("mergeRow", true);
                    }
                }
                XWPFTableCell tableCell = table.getRow(row).getCell(col);
                // add start by 明快de玄米61 time 2022/9/16 reason 设置单元格边距
                setTableCellMar(tableCell, CELL_MARGIN, CELL_MARGIN, CELL_MARGIN, CELL_MARGIN);
                // add add by 明快de玄米61 time 2022/9/16 reason 设置单元格边距
                if(StringUtils.isEmpty(colspan)){
                    if(col == 0){
                        if(tableCell.getCTTc().getTcPr() == null){
                            tableCell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
                        }else{
                            if(tableCell.getCTTc().getTcPr().getHMerge() == null){
                                tableCell.getCTTc().getTcPr().addNewHMerge().setVal(STMerge.RESTART);
                            }else{
                                tableCell.getCTTc().getTcPr().getHMerge().setVal(STMerge.RESTART);
                            }
                        }
                    }else{
                        if(array[row][col]!=null && array[row][col].get("mergeCol")!=null && array[row][col].get("mergeCol")){
                            if(tableCell.getCTTc().getTcPr() == null){
                                tableCell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
                            }else{
                                if(tableCell.getCTTc().getTcPr().getHMerge() == null){
                                    tableCell.getCTTc().getTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
                                }else{
                                    tableCell.getCTTc().getTcPr().getHMerge().setVal(STMerge.CONTINUE);
                                }
                            }
                            continue;
                        }else{
                            if(tableCell.getCTTc().getTcPr() == null){
                                tableCell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
                            }else{
                                if(tableCell.getCTTc().getTcPr().getHMerge() == null){
                                    tableCell.getCTTc().getTcPr().addNewHMerge().setVal(STMerge.RESTART);
                                }else{
                                    tableCell.getCTTc().getTcPr().getHMerge().setVal(STMerge.RESTART);
                                }
                            }
                        }
                    }
                }else{
                    if(tableCell.getCTTc().getTcPr() == null){
                        tableCell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
                    }else{
                        if(tableCell.getCTTc().getTcPr().getHMerge() == null){
                            tableCell.getCTTc().getTcPr().addNewHMerge().setVal(STMerge.RESTART);
                        }else{
                            tableCell.getCTTc().getTcPr().getHMerge().setVal(STMerge.RESTART);
                        }
                    }
                }
                if(StringUtils.isEmpty(rowspan)){
                    if(array[row][col]!=null && array[row][col].get("mergeRow")!=null && array[row][col].get("mergeRow")){
                        if(tableCell.getCTTc().getTcPr() == null){
                            tableCell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
                        }else{
                            if(tableCell.getCTTc().getTcPr().getVMerge() == null){
                                tableCell.getCTTc().getTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
                            }else{
                                tableCell.getCTTc().getTcPr().getVMerge().setVal(STMerge.CONTINUE);
                            }
                        }
                        continue;
                    }else{
                        if(tableCell.getCTTc().getTcPr() == null){
                            tableCell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
                        }else{
                            if(tableCell.getCTTc().getTcPr().getVMerge() == null){
                                tableCell.getCTTc().getTcPr().addNewVMerge().setVal(STMerge.RESTART);
                            }else{
                                tableCell.getCTTc().getTcPr().getVMerge().setVal(STMerge.RESTART);
                            }
                        }
                    }
                }else{
                    if(tableCell.getCTTc().getTcPr() == null){
                        tableCell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
                    }else{
                        if(tableCell.getCTTc().getTcPr().getVMerge() == null){
                            tableCell.getCTTc().getTcPr().addNewVMerge().setVal(STMerge.RESTART);
                        }else{
                            tableCell.getCTTc().getTcPr().getVMerge().setVal(STMerge.RESTART);
                        }
                    }
                }
                tableCell.removeParagraph(0);
                tableCell.setVerticalAlignment(XWPFVertAlign.CENTER);
                parseingCell(tableCell,colElement);
            }
        }
    }

    /**
     * @Description 设置单元格边距
     * @param cell 待设置的单元格
     * @param top   上边距 磅
     * @param bottom 下边距 磅
     * @param left  左边距 磅
     * @param right 右边距 磅
     */
    private static void setTableCellMar(XWPFTableCell cell, double top, double bottom, double left, double right) {
        CTTcPr cttcpr = getCttcpr(cell);
        CTTcMar ctTcMar = cttcpr.isSetTcMar() ? cttcpr.getTcMar() : cttcpr.addNewTcMar();
        if(left >= 0){
            (ctTcMar.isSetLeft() ? ctTcMar.getLeft() : ctTcMar.addNewLeft()).setW(BigInteger.valueOf(Math.round(left * PER_POUND)));
        }
        if(top >= 0){
            (ctTcMar.isSetTop() ? ctTcMar.getTop() : ctTcMar.addNewTop()).setW(BigInteger.valueOf(Math.round(top * PER_POUND)));
        }
        if(right >= 0){
            (ctTcMar.isSetRight() ? ctTcMar.getRight() : ctTcMar.addNewRight()).setW(BigInteger.valueOf(Math.round(right * PER_POUND)));
        }
        if(bottom >= 0){
            (ctTcMar.isSetBottom() ? ctTcMar.getBottom() : ctTcMar.addNewBottom()).setW(BigInteger.valueOf(Math.round(bottom * PER_POUND)));
        }
    }

    private static CTTcPr getCttcpr(XWPFTableCell cell){
        CTTc ctTc = cell.getCTTc();
        return ctTc.isSetTcPr() ? ctTc.getTcPr() : ctTc.addNewTcPr();
    }

    private static void parseingCell(XWPFTableCell tableCell, Element colElement){
        Elements children = colElement.children();
        if(!children.isEmpty()) {
            for (Element element : children) {
                if(!element.children().isEmpty()) {
                    parseingCell(tableCell,element);
                }else {
                    parseingCellImg(tableCell, element);
                }

            }
        }
        if(colElement.hasText()) {
            parseingCellImg(tableCell, colElement);
        }
    }

    private static Element simplifyTable(Element child) {
        Elements trElements = child.select("tr");
        if (trElements != null) {
            Iterator<Element> eleIterator = trElements.iterator();
            Integer rowNum = 0;
            // 针对于colspan操作
            while (eleIterator.hasNext()) {
                rowNum++;
                Element trElement = eleIterator.next();
                // 去除所有样式
                trElement.removeAttr("class");
                Elements tdElements = trElement.select("td");
                List<Element> tdEleList = covertElements2List(tdElements);
                for (int i = 0; i < tdEleList.size(); i++) {
                    Element curTdElement = tdEleList.get(i);
                    // 去除所有样式
                    curTdElement.removeAttr("class");
                    Element ele = curTdElement.clone();
                    String colspanValStr = curTdElement.attr("colspan");
                    if (!StringUtils.isEmpty(colspanValStr)) {
                        ele.removeAttr("colspan");
                        Integer colspanVal = Integer.parseInt(colspanValStr);
                        for (int k = 0; k < colspanVal - 1; k++) {
                            curTdElement.after(ele.outerHtml());
                        }
                    }
                }
            }
            // 针对于rowspan操作
            List<Element> trEleList = covertElements2List(trElements);
            Element firstTrEle = trElements.first();
            Elements tdElements = firstTrEle.select("td");
            if(tdElements.isEmpty()) {
                tdElements=firstTrEle.select("th");
            }
            Integer tdCount = tdElements.size();
            for (int i = 0; i < tdElements.size(); i++) { // 获取该列下所有单元格
                for (Element trElement : trEleList) {
                    List<Element> tdElementList = covertElements2List(trElement.select("td"));
                    try {
                        tdElementList.get(i);
                    } catch (Exception e) {
                        continue;
                    }
                    Node curTdNode = tdElementList.get(i);
                    Node cNode = curTdNode.clone();
                    String rowspanValStr = curTdNode.attr("rowspan");
                    if (!StringUtils.isEmpty(rowspanValStr)) {
                        cNode.removeAttr("rowspan");
                        Element nextTrElement = trElement.nextElementSibling();
                        Integer rowspanVal = Integer.parseInt(rowspanValStr);
                        for (int j = 0; j < rowspanVal - 1; j++) {
                            Node tempNode = cNode.clone();
                            List<Node> nodeList = new ArrayList<Node>();
                            nodeList.add(tempNode);
                            if (j > 0) {
                                nextTrElement = nextTrElement.nextElementSibling();
                            }
                            Integer indexNum = i ;

                            if (i == 0)
                            {
                                indexNum = 0;
                            }
                            if (indexNum == tdCount) {
                                nextTrElement.appendChild(tempNode);
                            }else {
                                nextTrElement.insertChildren(indexNum, nodeList);
                            }
                        }
                    }
                }
            }
        }
        Element tableEle = child.getElementsByTag("table").first();
        return tableEle;
    }

    private static List<Element> covertElements2List(Elements curElements) {
        List<Element> elementList = new ArrayList<Element>();
        Iterator<Element> eleIterator = curElements.iterator();
        while (eleIterator.hasNext()) {
            Element curlement = eleIterator.next();
            elementList.add(curlement);
        }
        return elementList;
    }

    private static void parseingCellImg(XWPFTableCell tableCell, Element element) {
        if((element.toString().startsWith("<img")||element.toString().startsWith("<p><img"))) {
            String src=element.attr("src");
            int res=getPictureType(src);
            String width=element.attr("width");
            String height=element.attr("height");
            XWPFParagraph paragraph = tableCell.addParagraph();
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun run = paragraph.createRun();
            // add start by 明快de玄米61 time 2022/9/8 reason 图片下载
            File imgFile = null;
            try {
                imgFile = ImageUtil.getImgFile(src);
                FileInputStream iss = new FileInputStream(imgFile);
                BufferedImage image = getImgByFilePath(imgFile.getAbsolutePath());
                if (image == null) {
                    width = StringUtils.isEmpty(width) ? String.valueOf(MAX_TABLE_IMG_WIDTH) : width;
                    height = StringUtils.isEmpty(width) ? String.valueOf(MAX_TABLE_IMG_WIDTH) : width;
                } else {
                    width = StringUtils.isEmpty(width) ? String.valueOf(image.getWidth()) : width;
                    BigDecimal originalWidth = new BigDecimal(image.getWidth());
                    BigDecimal originalHeight = new BigDecimal(image.getHeight());
                    height = StringUtils.isBlank(height) ? originalHeight.multiply(new BigDecimal(width).divide(originalWidth, 10, BigDecimal.ROUND_HALF_UP)).toBigInteger().toString() : height;
                    if (Integer.valueOf(width) >= MAX_TABLE_IMG_WIDTH) {
                        BigDecimal widthBigDecimal = new BigDecimal(width);
                        BigDecimal heightBigDecimal = new BigDecimal(height);
                        BigDecimal divide = new BigDecimal(MAX_TABLE_IMG_WIDTH).divide(widthBigDecimal, 10, BigDecimal.ROUND_HALF_UP);
                        width = widthBigDecimal.multiply(divide).toBigInteger().toString();
                        height = heightBigDecimal.multiply(divide).toBigInteger().toString();
                    }
                }
                // add end by 明快de玄米61 time 2022/9/8 reason 图片下载
                // delete start by 明快de玄米61 time 2022/9/8 reason 本次图片是超链接
//                FileInputStream iss=new FileInputStream(src);
                // delete start by 明快de玄米61 time 2022/9/8 reason 本次图片是超链接
                try {
                    run.addPicture(iss, res, "", Units.toEMU(Double.valueOf(width)), Units.toEMU(Double.valueOf(height)));
//                                        r9.setTextPosition(28);
                    iss.close();
                } catch (NumberFormatException e1) {
                    e1.printStackTrace();
                } catch (InvalidFormatException e1) {
                    e1.printStackTrace();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                if (imgFile != null) {
                    FileUtil.deleteParentFile(imgFile);
                }
            }
        }
        else {
            XWPFParagraph paragraph = tableCell.addParagraph();
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun run = paragraph.createRun();
            run.setText(reverseEscapeChar(element.text()));
            run.setFontFamily(PARA_FONT_FAMILY);
            run.setFontSize(PARA_FONT_SIZE);
        }
    }

    /**
     * 根据图片路径获取图片
     * @param path
     * @return
     * @throws Exception
     */
    private static BufferedImage getImgByFilePath(String path) {
        try {

            FileInputStream fis = new FileInputStream(path);
            byte[] byteArray = IOUtils.toByteArray(fis);
            ByteArrayInputStream bais = new ByteArrayInputStream(byteArray);
            return  ImageIO.read(bais);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private static int getPictureType(String src){
        String picType = "jpg";
        if(src.contains("png")) {
            picType="png";
        }else if(src.contains("jpg")) {
            picType="jpg";
        }else if(src.contains("jpeg")) {
            picType="jpeg";
        }else if(src.contains("gif")) {
            picType="gif";
        }
        int res = XWPFDocument.PICTURE_TYPE_PICT;
        if(!StringUtils.isEmpty(picType)){
            if(picType.equalsIgnoreCase("png")){
                res = XWPFDocument.PICTURE_TYPE_PNG;
            }else if(picType.equalsIgnoreCase("dib")){
                res = XWPFDocument.PICTURE_TYPE_DIB;
            }else if(picType.equalsIgnoreCase("emf")){
                res = XWPFDocument.PICTURE_TYPE_EMF;
            }else if(picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")){
                res = XWPFDocument.PICTURE_TYPE_JPEG;
            }else if(picType.equalsIgnoreCase("wmf")){
                res = XWPFDocument.PICTURE_TYPE_WMF;
            }
        }else {
            res = XWPFDocument.PICTURE_TYPE_JPEG;
        }
        return res;
    }

    private static void createDefaultHeader(XWPFDocument docx, String text) throws IOException, XmlException{
        CTP ctp = CTP.Factory.newInstance();
        XWPFParagraph paragraph = new XWPFParagraph(ctp, docx);
        ctp.addNewR().addNewT().setStringValue(text);
        ctp.addNewR().addNewT().setSpace(SpaceAttribute.Space.PRESERVE);
        CTSectPr sectPr = docx.getDocument().getBody().isSetSectPr() ? docx.getDocument().getBody().getSectPr() : docx.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(docx, sectPr);
        XWPFHeader header = policy.createHeader(STHdrFtr.DEFAULT, new XWPFParagraph[] { paragraph });
        header.setXWPFDocument(docx);
    }

    private static void setDocumentMargin(XWPFDocument document, String left, String top, String right, String bottom) {
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        CTPageMar ctpagemar = sectPr.addNewPgMar();
        if (StringUtils.isNotBlank(left)) {
            ctpagemar.setLeft(new BigInteger(left));
        }
        if (StringUtils.isNotBlank(top)) {
            ctpagemar.setTop(new BigInteger(top));
        }
        if (StringUtils.isNotBlank(right)) {
            ctpagemar.setRight(new BigInteger(right));
        }
        if (StringUtils.isNotBlank(bottom)) {
            ctpagemar.setBottom(new BigInteger(bottom));
        }
    }

    private static void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel) {
        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);

        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);

        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        ctStyle.setUiPriority(indentNumber);

        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);


        ctStyle.setQFormat(onoffnull);


        CTPPr ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);

        XWPFStyle style = new XWPFStyle(ctStyle);
        XWPFStyles styles = docxDocument.createStyles();

        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);
    }

    private static void setImgSpacing(XWPFParagraph paragraph) {
        CTPPr ppr = paragraph.getCTP().getPPr();
        if (ppr == null) ppr = paragraph.getCTP().addNewPPr();
        CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
        spacing.setAfter(BigInteger.valueOf(0));
        spacing.setBefore(BigInteger.valueOf(0));
        spacing.setLineRule(STLineSpacingRule.AUTO);
    }

    private static String getStringNoBlank(String str) {
        if(str!=null && !"".equals(str)) {
            Pattern p = Pattern.compile("\\*|\\\\|\\:|\\?|\\<|\\>|\\/|\"|\\|");
            Matcher m = p.matcher(str);
            String strNoBlank = m.replaceAll("");
            return strNoBlank;
        }else {
            return str;
        }
    }

    private static double length(String value) {
        double valueLength = 0;
        String chinese = "[\u0391-\uFFE5]";
        // 获取字段值的长度，如果含中文字符，则每个中文字符长度为2，否则为1
        for (int i = 0; i < value.length(); i++) {
            // 获取一个字符
            String temp = value.substring(i, i + 1);
            // 判断是否为中文字符
            if (temp.matches(chinese)) {
                // 中文字符长度为2
                valueLength += 1;
            } else {
                // 其他字符长度为1
                valueLength += 0.5;
            }
        }
        return valueLength;
    }

    private static List<String> dealHElement(String content){
        Pattern p=Pattern.compile("<h.*?>.*?</h.*?>");
        Matcher m=p.matcher(content);
        List<String> realcontents=new ArrayList<String>();
        int begin=0;
        while(m.find()) {
            String h=m.group();
            int length=content.indexOf(h);
            realcontents.add(content.substring(begin,length));
            realcontents.add(h.replaceAll("<h.*?>|</h.*?>", ""));
            begin=length+h.length();
        }
        realcontents.add(content.substring(begin));
        return realcontents;
    }

    /**
     * 向Word中插入图片(仅支持png格式图片, 未完待续...)
     * @param imagePath 图片文件路径
     * @throws Exception
     */
    private static void writeImage(XWPFParagraph paragraph, String imagePath, Integer width) throws Exception {
        XWPFRun run = paragraph.createRun();
        BufferedImage image = getImgByFilePath(imagePath);
        int res=getPictureType(imagePath);
        int height;
        if (image == null) {
            width = width == null ? MAX_PAGE_IMG_WIDTH : width;
            height = width == null ? MAX_PAGE_IMG_WIDTH : width;
        } else {
            width = width == null ? image.getWidth() : width;
            BigDecimal originalWidth = new BigDecimal(image.getWidth());
            BigDecimal originalHeight = new BigDecimal(image.getHeight());
            height = originalHeight.multiply(new BigDecimal(width).divide(originalWidth,10,BigDecimal.ROUND_HALF_UP)).toBigInteger().intValue();
            if (width >= MAX_PAGE_IMG_WIDTH) {
                BigDecimal widthBigDecimal = new BigDecimal(width);
                BigDecimal heightBigDecimal = new BigDecimal(height);
                BigDecimal divide = new BigDecimal(MAX_PAGE_IMG_WIDTH).divide(widthBigDecimal,10,BigDecimal.ROUND_HALF_UP);
                width = widthBigDecimal.multiply(divide).toBigInteger().intValue();
                height = heightBigDecimal.multiply(divide).toBigInteger().intValue();
            }
        }
        run.addPicture(new FileInputStream(imagePath), res, "",
                Units.toEMU(width), Units.toEMU(height));
    }


    static class TableInfo {
        private Integer start;
        private Integer end;
        private String html;

        public TableInfo(Integer start, Integer end, String html) {
            this.start = start;
            this.end = end;
            this.html = html;
        }

        public Integer getStart() {
            return start;
        }

        public void setStart(Integer start) {
            this.start = start;
        }

        public Integer getEnd() {
            return end;
        }

        public void setEnd(Integer end) {
            this.end = end;
        }

        public String getHtml() {
            return html;
        }

        public void setHtml(String html) {
            this.html = html;
        }
    }

    static class LinkInfo {
        private Integer start;
        private Integer end;
        private String href;
        private String html;

        public LinkInfo(Integer start, Integer end, String href, String html) {
            this.start = start;
            this.end = end;
            this.href = href;
            this.html = html;
        }

        public Integer getStart() {
            return start;
        }

        public void setStart(Integer start) {
            this.start = start;
        }

        public Integer getEnd() {
            return end;
        }

        public void setEnd(Integer end) {
            this.end = end;
        }

        public String getHref() {
            return href;
        }

        public void setHref(String href) {
            this.href = href;
        }

        public String getHtml() {
            return html;
        }

        public void setHtml(String html) {
            this.html = html;
        }
    }

    /**
     * 标题样式
     * @author 明快de玄米61
     * @date   2022/9/13 10:36
     **/
    static class HeadingStyle implements Serializable {
        private static final long serialVersionUID = 1L;
        /** 字号 **/
        private Integer fontSize;
        /** 字体 **/
        private String fontFamily;
        /** 加粗 **/
        private boolean bold;

        public Integer getFontSize() {
            return fontSize;
        }

        public void setFontSize(Integer fontSize) {
            this.fontSize = fontSize;
        }

        public String getFontFamily() {
            return fontFamily;
        }

        public void setFontFamily(String fontFamily) {
            this.fontFamily = fontFamily;
        }

        public boolean isBold() {
            return bold;
        }

        public void setBold(boolean bold) {
            this.bold = bold;
        }

        HeadingStyle(Integer fontSize, String fontFamily, boolean bold) {
            this.fontSize = fontSize;
            this.fontFamily = fontFamily;
            this.bold = bold;
        }
    }

    static class LabelInfo {
        public LabelInfo(Integer start, Integer end, String html) {
            this.start = start;
            this.end = end;
            this.html = html;
        }

        private Integer start;
        private Integer end;
        private String html;

        public Integer getStart() {
            return start;
        }

        public void setStart(Integer start) {
            this.start = start;
        }

        public Integer getEnd() {
            return end;
        }

        public void setEnd(Integer end) {
            this.end = end;
        }

        public String getHtml() {
            return html;
        }

        public void setHtml(String html) {
            this.html = html;
        }
    }
}

