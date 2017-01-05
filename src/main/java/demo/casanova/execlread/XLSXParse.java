package demo.casanova.execlread;

/**
 * 工程:execlread
 * 文件:XLSXParse.java
 * @author:Casanova.Z
 * Time:2015/12/26
 */
import android.util.Xml;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;


public class XLSXParse {

    private String _armStr;
    private OutFileType _outFileType;
    private String _spiltStr;
    private ArmFileType _armFileType;

    public XLSXParse() {

    }

    public XLSXParse(Builder builder) {

        this._armStr = builder._armStr;
        this._outFileType = builder._outFileType;
        this._spiltStr = builder._spiltStr;
        this._armFileType =builder._armFileType;
    }
    // 定义最后输出的数据类型
    public enum OutFileType{
        FILE_TYPE_JSON,//json格式输出
        FILE_TYPE_SPILT,//分隔符字符串输出
        FILE_TYPE_LIST,//直接输出List类型数据
        FILE_TYPE_ARRAY//输出字符串型二位数组
    }

    public enum ArmFileType{
        XLS,
        XLSX
    }

    public static class Builder {
        private String _armStr = null;
        private OutFileType _outFileType;
        private String _spiltStr;
        private ArmFileType _armFileType;

        public Builder() {
            this._spiltStr="|";
            this._outFileType=OutFileType.FILE_TYPE_LIST;
        }

        public Builder(String armFilePath,String split,OutFileType fileType) {
            this._armStr=armFilePath;
            this._spiltStr=split;
            this._outFileType=fileType;
        }
        //设置以什么格式输出
        public Builder setOutFileType(OutFileType outFileType){

            this._outFileType=outFileType;
            return this;
        }

        // 设置要解析的XLSX文件路径,其中带文件名
        public Builder setArmFilePath(String armFilePath) {
            this._armStr = armFilePath;
            return this;
        }

        public Builder setSplitString(String splitString){
            this._spiltStr=splitString;
            return this;
        }

        public Builder setArmFileType(ArmFileType armFileType){
            this._armFileType = armFileType;
            return this;
        }

        public XLSXParse build() {

            return new XLSXParse(this);
        }
    }

    //
    private void judgeArmFileType(){
        String type=this._armStr.substring(this._armStr.lastIndexOf(".")+1);
        if(type!=null){
            if(type.equals("xlsx")){
                this._armFileType = ArmFileType.XLSX;

            }else if(type.equals("xls")){
                this._armFileType = ArmFileType.XLS;
            }
        }
    }

    public Object parseFile(){
        judgeArmFileType();
        Object arm=null;
        switch (this._armFileType){
            case XLSX:
                arm = parseXLSX();
                break;
            case XLS:
                arm = parseXLS();
                break;
        }
        return arm;
    }
    /**
     * 开始处理xlsx,根据设置返回相应的数据
     * 1⃣️ JSON格式的字符串
     * 2⃣️ List数据
     * 3⃣️ 用指定字符隔开的字符串
     * @return
     */
    private Object parseXLSX(){
        List<Map<String,String>> list = readXLSX();
        Object armObj=null;
        if(list.size()>0){
            switch (this._outFileType){
                case FILE_TYPE_JSON:
                    armObj = new JSONArray(list);
                    break;
                case FILE_TYPE_LIST:
                    armObj = list;
                    break;
                case FILE_TYPE_SPILT:
                    StringBuilder sb=new StringBuilder();
                    for (int i = 0; i <list.size() ; i++) {
                        Map<String,String> map =list.get(i);
                        for (Map.Entry entry : map.entrySet()) {
                            Object key = entry.getKey();
                            sb.append(key+":"+map.get(key)+this._spiltStr);
                        }
                    }
                    sb.deleteCharAt(sb.toString().trim().length() - 1);
                    armObj=sb.toString();
                    break;
                case FILE_TYPE_ARRAY:

                    break;
            }
        }
        return armObj;
    }
    // 读取文件内容并且解析
    private List<Map<String,String>> readXLSX() {

        List<Map<String,String>> armList=new ArrayList<>();
        String str = "";
        String v = null;
        boolean flat = false;
        List<String> ls = new ArrayList<String>();
        try {
            File file =new File(this._armStr);
            ZipFile xlsxFile = new ZipFile(file);
            ZipEntry sharedStringXML = xlsxFile
                    .getEntry("xl/sharedStrings.xml");
            InputStream inputStream = xlsxFile.getInputStream(sharedStringXML);
            XmlPullParser xmlParser = Xml.newPullParser();
            xmlParser.setInput(inputStream, "utf-8");
            int evtType = xmlParser.getEventType();
            while (evtType != XmlPullParser.END_DOCUMENT) {
                switch (evtType) {
                    case XmlPullParser.START_TAG:
                        String tag = xmlParser.getName();
                        if (tag.equalsIgnoreCase("t")) {
                            ls.add(xmlParser.nextText());
                        }
                        break;
                    case XmlPullParser.END_TAG:
                        break;
                    default:
                        break;
                }
                evtType = xmlParser.next();
            }
            ZipEntry sheetXML = xlsxFile.getEntry("xl/worksheets/sheet1.xml");
            InputStream inputStreamsheet = xlsxFile.getInputStream(sheetXML);
            XmlPullParser xmlParsersheet = Xml.newPullParser();
            xmlParsersheet.setInput(inputStreamsheet, "utf-8");
            int evtTypesheet = xmlParsersheet.getEventType();
            String r="";
            while (evtTypesheet != XmlPullParser.END_DOCUMENT) {
                switch (evtTypesheet) {
                    case XmlPullParser.START_TAG:
                        // 获取文件中的各个节点
                        String tag = xmlParsersheet.getName();
                        /**
                         * 判断各个节点的值属于哪一类
                         */
                        if (tag.equalsIgnoreCase("row")) {// 如果xlsx读取到的节点值为row

                        } else if (tag.equalsIgnoreCase("c")) {// 如果xlsx读取到的节点值为c

                            String t = xmlParsersheet.getAttributeValue(null, "t");
                            r= null;
                            r = xmlParsersheet.getAttributeValue(null, "r");
                            if (t != null) {
                                flat = true;
                            } else {
                                flat = false;
                            }
                        } else if (tag.equalsIgnoreCase("v")) {
                            v = xmlParsersheet.nextText();
                            if (v != null) {
                                Map<String, String> map = new HashMap<>();
                                if (flat) {
                                    str = ls.get(Integer.parseInt(v)) + "";
                                }else{
                                    str = v+"";
                                }
                                map.put(r, str);
                                armList.add(map);
                            }
                        }
                        break;
                    case XmlPullParser.END_TAG:
                        break;
                }
                evtTypesheet = xmlParsersheet.next();
            }
        } catch (ZipException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (XmlPullParserException e) {
            e.printStackTrace();
        }
        return armList;
    }
    /**
     * 开始处理xls,根据设置返回相应的数据
     * 1⃣️ JSON格式的字符串
     * 2⃣️ List数据
     * 3⃣️ 用指定字符隔开的字符串
     * 4⃣️ 二维数组
     * @return
     */
    private Object parseXLS(){
        String[][] data = readXLS();
        Object armObj=null;
        if(data.length>0){
            switch (this._outFileType){
                case FILE_TYPE_JSON:
                    JSONArray arm =new JSONArray();
                    for (int i = 0; i <data.length ; i++) {
                        JSONObject jsonArr=new JSONObject();
                        for(int j=0;j<data[i].length;j++){
                            try {
                                jsonArr.put(i+"_"+j,data[i][j]);
                            } catch (JSONException e) {
                                e.printStackTrace();
                            }
                        }
                        arm.put(jsonArr);
                    }
                    armObj = arm;
                    break;
                case FILE_TYPE_LIST:
                    ArrayList<Map<String,String>> list =new ArrayList<>();
                    for (int i = 0; i <data.length ; i++) {
                        Map<String,String> map=new HashMap<>();
                        for (int j=0;j<data[i].length;j++){
                            map.put(i+"_"+j,data[i][j]);
                            list.add(map);
                        }
                    }
                    armObj = list;
                    break;
                case FILE_TYPE_SPILT:
                    StringBuilder sb=new StringBuilder();
                    for (int i = 0; i <data.length ; i++) {
                        for (int j=0;j<data[i].length;j++){
                            sb.append(i+"_"+j+":"+data[i][j]+this._spiltStr);
                        }
                    }
                    sb.deleteCharAt(sb.toString().trim().length() - 1);
                    armObj=sb.toString();
                    break;
                case FILE_TYPE_ARRAY:
                    armObj =data;
                    break;
            }
        }
        return armObj;
    }
    //读取xls文件内容
    public String[][] readXLS() {
        String[][]data=null;
        try {
            Workbook workbook = null;
            try {
                File file=new File(this._armStr);
                workbook = Workbook.getWorkbook(file);
            } catch (Exception e) {
                throw new Exception("File not found");
            }
            //得到第一张表
            Sheet sheet = workbook.getSheet(0);
            //列数
            int columnCount = sheet.getColumns();
            //行数
            int rowCount = sheet.getRows();
            if(columnCount>0&&rowCount>0){
                data=new String[rowCount][columnCount];
                //单元格
                Cell cell = null;
                for (int everyRow = 0; everyRow < rowCount; everyRow++) {
                    for (int everyColumn = 0; everyColumn < columnCount; everyColumn++) {
                        cell = sheet.getCell(everyColumn, everyRow);
                        data[everyRow][everyColumn]=cell.getContents().trim();
                    }
                }
            }
            //关闭workbook,防止内存泄露
            workbook.close();
        } catch (Exception e) {

        }
        return data;
    }
}
