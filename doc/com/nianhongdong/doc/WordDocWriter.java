/**
  *@(#) DOCWriter.java
  */
package com.nianhongdong.doc;
import java.io.File;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
   
/**  
 *  
 * 部分内容，转自网络   修改丰富 modified by nianhongdong
 * 
 * 作用：利用jacob插件生成word 文件！    
 *    
 *
 */   
public class WordDocWriter {    
   
    /** 日志记录器 */   
    static private Logger logger = Logger.getLogger(WordDocWriter.class);    
    
    /** word文档   
     *    
     * 在本类中有两种方式可以进行文档的创建,<br>   
     * 第一种调用 createNewDocument   
     * 第二种调用 openDocument    
     */     
    private Dispatch document = null;     
   
    /** word运行程序对象 */     
    private ActiveXComponent word = null;     
   
    /** 所有word文档 */     
    private Dispatch documents = null;     
   
    /**   
     *  Selection 对象 代表窗口或窗格中的当前所选内容。 所选内容代表文档中选定（或突出显示）的区域，如果文档中没有选定任何内容，则代表插入点。   
     *  每个文档窗格只能有一个Selection 对象，并且在整个应用程序中只能有一个活动的 Selection 对象。   
     */   
    private Dispatch selection = null;     
   
    /**   
     *    
     * Range 对象 代表文档中的一个连续区域。 每个 Range 对象由一个起始字符位置和一个终止字符位置定义。   
     * 说明：与书签在文档中的使用方法类似，Range 对象在 Visual Basic 过程中用来标识文档的特定部分。   
     * 但与书签不同的是，Range对象只在定义该对象的过程运行时才存在。   
     * Range对象独立于所选内容。也就是说，您可以定义和处理一个范围而无需更改所选内容。还可以在文档中定义多个范围，但每个窗格中只能有一个所选内容。   
     */   
    private Dispatch range = null;    
   
    /**   
     * PageSetup 对象 该对象包含文档的所有页面设置属性（如左边距、下边距和纸张大小）。   
     */   
    private Dispatch pageSetup = null;    
   
    /** 文档中的所有表格对象 */   
    private Dispatch tables = null;    
   
    /** 一个表格对象 */   
    private Dispatch table = null;    
   
    /** 表格所有行对象 */   
    private Dispatch rows = null;    
   
    /** 表格所有列对象 */   
    private Dispatch cols = null;    
   
    /** 表格指定行对象 */   
    private Dispatch row = null;    
   
    /** 表格指定列对象 */   
    private Dispatch col = null;    
   
    /** 表格中指定的单元格 */   
    private Dispatch cell = null;    
        
    /** 字体 */   
    private Dispatch font = null;    
        
    /** 对齐方式 */   
    private Dispatch alignment = null;    
   
    /**   
     * 构造方法   
     */   
    public WordDocWriter() {     
   
        if(this.word == null){    
            /* 初始化应用所要用到的对象实例 */   
            this.word = new ActiveXComponent("Word.Application");     
            /* 设置Word文档是否可见，true-可见false-不可见 */   
            this.word.setProperty("Visible",new Variant(false));    
            /* 禁用宏 */   
            this.word.setProperty("AutomationSecurity", new Variant(3));    
        }    
        if(this.documents == null){    
            this.documents = word.getProperty("Documents").toDispatch();    
        }    
    }    
   
    /**   
     * 设置页面方向和页边距   
     *   
     * @param orientation   
     *            可取值0或1，分别代表横向和纵向   
     * @param leftMargin   
     *            左边距的值   
     * @param rightMargin   
     *            右边距的值   
     * @param topMargin   
     *            上边距的值   
     * @param buttomMargin   
     *            下边距的值   
     */   
    public void setPageSetup(int orientation, int leftMargin,int rightMargin, int topMargin, int buttomMargin) {    
            
        logger.debug("设置页面方向和页边距...");    
        if(this.pageSetup == null){    
            this.getPageSetup();    
        }    
        Dispatch.put(pageSetup, "Orientation", ""+orientation);    
        Dispatch.put(pageSetup, "LeftMargin", ""+leftMargin);    
        Dispatch.put(pageSetup, "RightMargin", ""+rightMargin);    
        Dispatch.put(pageSetup, "TopMargin", ""+topMargin);    
        Dispatch.put(pageSetup, "BottomMargin", ""+buttomMargin);    
    } 
    
    /**
     * 
     * 功能说明:打印方向
     * 
     * @param sFlag  可取值0或1，分别代表横向和纵向  
     * @author nianhongdong
     * @date   2014-5-17 下午3:37:16
     * @see [相关类/方法]（可选）
     */
    public void setOrientation(String sFlag){
    	
    	Dispatch.put(pageSetup, "Orientation",sFlag);
    	
    }
    
    /**
     * 
     * 功能说明:左边距 
     * 
     * @param sFlag 单位厘米
     * @author nianhongdong
     * @date   2014-5-17 下午3:37:48
     * @see [相关类/方法]（可选）
     */
    public void setLeftMargin(String sFlag){
    	
    	Dispatch.put(pageSetup, "LeftMargin",sFlag);
    	
    }
 
 
   /**
    *  
    * 功能说明:右边距 
    * 
    * @param sFlag 单位厘米
    * @author nianhongdong
    * @date   2014-5-17 下午3:38:35
    * @see [相关类/方法]（可选）
    */
   public void setRightMargin(String sFlag){
 	
 	Dispatch.put(pageSetup, "RightMargin",sFlag);
 	
   }
   
   /**
    * 
    * 功能说明:上边距 
    * 
    * @param sFlag 单位厘米
    * @author nianhongdong
    * @date   2014-5-17 下午3:38:55
    * @see [相关类/方法]（可选）
    */
   public void setTopMargin(String sFlag){
	 	
	 	Dispatch.put(pageSetup, "TopMargin",sFlag);
	 	
   }
   
   /**
    * 
    * 功能说明:下边距 
    * 
    * @param sFlag 单位厘米
    * @author nianhongdong
    * @date   2014-5-17 下午3:39:27
    * @see [相关类/方法]（可选）
    */
   public void setBottomMargin(String sFlag){
	 	
	 	Dispatch.put(pageSetup, "BottomMargin",sFlag);
	 	
   }
    
    
   
    /**    
     * 打开文件    
     *    
     * @param inputDoc    
     *            要打开的文件，全路径    
     * @return Dispatch    
     *            打开的文件    
     */     
    public Dispatch openDocument(String inputDoc) {     
   
        logger.debug("打开Word文档...");    
        this.document = Dispatch.call(documents,"Open",inputDoc).toDispatch();    
        this.getSelection();    
        this.getRange();    
        this.getAlignment();    
        this.getFont();    
        this.getPageSetup();    
        return this.document;     
    }     
   
    /**   
     * 创建新的文件   
     *    
     * @return Dispache 返回新建文件   
     */   
    public Dispatch createNewDocument(){    
            
        logger.debug("创建新的文件...");    
        this.document = Dispatch.call(documents,"Add").toDispatch();    
        this.getSelection();    
        this.getRange();    
        this.getPageSetup();    
        this.getAlignment();    
        this.getFont();    
        return this.document;    
    }    
   
    /**    
     * 选定内容    
     * @return Dispatch 选定的范围或插入点    
     */     
    public Dispatch getSelection() {     
   
        logger.debug("获取选定范围的插入点...");    
        this.selection = word.getProperty("Selection").toDispatch();    
        return this.selection;     
    }     
   
    /**   
     * 获取当前Document内可以修改的部分<p><br>   
     * 前提条件：选定内容必须存在   
     *    
     * @param selectedContent 选定区域   
     * @return 可修改的对象   
     */   
    public Dispatch getRange() {    
   
        logger.debug("获取当前Document内可以修改的部分...");    
        this.range = Dispatch.get(this.selection, "Range").toDispatch();    
        return this.range;    
    }    
   
    /**   
     * 获得当前文档的文档页面属性   
     */   
    public Dispatch getPageSetup() {    
            
        logger.debug("获得当前文档的文档页面属性...");    
        if(this.document == null){    
            logger.warn("document对象为空...");    
            return this.pageSetup;    
        }    
        this.pageSetup = Dispatch.get(this.document, "PageSetup").toDispatch();    
        return this.pageSetup;    
    }    
   
    /**    
     * 把选定内容或插入点向上移动    
     * @param count 移动的距离    
     */     
    public void moveUp(int count) {     
   
        logger.debug("把选定内容或插入点向上移动...");    
        for(int i = 0;i < count;i++) {    
            Dispatch.call(this.selection,"MoveUp");    
        }    
    }     
    
    /**    
     * 把选定内容或插入点向下移动    
     * @param count 移动的距离    
     */     
    public void moveDown(int count) {     
   
        logger.debug("把选定内容或插入点向下移动...");    
        for(int i = 0;i < count;i++) {    
            Dispatch.call(this.selection,"MoveDown");    
        }    
    }
    
    /**
     * 
     * 功能说明:设置下一页
     * 
     * @author nianhongdong
     * @date   2014-5-17 下午4:29:49
     * @see [相关类/方法]（可选）
     */
    public void setNextPage() {     
    	   
        logger.debug("写入光标移到到下一页...");    
        Dispatch.call(this.selection,"InsertBreak");    
          
    }
    
 
   
    /**    
     * 把选定内容或插入点向左移动    
     * @param count 移动的距离    
     */     
    public void moveLeft(int count) {     
   
        logger.debug("把选定内容或插入点向左移动...");    
        for(int i = 0;i < count;i++) {    
            Dispatch.call(this.selection,"MoveLeft");    
        }    
    }     
   
    /**    
     * 把选定内容或插入点向右移动    
     * @param count 移动的距离    
     */     
    public void moveRight(int count) {     
   
        logger.debug("把选定内容或插入点向右移动...");    
        for(int i = 0;i < count;i++) {    
            Dispatch.call(this.selection,"MoveRight");    
        }    
    }    
        
    /**   
     * 回车键   
     */   
    public void enterDown(int count){    
            
        logger.debug("按回车键...");    
        for(int i = 0;i < count;i++) {    
            Dispatch.call(this.selection, "TypeParagraph");    
        }    
    }
    
    public void setEnterDown(String sCount){
    	
    	int count = Integer.parseInt(sCount);
    	logger.debug("按回车键...");    
        for(int i = 0;i < count;i++) {    
            Dispatch.call(this.selection, "TypeParagraph");    
        }  
    }
   
    /**    
     * 把插入点移动到文件首位置    
     */     
    public void moveStart() {     
   
        logger.debug("把插入点移动到文件首位置...");    
        Dispatch.call(this.selection,"HomeKey",new Variant(6));     
    }     
   
    /**    
     * 从选定内容或插入点开始查找文本    
     * @param selection 选定内容    
     * @param toFindText 要查找的文本    
     * @return boolean true-查找到并选中该文本，false-未查找到文本    
     */     
    public boolean find(String toFindText) {     
   
        logger.debug("从选定内容或插入点开始查找文本"+" 要查找内容：  "+toFindText);    
        /* 从selection所在位置开始查询 */   
        Dispatch find = Dispatch.call(this.selection,"Find").toDispatch();     
        /* 设置要查找的内容 */   
        Dispatch.put(find,"Text",toFindText);     
        /* 向前查找 */   
        Dispatch.put(find,"Forward","True");     
        /* 设置格式 */   
        Dispatch.put(find,"Format","True");     
        /* 大小写匹配 */   
        Dispatch.put(find,"MatchCase","True");     
        /* 全字匹配 */   
        Dispatch.put(find,"MatchWholeWord","True");     
        /* 查找并选中 */   
        return Dispatch.call(find,"Execute").getBoolean();     
    }     
   
    /**    
     * 把选定内容替换为设定文本    
     * @param selection 选定内容    
     * @param newText 替换为文本    
     */     
    public void replace(String newText) {     
   
        logger.debug("把选定内容替换为设定文本...");    
        /* 设置替换文本 */   
        Dispatch.put(this.selection,"Text",newText);     
    }     
   
    /**    
     * 全局替换    
     * @param selection 选定内容或起始插入点    
     * @param oldText 要替换的文本    
     * @param replaceObj 替换为文本   
     */     
    public void replaceAll(String oldText,Object replaceObj) {     
   
        logger.debug("全局替换...");    
        /* 移动到文件开头 */   
        moveStart();     
        /* 表格替换方式 */   
        String newText = (String) replaceObj;    
        /* 图片替换方式 */   
        if(oldText.indexOf("image") != -1 || newText.lastIndexOf(".bmp") != -1 || newText.lastIndexOf(".jpg") != -1 || newText.lastIndexOf(".gif") != -1){     
            while (find(oldText)) {     
                insertImage(newText);     
                Dispatch.call(this.selection,"MoveRight");     
            }     
            /* 正常替换方式 */   
        } else {    
            while (find(oldText)) {     
                replace(newText);     
                Dispatch.call(this.selection,"MoveRight");     
            }     
        }    
    }     
   
    /**    
     * 插入图片    
     * @param selection 图片的插入点    
     * @param imagePath 图片文件（全路径）    
     */     
    public void insertImage(String imagePath) {     
   
        logger.debug("插入图片...");    
        Dispatch.call(this.selection, "TypeParagraph");    
        Dispatch.call(Dispatch.get(this.selection,"InLineShapes").toDispatch(),"AddPicture",imagePath);     
    }     
   
    /**   
     * 合并表格   
     *   
     * @param selection 操作对象   
     * @param tableIndex 表格起始点   
     * @param fstCellRowIdx 开始行   
     * @param fstCellColIdx 开始列   
     * @param secCellRowIdx 结束行   
     * @param secCellColIdx 结束列   
     */   
    public void mergeCell(int tableIndex, int fstCellRowIdx, int fstCellColIdx, int secCellRowIdx, int secCellColIdx){    
   
        logger.debug("合并单元格...");    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return;    
        }    
        Dispatch fstCell = Dispatch.call(table, "Cell",new Variant(fstCellRowIdx), new Variant(fstCellColIdx)).toDispatch();    
        Dispatch secCell = Dispatch.call(table, "Cell",new Variant(secCellRowIdx), new Variant(secCellColIdx)).toDispatch();    
        Dispatch.call(fstCell, "Merge", secCell);    
    }    
   
    /**   
     * 想Table对象中插入数值<p>   
     *     参数形式：ArrayList<String[]>List.size()为表格的总行数<br>   
     *     String[]的length属性值应该与所创建的表格列数相同   
     *    
     * @param selection 插入点   
     * @param tableIndex 表格起始点   
     * @param list 数据内容   
     */   
    public void insertToTable(ArrayList list){    
   
        System.out.println("向Table对象中插入数据...");    
        logger.debug("向Table对象中插入数据...");    
        if(list == null || list.size() <= 0){    
            logger.warn("写出数据集为空...");    
            return;    
        }    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return;    
        }    
        for(int i = 0; i < list.size(); i++){    
            String[]  strs = (String[])list.get(i);    
            for(int j = 0; j<strs.length; j++){    
                /* 遍历表格中每一个单元格，遍历次数与所要填入的内容数量相同 */   
                Dispatch cell = this.getCell(i+1, j+1);    
                /* 选中此单元格 */   
                Dispatch.call(cell, "Select");    
                /* 写出内容到此单元格中 */   
                Dispatch.put(this.selection, "Text", strs[j]);    
                /* 移动游标到下一个位置 */   
            }    
            this.moveDown(1);    
        }    
        this.enterDown(1);    
    }  
    
    /**
     * 
     * 功能说明: 插入表格数据
     * 
     * @param lsData
     * @param HeaderNum
     * @author nianhongdong
     * @date   2014-4-10 下午6:16:31
     * @see [相关类/方法]（可选）
     */
    public void insertTableData(List lsData,int HeaderNum,String[] fields){
    	
    	System.out.println("向Table对象中插入数据...");    
        logger.debug("向Table对象中插入数据...");    
        if(lsData == null || lsData.size() <= 0){    
            logger.warn("写出数据集为空...");    
            return;    
        }    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return;    
        }
        if(fields == null||fields.length==0){    
            logger.warn("引用列为空...");    
            return;    
        }  
        
        for(int i = 0; i < lsData.size(); i++){ 
        	
            Map data = (Map)lsData.get(i);   
            for(int j = 0; j<fields.length; j++){    
                /* 遍历表格中每一个单元格，遍历次数与所要填入的内容数量相同 */   
                Dispatch cell = this.getCell(HeaderNum+i+1, j+1);    
                /* 选中此单元格 */   
                Dispatch.call(cell, "Select"); 
            
                /* 写出内容到此单元格中 */   
                Dispatch.put(this.selection, "Text", data.get(fields[j]));    
                /* 移动游标到下一个位置 */   
            } 
            this.moveDown(1); 
        }  
        this.enterDown(1);    
    }
    
    
   
    /**   
     * 在文档中正常插入文字内容   
     *    
     * @param selection 插入点   
     * @param list 数据内容   
     */   
    public void insertToDocument(List list){    
   
        logger.debug("向Document对象中插入数据...");    
        if(list == null || list.size() <= 0){    
            logger.warn("写出数据集为空...");    
            return;    
        }    
        if(this.document == null){    
            logger.warn("document对象为空...");    
            return;    
        }    
        
        for(int i=0;i<list.size();i++){
        		String str = (String)list.get(i);
        		 /* 写出至word中 */   
                //this.applyListTemplate(3, 2);    
                Dispatch.put(this.selection, "Text", str);    
                this.moveDown(1);    
                this.enterDown(1); 
        	}
              
            
    }    
    
    /**
     * 
     * 功能说明:插入文件
     * 
     * @param fileName
     * @author nianhongdong
     * @date   2014-4-11 下午1:57:56
     * @see [相关类/方法]（可选）
     */
    public void insertFileToDocument(String fileName){
    	
    	logger.debug("向Document对象文件...");  
    	//插入文件内容
    	Dispatch.call(this.selection,"insertFile",fileName);
    }
   
    /**   
     * 创建新的表格   
     *    
     * @param selection 插入点   
     * @param document 文档对象   
     * @param rowCount 行数   
     * @param colCount 列数   
     * @param width 边框数值 0浅色1深色   
     * @return 新创建的表格对象   
     */   
    public Dispatch createNewTable(int rowCount, int colCount, int width){    
   
        logger.debug("创建新的表格...");    
        if(this.tables == null){    
            this.getTables();    
        }    
        this.getRange();    
        if(rowCount > 0 && colCount > 0){    
            this.table = Dispatch.call(this.tables,"Add",this.range,new Variant(rowCount),new Variant(colCount),new Variant(width)).toDispatch();    
        }    
        /* 返回新创建表格 */   
        return this.table;    
    }    
    
    /**
     * 
     * 功能说明 ：创建表格 
     * 
     * @param rowCount
     * @param colCount
     * @param width
     * @return
     * @author nianhongdong
     * @date   2014-5-17 下午3:41:33
     * @see [相关类/方法]（可选）
     */
    public Dispatch createTable(String rowCount, String colCount, String width){    
    	   
        logger.debug("创建新的表格...");    
        if(this.tables == null){    
            this.getTables();    
        }    
        this.getRange(); 
        
        int iRowCount = Integer.parseInt(rowCount);
        int iColCount = Integer.parseInt(colCount);
        
        if(iRowCount > 0 && iColCount > 0){    
            this.table = Dispatch.call(this.tables,"Add",this.range,new Variant(rowCount),new Variant(colCount),new Variant(width)).toDispatch();    
        }    
        /* 返回新创建表格 */   
        return this.table;    
    }    
   
    /**   
     * 获取Document对象中的所有Table对象   
     *    
     * @return 所有Table对象   
     */   
    public Dispatch getTables(){    
   
   
        logger.debug("获取所有表格对象...");    
        if(this.document == null){    
            logger.warn("document对象为空...");    
            return this.tables;    
        }    
        this.tables = Dispatch.get(this.document, "Tables").toDispatch();    
        return this.tables;    
    }    
        
    /**   
     * 获取Document中Table的数量   
     *    
     * @return 表格数量   
     */   
    public int getTablesCount(){    
            
        logger.debug("获取文档中表格数量...");    
        if(this.tables == null){    
            this.getTables();    
        }    
        return Dispatch.get(tables, "Count").getInt();    
            
    }    
   
    /**   
     * 获取指定序号的Table对象   
     *    
     * @param tableIndex Table序列   
     * @return   
     */   
    public Dispatch getTable(int tableIndex){    
   
        logger.debug("获取指定表格对象...");    
        if(this.tables == null){    
            this.getTables();    
        }    
        if(tableIndex >= 0){    
            this.table = Dispatch.call(this.tables, "Item", new Variant(tableIndex)).toDispatch();    
        }    
        return this.table;    
    }    
   
    /**   
     * 获取表格的总列数   
     *    
     * @return 总列数   
     */   
    public int getTableColumnsCount() {    
   
        logger.debug("获取表格总行数...");    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return 0;    
        }    
        return Dispatch.get(this.cols,"Count").getInt();    
    }    
   
    /**   
     * 获取表格的总行数   
     *    
     * @return 总行数   
     */   
    public int getTableRowsCount(){    
   
        logger.debug("获取表格总行数...");    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return 0;    
        }    
        
      
        return Dispatch.get(this.rows,"Count").getInt();    
    }    
    /**   
     * 获取表格列对象   
     *    
     * @return 列对象   
     */   
    public Dispatch getTableColumns() {    
   
        logger.debug("获取表格行对象...");    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return this.cols;    
        }    
        this.cols = Dispatch.get(this.table,"Columns").toDispatch();    
        return this.cols;    
    }    
   
   
    /**   
     * 获取表格的行对象   
     *    
     * @return 总行数   
     */   
    public Dispatch getTableRows(){    
   
        logger.debug("获取表格总行数...");    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return this.rows;    
        }    
        this.rows = Dispatch.get(this.table,"Rows").toDispatch();    
        return this.rows;    
    }    
   
    /**   
     * 获取指定表格列对象   
     *    
     * @return 列对象   
     */   
    public Dispatch getTableColumn(int columnIndex) {    
   
        logger.debug("获取指定表格行对象...");    
        if(this.cols == null){    
            this.getTableColumns();    
        }    
        if(columnIndex >= 0){    
            this.col = Dispatch.call(this.cols, "Item", new Variant(columnIndex)).toDispatch();    
        }    
        return this.col;    
    }   

    /**
     * 
     * 在表格中增加行
     * 
     * @author nianhongdong
     * @date   2014-4-10 下午5:10:13
     * @see [相关类/方法]（可选）
     */
	public void addTableRow(int rowCout) {
		
		for(int i=0;i<rowCout;i++){
			Dispatch.call(this.getTableRows(), "Add"); 
		}
		
		
	}

    /**   
     * 获取表格中指定的行对象   
     *    
     * @param rowIndex 行序号   
     * @return 行对象   
     */   
    public Dispatch getTableRow(int rowIndex){    
   
        logger.debug("获取指定表格总行数...");    
        if(this.rows == null){    
            this.getTableRows();    
        }    
        if(rowIndex >= 0){    
            this.row = Dispatch.call(this.rows, "Item", new Variant(rowIndex)).toDispatch();    
        }    
        return this.row;    
    }    
        
    /**   
     * 自动调整表格   
     */   
    public void autoFitTable() {    
            
        logger.debug("自动调整表格...");    
        int count = this.getTablesCount();    
        for (int i = 0; i < count; i++) {    
            Dispatch table = Dispatch.call(tables, "Item", new Variant(i + 1)).toDispatch();    
            Dispatch cols = Dispatch.get(table, "Columns").toDispatch();    
            Dispatch.call(cols, "AutoFit");    
        }    
    }    
   
    /**   
     * 获取当前文档中，表格中的指定单元格   
     *   
     * @param CellRowIdx  单元格所在行   
     * @param CellColIdx 单元格所在列   
     * @return 指定单元格对象   
     */   
    public Dispatch getCell(int cellRowIdx, int cellColIdx) {    
   
            
        logger.debug("获取当前文档中，表格中的指定单元格...");    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return this.cell;    
        }    
        if(cellRowIdx >= 0 && cellColIdx >=0){    
            this.cell = Dispatch.call(this.table, "Cell", new Variant(cellRowIdx),new Variant(cellColIdx)).toDispatch();    
        }    
        return this.cell;    
    }    
   
    /**   
     * 设置文档标题   
     *    
     * @param title 标题内容   
     */   
    public void setTitle(String title){    
            
        logger.debug("设置文档标题...");    
        if(title == null || "".equals(title)){    
            logger.warn("文档标题为空...");    
            return;    
        }    
        Dispatch.call(this.selection, "TypeText", title);     
    }
    
    /** 
     * 
     * 功能说明:设置页面纸张大小 
     * 
     * @param sSize ：A3,A4,A5 可选
     * @author nianhongdong
     * @date   2014-5-17 下午3:42:12
     * @see [相关类/方法]（可选）
     */
    public void setPapersize(String sSize){
    	
    	String sPaperSize = "7";
    	if(sSize.equalsIgnoreCase("a3")){
    		sPaperSize = "6";
    	}else if(sSize.equalsIgnoreCase("a4")){
    		sPaperSize = "7";
    	}else if(sSize.equalsIgnoreCase("a5")){
    		sPaperSize = "9";
    	}

    	Dispatch.put(pageSetup, "PaperSize",sPaperSize);  
    	 
    }
    
 
        
    /**   
     * 设置当前表格线的粗细   
     *   
     * @param width   
     *        width范围：1<w<13,如果是0，就代表没有框   
     */   
    public void setTableBorderWidth(int width) {    
   
        logger.debug("设置当前表格线的粗细...");    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return;    
        }    
        /*   
         * 设置表格线的粗细 1：代表最上边一条线 2：代表最左边一条线 3：最下边一条线 4：最右边一条线 5：除最上边最下边之外的所有横线   
         * 6：除最左边最右边之外的所有竖线 7：从左上角到右下角的斜线 8：从左下角到右上角的斜线   
         */   
        Dispatch borders = Dispatch.get(table, "Borders").toDispatch();    
        Dispatch border = null;    
        for (int i = 1; i < 7; i++) {    
            border = Dispatch.call(borders, "Item", new Variant(i)).toDispatch();    
            if (width != 0) {    
                Dispatch.put(border, "LineWidth", new Variant(width));    
                Dispatch.put(border, "Visible", new Variant(true));    
            } else if (width == 0) {    
                Dispatch.put(border, "Visible", new Variant(false));    
            }    
        }    
    }    
        
    /**   
     * 对当前selection设置项目符号和编号   
     * @param tabIndex   
     *     1: 项目编号   
     *     2: 编号   
     *     3: 多级编号   
     *     4: 列表样式   
     * @param index   
     *     0:表示没有 ,其它数字代表的是该Tab页中的第几项内容   
     */   
    public void applyListTemplate(int tabIndex,int index){    
   
        logger.debug("对当前selection设置项目符号和编号...");    
        /* 取得ListGalleries对象列表 */   
        Dispatch listGalleries = Dispatch.get(this.word, "ListGalleries").toDispatch();    
        /* 取得列表中一个对象 */   
        Dispatch listGallery = Dispatch.call(listGalleries, "Item", new Variant(tabIndex)).toDispatch();    
        Dispatch listTemplates = Dispatch.get(listGallery, "ListTemplates").toDispatch();    
        if(this.range == null){    
            this.getRange();    
        }    
        Dispatch listFormat = Dispatch.get(this.range, "ListFormat").toDispatch();    
        Dispatch.call(listFormat,"ApplyListTemplate",Dispatch.call(listTemplates, "Item", new Variant(index)), new Variant(true),new Variant(1),new Variant(0));    
    }    
        
    /**   
     * 增加文档目录   
     *   
     * 目前采用固定参数方式，以后可以动态进行调整   
     */   
    public void addTablesOfContents()    
    {    
      /* 取得ActiveDocument、TablesOfContents、range对象 */   
      Dispatch ActiveDocument = word.getProperty("ActiveDocument").toDispatch();    
      Dispatch TablesOfContents = Dispatch.get(ActiveDocument,"TablesOfContents").toDispatch();    
      Dispatch range = Dispatch.get(this.selection, "Range").toDispatch();    
      /* 增加目录 */     
      Dispatch.call(TablesOfContents,"Add",range,new Variant(true),new Variant(1),new Variant(3),new Variant(true),new Variant(""),new Variant(true),new Variant(true));    
        
    }    
   
        
    /**   
     * 设置当前Selection 位置方式   
     * @param selectedContent 0－居左；1－居中；2－居右。   
     */   
    public void setAlignment(int alignmentType) {    
            
        logger.debug("设置当前Selection 位置方式...");    
        if(this.alignment == null){    
            this.getAlignment();    
        }    
        Dispatch.put(this.alignment, "Alignment", ""+alignmentType);    
    }    
        
    /**   
     * 获取当前选择区域的对齐方式   
     *    
     * @return 对其方式对象   
     */   
    public Dispatch getAlignment(){    
            
        logger.debug("获取当前选择区域的对齐方式...");    
        if(this.selection == null){    
            this.getSelection();    
        }    
        this.alignment = Dispatch.get(this.selection, "ParagraphFormat").toDispatch();    
        return this.alignment;    
    }    
        
    /**   
     * 获取字体对象   
     *    
     * @return 字体对象   
     */   
    public Dispatch getFont(){    
            
        logger.debug("获取字体对象...");    
        if(this.selection == null){    
            this.getSelection();    
        }    
        this.font = Dispatch.get(this.selection, "Font").toDispatch();    
        return this.font;    
    }    
        
    /**   
     * 设置选定内容的字体 注：在调用此方法前，选定区域对象selection必须存在   
     *   
     * @param fontName   
     *            字体名称，例如 "宋体"   
     * @param isBold   
     *            粗体   
     * @param isItalic   
     *            斜体   
     * @param isUnderline   
     *            下划线   
     * @param rgbColor   
     *            颜色，例如"255,255,255"   
     * @param fontSize   
     *            字体大小   
     * @param Scale   
     *            字符间距，百分比值。例如 70代表缩放为70%   
     */   
    public void setFontScale(String fontName, boolean isBold, boolean isItalic, boolean isUnderline,
    		String rgbColor, int Scale, int fontSize) {    
            
        logger.debug("设置字体...");    
        Dispatch.put(this.font, "Name", ""+fontName);    
        Dispatch.put(this.font, "Bold", new Variant(isBold));    
        Dispatch.put(this.font, "Italic", new Variant(isItalic));    
        Dispatch.put(this.font, "Underline", new Variant(isUnderline));    
        Dispatch.put(this.font, "Color", ""+rgbColor);    
        Dispatch.put(this.font, "Scaling", ""+Scale);    
        Dispatch.put(this.font, "Size", ""+fontSize);    
    }
    
    public void setFontName(String fontName){
    	
    	logger.debug("设置字体名称");
    	Dispatch.put(this.font, "Name", ""+fontName);    
    	 
    }
    
    public void setIsBold(String sFlag){
    	
    	logger.debug("设置字体粗体");
    	Dispatch.put(this.font, "Bold", new Variant(DocUtils.estimate(sFlag)));  
    	
    }
    
    public void setIsItalic(String sFlag){
    	
    	logger.debug("设置字体斜体");
    	Dispatch.put(this.font, "Italic", new Variant(DocUtils.estimate(sFlag)));  
    	
    }
    
    public void setIsUnderline(String sFlag){
    	
    	logger.debug("设置字体下划线");
    	Dispatch.put(this.font, "Underline", new Variant(DocUtils.estimate(sFlag)));  
    	
    }
    
    /**
     * 
     * 功能说明:设置字体颜色
     * 
     * @param sColor ： black:黑色 ,blue：蓝色,green：绿色,red：红色,yellow：黄色,brown：棕色。
     * @author nianhongdong
     * @date   2014-5-17 下午3:43:31
     * @see [相关类/方法]（可选）
     */
    public void setColor(String sColor){
    	
    	logger.debug("设置字体颜色");
    	String sColorValue = "0";
    	
    	if("black".equalsIgnoreCase(sColor)){
    		sColorValue = "0";
    	}else if("blue".equalsIgnoreCase(sColor)){
    		sColorValue = "16711680";
    	}else if("green".equalsIgnoreCase(sColor)){
    		sColorValue = "32768";
    	}else if("red".equalsIgnoreCase(sColor)){
    		sColorValue = "255";
    	}else if("yellow".equalsIgnoreCase(sColor)){
    		sColorValue = "65535";
    	}else if("brown".equalsIgnoreCase(sColor)){
    		sColorValue = "13209";
    	}
    	
    	Dispatch.put(this.font, "Color",sColorValue);  
    	
    }
    
    /**
     * 
     * 功能说明:字体大小
     * 
     * @param fontSize
     * @author nianhongdong
     * @date   2014-5-17 下午3:46:41
     * @see [相关类/方法]（可选）
     */
    public void setFontSize(String fontSize){
    	
    	logger.debug("设置字体大小");
    	Dispatch.put(this.font,"Size",fontSize);  
    }
    
        
    /**    
     * 保存文件    
     * @param outputPath 输出文件（包含路径）    
     */     
    public void saveAs(String outputPath) {     
   
        logger.debug("保存文件...");    
        if(this.document == null){    
            logger.warn("document对象为空...");    
            return;    
        }    
        if(outputPath ==null || "".equals(outputPath)){    
            logger.warn("文件保存路径为空...");    
            return;    
        }    
        Dispatch.call(this.document,"SaveAs",outputPath);     
    }    
        
    public void saveAsHtml(String htmlFile){    
        Dispatch.invoke(this.document,"SaveAs",Dispatch.Method, new Object[]{htmlFile,new Variant(8)}, new int[1]);    
    }    
   
    /**    
     * 关闭文件    
     * @param document 要关闭的文件    
     */     
    public void close() {    
   
        logger.debug("关闭文件...");    
        if(document == null){    
            logger.warn("document对象为空...");    
            return;    
        }    
        Dispatch.call(document,"Close",new Variant(0));     
    }    
   
    /**   
     * 列印word文件   
     *   
     */   
    public void printFile(){    
        logger.debug("打印文件...");    
        if(document == null){    
            logger.warn("document对象为空...");    
            return;    
        }    
        Dispatch.call(document,"PrintOut");    
    }    
   
    /**    
     * 退出程序    
     */     
    public void quit() {     
   
        logger.debug("退出程序");    
        word.invoke("Quit",new Variant[0]);     
        ComThread.Release();     
    }    
    
    /**
     * 
     * 功能说明:读取xml模板文件
     * 
     * @param fileName
     * @author nianhongdong
     * @date   2014-5-1 下午10:21:01
     * @see [相关类/方法]（可选）
     */
    public static void readXMLFile(String fileName,String saveFileName) throws Exception{
    	
    	SAXReader reader = new SAXReader();
    	Document docXML = reader.read(new File(fileName));
    	processXMLDocument(docXML,saveFileName);
    }
    
    /**
     * 
     * 功能说明: 读取xml字符串
     * 
     * @param sXML
     * @param saveFileName
     * @throws Exception
     * @author nianhongdong
     * @date   2014-5-17 下午3:47:06
     * @see [相关类/方法]（可选）
     */
    public static void readXMLString(String sXML,String saveFileName)throws Exception{
    	
    	Document docXML = DocumentHelper.parseText(sXML);
    	processXMLDocument(docXML,saveFileName);
    }
    
    /**
     * 
     * 功能说明:处理xml文档
     * 
     * @param docXML
     * @param saveFileName
     * @throws Exception
     * @author nianhongdong
     * @date   2014-5-17 下午3:47:36
     * @see [相关类/方法]（可选）
     */
    public static void processXMLDocument(Document docXML,String saveFileName) throws Exception{
    	
    	//获取根节点
    	Element rootElt = docXML.getRootElement(); 
        if(!"worddocument".equals(rootElt.getName())){
         	throw new RuntimeException("模板文档根节点名称不对！");
        }
        WordDocWriter docWriter = new WordDocWriter(); 
        docWriter.createNewDocument();
        Map mapMethod = getMethodMap(docWriter);
        
        //根节点属性设置
        setElementAttributs(docWriter,rootElt,mapMethod);
        
        Iterator iteEle = rootElt.elementIterator();
       
         while (iteEle.hasNext()) {
        	 Element recordEle = (Element)iteEle.next();
        	 setElementAttributs(docWriter,recordEle,mapMethod);
         }
         
        docWriter.saveAs(saveFileName); 
        docWriter.quit();
        docWriter.close();
    }
    
    /**
     * 
     * 功能说明:获取自身设置方法
     * 
     * @param docWriter
     * @return
     * @throws Exception
     * @author nianhongdong
     * @date   2014-5-17 下午3:48:24
     * @see [相关类/方法]（可选）
     */
    private static Map getMethodMap(WordDocWriter docWriter) throws Exception{
    	
    	Method[] methods = DocUtils.getAllMethod(docWriter);
    	Map mapMethod = new HashMap();
    	
    	for(int i=0;i<methods.length;i++){
    		if(methods[i].getName().startsWith("set")){
    			mapMethod.put(methods[i].getName().substring(3).toLowerCase(),
    					methods[i].getName());
    		}
    	}
    	
    	return mapMethod;
    }
    
   
    /** 
     * 
     * 功能说明:设置节点属性
     * 
     * @param rootElt
     * @author nianhongdong
     * @date   2014-5-1 下午10:28:12
     * @see [相关类/方法]（可选）
     */
    private static void setElementAttributs(Object objectClass,Element element,Map mapMethod) 
    		throws Exception{
    	
    	 if(element.attributeCount()>0 && "[worddocument][p][enterdown][table]".
    			 indexOf("["+element.getName().toLowerCase()+"]")>=0){
    		 
    		 
    		 List lsAttr = (List)element.attributes();
    		 
    		 //表格
    		 if("table".equalsIgnoreCase(element.getName())){
    			 
    			 String rowC = element.attributeValue("rowcount");
    			 String colC = element.attributeValue("colcount");
    			 
    			 if(rowC==null||!DocUtils.isNumber(rowC)){
    				 throw new RuntimeException("表格行参数没有或者非数值属性！");
    			 }
    			 
    			 if(colC==null||!DocUtils.isNumber(colC)){
    				 throw new RuntimeException("表格列参数没有或者非数值属性！");
    			 }
    			 
    			 DocUtils.invokeMethod(objectClass,"createTable",new Object[]{rowC,colC,"1"});
    			 
    			 Iterator itData = element.elementIterator("data");
    			 List lsTableData = new ArrayList();
    			 
    			 while(itData.hasNext()){
    				 Element itemEle = (Element) itData.next();
    				 String[] values = itemEle.attributeValue("values").split(",");
    				 String[] fixValues = new String[Integer.parseInt(colC)];
    			
    				 System.arraycopy(values,0,fixValues,0,Integer.parseInt(colC));
    				 
    				 if(lsTableData.size()+1 > Integer.parseInt(rowC)) continue;
    				 lsTableData.add(fixValues);
    				 
    			 }
    			 
    			 DocUtils.invokeMethod(objectClass,"insertToTable",lsTableData);
    		 }
    		 //回车
    		 if("enterdown".equalsIgnoreCase(element.getName())){ 
    			if(lsAttr!=null&&!lsAttr.isEmpty()){
    				Attribute at = (Attribute)lsAttr.get(0);
    				if("num".equalsIgnoreCase(at.getName())){
    					DocUtils.invokeMethod(objectClass,"setEnterDown",at.getValue());
    				}
    			}
    		 }
    		
         	//文本输出
         	String sTextContent = null;
         	
         	for(int i=0;i<lsAttr.size();i++){
         		Attribute at = (Attribute)lsAttr.get(i);
         		logger.debug("节点属性:"+at.getName()+"|"+"节点属性值:"+at.getValue());
         		if(DocUtils.isNullStr(at.getValue())) continue;
         		if("text".equalsIgnoreCase(at.getName())){
         			sTextContent = at.getValue();
         		}
         		if(mapMethod.containsKey(at.getName().toLowerCase())){
         			String sMethodName = (String)mapMethod.get(at.getName().toLowerCase());	
         			DocUtils.invokeMethod(objectClass,sMethodName,at.getValue());
         		}
         	}
         	
         	if(!DocUtils.isNullStr(sTextContent)){
         		DocUtils.invokeMethod(objectClass,"setTitle",sTextContent);
         	}
         	
         }else{
        	 
        	 if(mapMethod.containsKey(element.getName().toLowerCase())){
      			String sMethodName = (String)mapMethod.get(element.getName().toLowerCase());	
      			DocUtils.invokeMethod(objectClass,sMethodName,null);
      		}
         }
	}

    public static void main(String args[]) throws Exception{    
    	readXMLFile("d:/test.xml","d:/testNhd.docx");
    }    
   
}  