/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
Új sheet bejegyzése: 
[Content_Types].xml - ok
docProps/app.xml - ok
docProps/core.xml-be be lehet jegyezni a módosítás idejét, szerzõjét
xl/workbook.xml A sheets tagba kell bejegyezni. - ok
xl/styles.xml -be a cellák stílusait, ami a kirakáshoz kell
xl/_rels/workbook.xml.rels sheetet bejegyezni - ok
xl/worksheets mappába pedig létre kell hozni a sheetet, ha nincs felülírás

*/


//import org.w3c.dom.*;
package excelcasirtools;

import java.io.FileWriter;
import java.io.PrintWriter;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.LinkedList;
import java.util.Iterator;
import java.io.File;
import java.io.*;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.Date;
import java.text.SimpleDateFormat;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.io.SAXReader;
import org.dom4j.Element;
import org.dom4j.Attribute;
import org.dom4j.Node;
import org.dom4j.io.XMLWriter;
import org.dom4j.io.OutputFormat;
import java.io.IOException;
import java.io.Writer;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;
import java.nio.charset.StandardCharsets;

/**
 *
 * @author dell
 */
public class XlsxWriter {
    static String DATE="D";
    static String TEXT="C";
    static String NUM="N";
    static String TIME="T";
    static String DATETIME="DT";
    static String PERCENT="P";
    private String textFile="c:/tmp/forras.txt";
    private String outFile="c:/tmp/out.xlsx";
    private String workLibrary="c:/tmp";
    private String xlsxTemplate="c:/tmp/null.xlsx";
    private String separator="¤";
    private String logFile=null;
    private String charSet=null;
    private String sheetId=null;
    private String sheetName="LaciBacsi";
    private boolean logEnabled=true;
    private PrintWriter log=null;
    private boolean sheetOverwriting=false;
    private LinkedList<Element> sheetList;
    private boolean analogSys=false;
    private int fontCount=2,borderCount=1,fillCount=1,xfsCount=2,numFmtCount=2;
    private String[] columnNames,columnTypes;
    private OutputStreamWriter sheetFileWriter;
    XMLStreamWriter sheetWriter;
    private int rowCount=0,rowCounter=1,colCounter=1;
    static String ABC="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    private String title=null,subtitle=null;

    public static void main(String[] args) {
        // TODO code application logic here
     XlsxWriter xw=new XlsxWriter();   
     for (int i=0;i<args.length;i++){
        if (args[i].toLowerCase().equals("-title")) xw.title=args[i+1];
        if (args[i].toLowerCase().equals("-subtitle")) xw.subtitle=args[i+1];
        if (args[i].toLowerCase().equals("-filename")) xw.outFile=args[i+1];
        if (args[i].toLowerCase().equals("-rowcount")) xw.rowCount=Integer.parseInt(args[i+1]);
        if (args[i].toLowerCase().equals("-templatefile")) xw.xlsxTemplate=args[i+1];
	if (args[i].toLowerCase().equals("-logfile")) xw.logFile=args[i+1];
	if (args[i].toLowerCase().equals("-workpath")) xw.workLibrary=args[i+1];
	if (args[i].toLowerCase().equals("-charset")) xw.charSet=args[i+1];
	if (args[i].toLowerCase().equals("-separator")) xw.separator=args[i+1];
	if (args[i].toLowerCase().equals("-sheetname")) xw.sheetName=args[i+1];
	if (args[i].toLowerCase().equals("-log")) xw.logEnabled=args[i+1].toLowerCase().equals("yes") ? true:false;
     }
     xw.init();
    }
  public void setParams(String paramName,String paramValue){
        if (paramName.toLowerCase().equals("title")) this.title=paramValue;
        if (paramName.toLowerCase().equals("subtitle")) this.subtitle=paramValue;
        if (paramName.toLowerCase().equals("filename")) this.outFile=paramValue;
        if (paramName.toLowerCase().equals("rowcount")) this.rowCount=Integer.parseInt(paramValue);
        if (paramName.toLowerCase().equals("templatefile")) this.xlsxTemplate=paramValue;
	if (paramName.toLowerCase().equals("logfile")) this.logFile=paramValue;
	if (paramName.toLowerCase().equals("workpath")) this.workLibrary=paramValue;
	if (paramName.toLowerCase().equals("charset")) this.charSet=paramValue;
	if (paramName.toLowerCase().equals("separator")) this.separator=paramValue;
	if (paramName.toLowerCase().equals("sheetname")) this.sheetName=paramValue;
	if (paramName.toLowerCase().equals("log")) this.logEnabled=paramValue.toLowerCase().equals("yes") ? true:false;
       }  
  public void log(String msg){
      if (logEnabled){
      if (log!=null)log.println(msg);
      else System.out.println(msg);
      }
    }
  public void setColumnNames(String columnNames,String separator){
      this.columnNames=columnNames.split(separator);
  }
  public void setColumnTypes(String columnTypes,String separator){
      this.columnTypes=columnTypes.split(separator);
  }
  public void writeHead(){
      try {
       Writer fstream = null;
       BufferedWriter out = null;
    
    if (title!=null) rowCount=rowCount+2;
    if (subtitle!=null) rowCount=rowCount+2;
    sheetFileWriter = new OutputStreamWriter(new FileOutputStream(
            new File(workLibrary+"/XlsxWriter/xl/worksheets/sheet"+sheetId+".xml")),
            StandardCharsets.UTF_8);   
       XMLOutputFactory xMLOutputFactory = XMLOutputFactory.newInstance();
          sheetWriter =
            xMLOutputFactory.createXMLStreamWriter(sheetFileWriter);
 
         sheetWriter.writeStartDocument();
         sheetWriter.writeStartElement("worksheet");
         sheetWriter.writeAttribute("xmlns",
                 "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
         sheetWriter.writeAttribute("xmlns:r",
                 "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
         sheetWriter.writeAttribute("xmlns:mc",
                 "http://schemas.openxmlformats.org/markup-compatibility/2006");
         sheetWriter.writeAttribute("mc:Ignorable",
                 "x14ac");
         sheetWriter.writeAttribute("xmlns:x14ac",
                 "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
         sheetWriter.writeStartElement("sheetPr");
         sheetWriter.writeStartElement("pageSetUpPr");
         sheetWriter.writeAttribute("fitToPage","1");
         sheetWriter.writeEndElement();
         sheetWriter.writeEndElement();
         if (rowCount>0){
            sheetWriter.writeStartElement("dimension");
            sheetWriter.writeAttribute("ref","A1:"+getExcelColumn(columnNames.length)+rowCount);
            sheetWriter.writeEndElement();
         }
         sheetWriter.writeStartElement("sheetViews");
         sheetWriter.writeStartElement("sheetView");
         sheetWriter.writeAttribute("tabSelected","1");
         sheetWriter.writeAttribute("workbookViewId","0");
         sheetWriter.writeStartElement("selection");
         sheetWriter.writeAttribute("activeCell","A1");
         sheetWriter.writeAttribute("sqref","A1");
         sheetWriter.writeEndElement();
         sheetWriter.writeEndElement();
         sheetWriter.writeStartElement("sheetFormatPr");
         sheetWriter.writeAttribute("defaultRowHeight","15");
         sheetWriter.writeAttribute("x14ac:dyDescent","0.25");
         sheetWriter.writeEndElement();
         boolean kellCols=false;
         for (int i=1;i<=columnTypes.length;i++){
             if (columnTypes[i-1].equals(DATE)){
                 if (!kellCols){
                   sheetWriter.writeStartElement("cols"); 
                   kellCols=true;
                 }
                 sheetWriter.writeStartElement("col");
                 sheetWriter.writeAttribute("min",""+i);
                 sheetWriter.writeAttribute("max",""+i);
                 sheetWriter.writeAttribute("width","10.140625");
                 sheetWriter.writeAttribute("bestFit","1");
                 sheetWriter.writeAttribute("customWidth","1");
                 sheetWriter.writeEndElement();
             }
             if (columnTypes[i-1].equals(DATETIME)){
                 if (!kellCols){
                   sheetWriter.writeStartElement("cols"); 
                   kellCols=true;
                 }
                 sheetWriter.writeStartElement("col");
                 sheetWriter.writeAttribute("min",""+i);
                 sheetWriter.writeAttribute("max",""+i);
                 sheetWriter.writeAttribute("width","19.85546875");
                 sheetWriter.writeAttribute("bestFit","1");
                 sheetWriter.writeAttribute("customWidth","1");
                 sheetWriter.writeEndElement();
             }
         }
         if (kellCols) sheetWriter.writeEndElement();
         sheetWriter.writeStartElement("sheetData");
         boolean elvalasztosor=false;
         if (title!=null) {
             sheetWriter.writeStartElement("row");
             sheetWriter.writeAttribute("r",""+rowCounter);
             sheetWriter.writeAttribute("spans","1:1");
             sheetWriter.writeAttribute("x14ac:dyDescent","0.25");
             colCounter=0;
    	     sheetWriter.writeStartElement("c");
    	     sheetWriter.writeAttribute("r","A"+rowCounter);
    	     sheetWriter.writeAttribute("t","inlineStr");
    	     sheetWriter.writeAttribute("s",""+(xfsCount-5));
    	     sheetWriter.writeStartElement("is");
    	     sheetWriter.writeStartElement("t");
    	     sheetWriter.writeCharacters(title);
    	     sheetWriter.writeEndElement();
    	     sheetWriter.writeEndElement();
    	     sheetWriter.writeEndElement();
    	     sheetWriter.writeEndElement();
    	     rowCounter++;
    	     elvalasztosor=true;
    	     }
         if (subtitle!=null) {
             sheetWriter.writeStartElement("row");
             sheetWriter.writeAttribute("r",""+rowCounter);
             sheetWriter.writeAttribute("spans","1:1");
             sheetWriter.writeAttribute("x14ac:dyDescent","0.25");
             colCounter=0;
    	     sheetWriter.writeStartElement("c");
    	     sheetWriter.writeAttribute("r","A"+rowCounter);
    	     sheetWriter.writeAttribute("t","inlineStr");
    	     sheetWriter.writeAttribute("s",""+(xfsCount-4));
    	     sheetWriter.writeStartElement("is");
    	     sheetWriter.writeStartElement("t");
    	     sheetWriter.writeCharacters(subtitle);
    	     sheetWriter.writeEndElement();
    	     sheetWriter.writeEndElement();
    	     sheetWriter.writeEndElement();
    	     sheetWriter.writeEndElement();
    	     rowCounter++;
    	     elvalasztosor=true;
    	     }
         if (elvalasztosor) rowCounter++;
      }
      catch(Exception e){
          e.printStackTrace(log);
      }
  }
  
  public void startRow(){
  try{
      sheetWriter.writeStartElement("row");
      sheetWriter.writeAttribute("r",""+rowCounter);
      sheetWriter.writeAttribute("spans","1:"+(columnNames.length+1));
      sheetWriter.writeAttribute("x14ac:dyDescent","0.25");
      colCounter=0;
  }    
  catch(Exception e){
      e.printStackTrace(log);
  }    
  }

  public void endRow(){
  rowCounter++;
  try{
      sheetWriter.writeEndElement();
  }    
  catch(Exception e){
      e.printStackTrace(log);
  }    
  }
  public void writeDataCell(String textValue){
  try{
	     sheetWriter.writeStartElement("c");
	     sheetWriter.writeAttribute("r",getExcelColumn(colCounter)+rowCounter);
	     sheetWriter.writeAttribute("t","inlineStr");
	     sheetWriter.writeAttribute("s",""+(xfsCount-4));
	     sheetWriter.writeStartElement("is");
	     sheetWriter.writeStartElement("t");
	     sheetWriter.writeCharacters(textValue);
	     sheetWriter.writeEndElement();
	     sheetWriter.writeEndElement();
	     sheetWriter.writeEndElement();
	     colCounter++;
  }    
  catch(Exception e){
      e.printStackTrace(log);
  }    
  }
  
  public void writeDataCell(double numValue){
  try{

     sheetWriter.writeStartElement("c");
     sheetWriter.writeAttribute("r",getExcelColumn(colCounter)+rowCounter);
     if (columnTypes[colCounter].equals(DATE))sheetWriter.writeAttribute("s",""+(xfsCount-2));
     else if (columnTypes[colCounter].equals(DATETIME))sheetWriter.writeAttribute("s",""+xfsCount);
     else if (columnTypes[colCounter].equals(TIME))sheetWriter.writeAttribute("s",""+(xfsCount-1));
     sheetWriter.writeStartElement("v");
     if (!Double.isNaN(numValue)) sheetWriter.writeCharacters(""+numValue);
     sheetWriter.writeEndElement();
     sheetWriter.writeEndElement();
     colCounter++;
  }    
  catch(Exception e){
      e.printStackTrace(log);
  }    
  }
 
  public void closeTable(){
      try{
         sheetWriter.writeEndElement();
         sheetWriter.writeStartElement("pageMargins");
         sheetWriter.writeAttribute("left","0.7");
         sheetWriter.writeAttribute("right","0.7");
         sheetWriter.writeAttribute("top","0.75");
         sheetWriter.writeAttribute("bottom","0.75");
         sheetWriter.writeAttribute("header","0.3");
         sheetWriter.writeAttribute("footer","0.3");
         sheetWriter.writeEndElement();
         sheetWriter.writeStartElement("pageSetup");
         sheetWriter.writeAttribute("paperSize","9");
         sheetWriter.writeAttribute("orientation","portrait");
         sheetWriter.writeAttribute("verticalDpi","0");
         sheetWriter.writeAttribute("r:id","rId1");
         sheetWriter.writeEndElement();
         sheetWriter.writeEndElement();
         
         sheetWriter.writeEndDocument();

         sheetWriter.flush();
         sheetWriter.close();
      }
      catch(Exception e){
          e.printStackTrace(log);
      }
  }
  public void init(){
      try {
      if (logFile!=null)log=new PrintWriter(new FileWriter(logFile));
      else log=new PrintWriter(System.out);
      File workDir = new File(workLibrary+"/XlsxWriter");
        if (!workDir.exists()){
        workDir.mkdirs();
        }
      unzipExcelFile(xlsxTemplate);
      xl_workbook();
      docProps_core();
      if (!sheetOverwriting){
          Content_Types();
          xl_rels_workbook();
          docProps_app();
          xl_styles();
      }
      if (log!=null){
          log.flush();
          log.close();
      }
      }
      catch(Exception e){
          e.printStackTrace(log);
      }
  }  
  private void createDirectoryFromFilePath(String path){
      String[] darabok=path.split("/");
      String ut=workLibrary+"/XlsxWriter";
      if (darabok.length>1){
          for (int i=0;i<darabok.length-1;i++){
              ut=ut+"/"+darabok[i];
              File f=new File(ut);
              if (!f.exists()) f.mkdirs();
            } 
        }
  }
  private void unzipExcelFile(String excelFile){
      Enumeration enumEntries;
      ZipFile zip;
      String fileName="";

      try {
          zip = new ZipFile(excelFile);
          enumEntries = zip.entries();
          while (enumEntries.hasMoreElements()) {
              ZipEntry zipentry = (ZipEntry) enumEntries.nextElement();
              fileName= zipentry.getName();
              createDirectoryFromFilePath(fileName);
              if (zipentry.isDirectory()) {
                  fileName= zipentry.getName();
                  log("Name of Extract directory : " +fileName);
                  File f=new File(workLibrary+"/XlsxWriter/"+fileName);
                  if (!f.exists()) f.mkdirs();
                  continue;
              }
              log("Name of Extract fille : " + zipentry.getName());
              extractFile(zip.getInputStream(zipentry), new FileOutputStream(workLibrary+"/XlsxWriter/"+fileName));
          }
          zip.close();
     } catch (IOException ioe) {
         log("There is an IoException Occured :" + ioe);
         ioe.printStackTrace(log);
     }
  }  
  
  private void extractFile(InputStream inStream, OutputStream outStream) throws IOException {
      byte[] buf = new byte[1024];
      int l;
      while ((l = inStream.read(buf)) >= 0) {
           outStream.write(buf, 0, l);
      }
      inStream.close();
      outStream.close();
  }
  /*
  - <HeadingPairs>
- <vt:vector size="2" baseType="variant">
- <vt:variant>
  <vt:lpstr>Munkalapok</vt:lpstr> 
  </vt:variant>
- <vt:variant>
  <vt:i4>2</vt:i4> 
  </vt:variant>
  </vt:vector>
  </HeadingPairs>
- <TitlesOfParts>
- <vt:vector size="2" baseType="lpstr">
  <vt:lpstr>work1</vt:lpstr> 
  <vt:lpstr>job</vt:lpstr> 
  </vt:vector>
  </TitlesOfParts>
  */
  private void docProps_app(){
    try{
    List<Element> elements;
    File ct=new File(workLibrary+"/XlsxWriter/docProps/app.xml");
    SAXReader reader = new SAXReader();
    Document document = reader.read(ct);
    Element root = document.getRootElement();
    log("RootElement: "+root.getName());
    Element headingPairs=root.element("HeadingPairs");
    Element titlesOfParts=root.element("TitlesOfParts");
    Element vt_vector=titlesOfParts.element("vector");
    Attribute sizeA=vt_vector.attribute("size");
    int size=Integer.parseInt(sizeA.getValue());
    size++;
    vt_vector.remove(sizeA);
    vt_vector.addAttribute("size",""+size);
    Element bejegyzes=vt_vector.addElement("vt:lpstr").addText(sheetName);
    
    Element vtvector=headingPairs.element("vector");
    elements=vtvector.elements("variant");
    Element i4=null,variant=null;
    for (int i=0;i<elements.size();i++){
        if(elements.get(i).element("i4")!=null)variant=elements.get(i);
    }
    i4=variant.element("i4");
    size=Integer.parseInt(i4.getText());
    size++;
    variant.remove(i4);
    Element i_4=variant.addElement("vt:i4");
    i_4.addText(""+size);
    OutputFormat format = OutputFormat.createPrettyPrint();
    XMLWriter writer;
    FileWriter fw=new FileWriter(ct);
    writer = new XMLWriter(fw,format);
    writer.write( document );
    writer.close();
    }
    catch(Exception e){
        e.printStackTrace(log);
    } 
  }
/*
  <sheets>
  <sheet name="work1" sheetId="1" r:id="rId1" /> 
  <sheet name="job" sheetId="2" r:id="rId2" /> 
  </sheets>
*/ 
private void xl_workbook(){
    sheetList=new LinkedList();
    try{
    File ct=new File(workLibrary+"/XlsxWriter/xl/workbook.xml");
    SAXReader reader = new SAXReader();
    Document document = reader.read(ct);
    Element root = document.getRootElement();
    log("RootElement: "+root.getName());
    Element sheets=root.element("sheets");
    Element sheet=null;
    int id=1;
    for (Iterator<Element> sit = sheets.elementIterator("sheet"); sit.hasNext();) {
        sheet= sit.next();
        id++;
        if (sheet.attribute("name").getValue().equalsIgnoreCase(sheetName)){
            sheetName=sheet.attribute("name").getValue();
            sheetId=sheet.attribute("sheetId").getValue();
            log("sheetName: "+sheetName+" will be overwritten.");
            sheetOverwriting=true;
        }
    }
    if (sheetId==null)sheetId=""+id;
    if (!sheetOverwriting){
        Element bejegyzes=sheets.addElement("sheet")
            .addAttribute("name",sheetName)
            .addAttribute("sheetId", sheetId)
            .addAttribute("r:id","rId"+sheetId);
        OutputFormat format = OutputFormat.createPrettyPrint();
        XMLWriter writer;
        FileWriter fw=new FileWriter(ct);
        writer = new XMLWriter(fw,format);
        writer.write( document );
        writer.close();
        }
    }
    catch(Exception e){
        e.printStackTrace(log);
    }
}
/*
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" /> 
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml" /> 
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" /> 
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" /> 
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" /> 
  </Relationships>
*/
private void xl_rels_workbook(){
    try{
    File ct=new File(workLibrary+"/XlsxWriter/xl/_rels/workbook.xml.rels");
    SAXReader reader = new SAXReader();
    Document document = reader.read(ct);
    Element root = document.getRootElement();
    for (Iterator<Element> it = root.elementIterator("Relationship"); it.hasNext();) {
        Element relationShip=it.next();
        if (!relationShip.attribute("Target").getValue().contains("worksheets/sheet"))
        {
            Attribute rId=relationShip.attribute("Id");
            relationShip.remove(rId);
            int sorszam=Integer.parseInt(rId.getValue().substring(3));
            sorszam++;
            relationShip.addAttribute("Id","rId"+sorszam);
        }
        }    
    Element bejegyzes=root.addElement("Relationship")
        .addAttribute("Id","rId"+sheetId)
        .addAttribute("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
        .addAttribute("Target","worksheets/sheet"+sheetId+".xml");
    
        OutputFormat format = OutputFormat.createPrettyPrint();
        XMLWriter writer;
        FileWriter fw=new FileWriter(ct);
        writer = new XMLWriter(fw,format);
        writer.write( document );
        writer.close();
    }
    catch(Exception e){
        e.printStackTrace(log);
    }
}
/*
  <Override PartName="/xl/worksheets/sheet1.xml" 
  ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" /> 

  */

private void Content_Types(){
    try{
    File ct=new File(workLibrary+"/XlsxWriter/[Content_Types].xml");
    SAXReader reader = new SAXReader();
    Document document = reader.read(ct);
    Element root = document.getRootElement();
    log("RootElement: "+root.getName());
    Element bejegyzes=root.addElement("Override")
            .addAttribute("PartName","/xl/worksheets/sheet"+sheetId+".xml")
            .addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    OutputFormat format = OutputFormat.createPrettyPrint();
    XMLWriter writer;
    FileWriter fw=new FileWriter(ct);
    writer = new XMLWriter(fw,format);
    writer.write( document );
    writer.close();
    }
    catch(Exception e){
        e.printStackTrace(log);
    }
}  
/*
  <?xml version="1.0" encoding="UTF-8" standalone="yes" ?> 
- <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Apache POI</dc:creator> 
  <cp:lastModifiedBy>dell</cp:lastModifiedBy> 
  <dcterms:created xsi:type="dcterms:W3CDTF">2021-02-02T08:00:56Z</dcterms:created> 
  <dcterms:modified xsi:type="dcterms:W3CDTF">2021-02-02T21:33:07Z</dcterms:modified> 
  </cp:coreProperties>
*/

private void docProps_core(){
    try{
    File ct=new File(workLibrary+"/XlsxWriter/docProps/core.xml");
    SAXReader reader = new SAXReader();
    Document document = reader.read(ct);
    Element root = document.getRootElement();
    log("RootElement: "+root.getName());
    Element creator=root.element("creator");
    if(creator.getText().equals("AnalogSys")) analogSys=true;
    if (!analogSys){
    root.remove(creator);
    Element bejegyzes=root.addElement("dc:creator");
    bejegyzes.addText("AnalogSys");
    }
    System.getProperty("user.name");
    Element lastModifiedBy=root.element("lastModifiedBy");
    root.remove(lastModifiedBy);
    Element lmb=root.addElement("cp:lastModifiedBy");
    lmb.addText(System.getProperty("user.name"));
    Date date = new Date();  
    SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ssz");  
    Element modify=root.addElement("dcterms:modified")
    .addAttribute("xsi:type","dcterms:W3CDTF");
    modify.addText(formatter.format(date));
    OutputFormat format = OutputFormat.createPrettyPrint();
    XMLWriter writer;
    FileWriter fw=new FileWriter(ct);
    writer = new XMLWriter(fw,format);
    writer.write( document );
    writer.close();
    }
    catch(Exception e){
        e.printStackTrace(log);
    } 
    
}

private void xl_styles(){
    try{
       
    File ct=new File(workLibrary+"/XlsxWriter/xl/styles.xml");
    SAXReader reader = new SAXReader();
    Document document = reader.read(ct);
    Element root = document.getRootElement();
    log("RootElement: "+root.getName());
    Element borders=root.element("borders");
    borderCount=Integer.parseInt(borders.attributeValue("count"));
    borderCount++;
    borders.remove(borders.attribute("count"));
    borders.addAttribute("count",""+borderCount);
    Element border=borders.addElement("border");
    Element left=border.addElement("left").addAttribute("style","thin");
    Element right=border.addElement("right").addAttribute("style","thin");
    Element top=border.addElement("top").addAttribute("style","thin");
    Element bottom=border.addElement("bottom").addAttribute("style","thin");
    left.addElement("color").addAttribute("auto","1");
    right.addElement("color").addAttribute("auto","1");
    top.addElement("color").addAttribute("auto","1");
    bottom.addElement("color").addAttribute("auto","1");
    border.addElement("diagonal");
    Element fills=root.element("fills");
    //if (fills==null)fills=root.addElement("fills").addAttribute("count","1");
    fillCount=Integer.parseInt(fills.attributeValue("count"));
    fillCount++;
    fills.remove(fills.attribute("count"));
    fills.addAttribute("count",""+fillCount);
    Element fill=fills.addElement("fill");
    Element patternFill=fill.addElement("patternFill").addAttribute("patternType","none");
    patternFill.addElement("bgColor").addAttribute("indexed","22");
    Element fonts=root.element("fonts");
    if (fonts==null)fonts=root.addElement("fonts").addAttribute("count","2")
            .addAttribute("x14ac:knownFonts","1");
    else{
        fontCount=Integer.parseInt(fonts.attributeValue("count"));
        fontCount=fontCount+3;
        fonts.remove(fonts.attribute("count"));
        fonts.addAttribute("count",""+fontCount);
    }
    Element font=fonts.addElement("font");
    font.addElement("b");
    font.addElement("sz").addAttribute("val","14");
    font.addElement("color").addAttribute("indexed","8");
    font.addElement("name").addAttribute("val","Arial");
    font=fonts.addElement("font");
    font.addElement("b");
    font.addElement("sz").addAttribute("val","10");
    font.addElement("color").addAttribute("indexed","8");
    font.addElement("name").addAttribute("val","Arial");
    font=fonts.addElement("font");
    font.addElement("sz").addAttribute("val","10");
    font.addElement("color").addAttribute("indexed","8");
    font.addElement("name").addAttribute("val","Arial");
    Element numFmts=root.element("numFmts");
    if (numFmts==null)numFmts=root.addElement("numFmts").addAttribute("count","2");
    else {
        numFmtCount=Integer.parseInt(numFmts.attributeValue("count"));
        numFmtCount++;
        numFmtCount++;
        numFmts.remove(numFmts.attribute("count"));
        numFmts.addAttribute("count",""+numFmtCount);
    }
    int id=163,fmtid;
    for (Iterator<Element> it = numFmts.elementIterator("numFmt"); it.hasNext();) {
        Element fmt=it.next();
        fmtid=Integer.parseInt(fmt.attributeValue("numFmtId"));
        if (id<fmtid)id=fmtid;
        }    
    id++;
    Element numFmt=numFmts.addElement("numFmt").addAttribute("numFmtId",""+id)
            .addAttribute("formatCode","[$-F400]h:mm:ss\\ AM/PM");
    id++;
    numFmt=numFmts.addElement("numFmt").addAttribute("numFmtId",""+id)
            .addAttribute("formatCode","yyyy/mm/dd\\ hh:mm:ss");
    Element xfs=root.element("cellXfs");
    xfsCount=Integer.parseInt(xfs.attributeValue("count"));
    xfsCount=xfsCount+5;
    xfs.remove(xfs.attribute("count"));
    xfs.addAttribute("count",""+xfsCount);
    //Cím cella
    Element xf=xfs.addElement("xf").addAttribute("numFmtId","0")
            .addAttribute("fontId",""+(fontCount-3)) //bold
            .addAttribute("fillId","0") //default
            .addAttribute("borderId","0") //default
            .addAttribute("xfId","0")
            .addAttribute("applyBorder","0");
  //Fejléc cella
    xf=xfs.addElement("xf").addAttribute("numFmtId","0")
            .addAttribute("fontId",""+(fontCount-2)) //bold
            .addAttribute("fillId",""+(fillCount-1)) //szürke
            .addAttribute("borderId",""+(borderCount-1)) //keretezett
            .addAttribute("xfId","0")
            .addAttribute("applyBorder","1");
    //xf.addElement("alignment").addAttribute("horizontal","center"); //középre igazított
    // Adatcella String vagy sima szám
    xf=xfs.addElement("xf").addAttribute("numFmtId","0")
            .addAttribute("fontId",""+(fontCount-1)) //normal
            .addAttribute("fillId","0") //default
            .addAttribute("borderId",""+(borderCount-1)) //keretezett
            .addAttribute("xfId","0")
            .addAttribute("applyBorder","1");
    //xf.addElement("alignment").addAttribute("horizontal","left"); //balra igazított
    // Adatcella dátum
    xf=xfs.addElement("xf").addAttribute("numFmtId","14") //yyyy-mm-dd
            .addAttribute("fontId",""+(fontCount-1)) //normal
            .addAttribute("fillId","0") //default
            .addAttribute("borderId",""+(borderCount-1)) //keretezett
            .addAttribute("xfId","0")
            .addAttribute("applyNumberFormat","1")
            .addAttribute("applyBorder","1");
    //xf.addElement("alignment").addAttribute("horizontal","right"); //jobbra igazított
    // Adatcella idõ
    xf=xfs.addElement("xf").addAttribute("numFmtId",""+(id-1)) //hh:mm:ss
            .addAttribute("fontId",""+(fontCount-1)) //normal
            .addAttribute("fillId","0") //default
            .addAttribute("borderId",""+(borderCount-1)) //keretezett
            .addAttribute("xfId","0")
            .addAttribute("applyNumberFormat","1")
            .addAttribute("applyBorder","1");
    //xf.addElement("alignment").addAttribute("horizontal","right"); //jobbra igazított
    // Adatcella dátumidõ
    xf=xfs.addElement("xf").addAttribute("numFmtId",""+(id)) //yyyy-mm-dd hh:mm:ss
            .addAttribute("fontId",""+(fontCount-1)) //normal
            .addAttribute("fillId","0") //default
            .addAttribute("borderId",""+(borderCount-1)) //keretezett
            .addAttribute("xfId","0")
            .addAttribute("applyNumberFormat","1")
            .addAttribute("applyBorder","1");
    //xf.addElement("alignment").addAttribute("horizontal","right"); //jobbra igazított
    OutputFormat format = OutputFormat.createPrettyPrint();
    XMLWriter writer;
    FileWriter fw=new FileWriter(ct);
    writer = new XMLWriter(fw,format);
    writer.write( document );
    writer.close();
    }
    catch(Exception e){
        e.printStackTrace(log);
    } 
}
  private String getExcelColumn(int pos){
      int tizesek,szazasok,maradek;
      int base=ABC.length();
      if(pos<base) return ""+ABC.charAt(pos);
      tizesek=(int)(pos/base);
      if (tizesek<ABC.length()){
          maradek=pos-(tizesek*base);
          return ""+ABC.charAt(tizesek)+ABC.charAt(maradek);
      }
      szazasok=(int)(pos/(base*base));
      maradek=(int)(pos-(szazasok*base*base));
      tizesek=(int)maradek/base;
      maradek=maradek-(tizesek*base);
      return ""+ABC.charAt(szazasok)+ABC.charAt(tizesek)+ABC.charAt(maradek);
      }
      
      int power(int mit,int mire){
          int value=1;
          for (int x=0;x<mire;x++)value=value*mit;
          //log("mit: "+mit+" mire: "+mire+" value:"+value);
          return value;
      }

}
