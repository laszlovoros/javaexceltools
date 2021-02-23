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
import java.io.BufferedReader;
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
    private String outFile=null;
    private String workLibrary=null;
    private String xlsxTemplate=null;
    private String separator="¤";
    private String logFile=null;
    private String charSet=null;
    private String sheetId=null;
    private String sheetName=null;
    private boolean logEnabled=true;
    private PrintWriter log=null;
    private boolean sheetOverwriting=false,generatedByThis=false;
    private LinkedList<Element> sheetList;
    private boolean analogSys=false;
    private int fontCount=2,borderCount=1,fillCount=1,xfsCount=2,numFmtCount=2;
    private String[] columnNames=null,columnTypes=null;
    private OutputStreamWriter sheetFileWriter;
    private XMLStreamWriter sheetWriter;
    private BufferedReader input=null;
    private int rowCount=0,rowCounter=1,columnCount=0,colCounter=1;
    static String ABC="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    private String title=null,subtitle=null;

    public static void main(String[] args) {
        // TODO code application logic here
     XlsxWriter xw=new XlsxWriter();   
     for (int i=0;i<args.length;i++){
        if (args[i].toLowerCase().equals("-title")) xw.title=args[i+1];
        if (args[i].toLowerCase().equals("-subtitle")) xw.subtitle=args[i+1];
        if (args[i].toLowerCase().equals("-filename")) xw.outFile=args[i+1];
        if (args[i].toLowerCase().equals("-columncount")) xw.columnCount=Integer.parseInt(args[i+1]);
        if (args[i].toLowerCase().equals("-rowcount")) xw.rowCount=Integer.parseInt(args[i+1]);
        if (args[i].toLowerCase().equals("-templatefile")) xw.xlsxTemplate=args[i+1];
	if (args[i].toLowerCase().equals("-logfile")) xw.logFile=args[i+1];
	if (args[i].toLowerCase().equals("-workpath")) xw.workLibrary=args[i+1];
	if (args[i].toLowerCase().equals("-datasourcefile")) xw.textFile=args[i+1];
        if (args[i].toLowerCase().equals("-charset")) xw.charSet=args[i+1];
	if (args[i].toLowerCase().equals("-separator")) xw.separator=args[i+1];
	if (args[i].toLowerCase().equals("-sheetname")) xw.sheetName=args[i+1];
	if (args[i].toLowerCase().equals("-log")) xw.logEnabled=args[i+1].toLowerCase().equals("yes") ? true:false;
     }
     xw.init();
    }
  public void setParams(String[] params){
      String paramName=params[0];
      String paramValue=null;
      if (params.length>1)paramValue=params[1];
      log(paramName+":"+paramValue);
      if (paramName.toLowerCase().equals("title")) this.title=paramValue;
      if (paramName.toLowerCase().equals("subtitle")) this.subtitle=paramValue;
      if (paramName.toLowerCase().equals("filename")) this.outFile=paramValue;
      if (paramName.toLowerCase().equals("rowcount")) this.rowCount=Integer.parseInt(paramValue);
      if (paramName.toLowerCase().equals("columncount")) this.columnCount=Integer.parseInt(paramValue);
      if (paramName.toLowerCase().equals("templatefile")) this.xlsxTemplate=paramValue;
      if (paramName.toLowerCase().equals("logfile")) this.logFile=paramValue;
      if (paramName.toLowerCase().equals("datasourcefile")) this.textFile=paramValue;
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
  public void startSheet() throws XMLStreamException,IOException{
      
    Writer fstream = null;
    BufferedWriter out = null;
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
             else if (columnTypes[i-1].equals(DATETIME)){
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
             else{
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
            // sheetWriter.writeAttribute("spans","1:"+columnNames.length);
            // sheetWriter.writeAttribute("x14ac:dyDescent","0.25");
             colCounter=0;
    	     sheetWriter.writeStartElement("c");
    	     sheetWriter.writeAttribute("r","A"+rowCounter);
    	     sheetWriter.writeAttribute("t","inlineStr");
    	     sheetWriter.writeAttribute("s",""+(xfsCount-6));
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
             //sheetWriter.writeAttribute("spans","1:"+columnNames.length);
             //sheetWriter.writeAttribute("x14ac:dyDescent","0.25");
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
         startRow();
         for (int i=0;i<columnNames.length;i++) writeCell(columnNames[i],columnTypes[i],true);
         endRow();
  }
  public void processSheet(){
      
      try {
          columnNames=new String[columnCount];
          columnTypes=new String[columnCount];
          for (int i=0;i<columnCount;i++){
              columnNames[i]=input.readLine();
              columnTypes[i]=input.readLine();
          }
          if (title!=null)rowCount=rowCount+2;
          if (subtitle!=null)rowCount++;
          startSheet();
          writeSheet();
          endSheet();
      }
      catch(Exception e){
          e.printStackTrace();
      }
  }
  public void writeSheet() throws XMLStreamException,IOException{
    String line=input.readLine();
    while((line!=null)&&(rowCounter<=rowCount)){
        startRow();
        for(int i=0;i<columnCount;i++){
            line=input.readLine();
            if (line!=null)
            writeCell(line,columnTypes[i],false);
            else i=columnCount;
        }
        endRow();
    }
  }
  
  public void startRow() throws XMLStreamException{
  
      sheetWriter.writeStartElement("row");
      sheetWriter.writeAttribute("r",""+rowCounter);
      //sheetWriter.writeAttribute("spans","1:"+(columnNames.length));
      //sheetWriter.writeAttribute("x14ac:dyDescent","0.25");
      colCounter=0;
      
  }

  public void endRow() throws XMLStreamException{
  rowCounter++;
  sheetWriter.writeEndElement();
  }

  public void writeCell(String textValue,String fieldType,boolean header)
  throws XMLStreamException
  {
 
     sheetWriter.writeStartElement("c");
     sheetWriter.writeAttribute("r",getExcelColumn(colCounter)+rowCounter);
     if (fieldType.equals(TEXT)){
	     sheetWriter.writeAttribute("t","inlineStr");
	     if (header)sheetWriter.writeAttribute("s",""+(xfsCount-5));
             else sheetWriter.writeAttribute("s",""+(xfsCount-4));
	     sheetWriter.writeStartElement("is");
	     sheetWriter.writeStartElement("t");
	     sheetWriter.writeCharacters(textValue);
	     sheetWriter.writeEndElement();
     }
     else {
     if (fieldType.equals(DATE))sheetWriter.writeAttribute("s",""+(xfsCount-3));
     else if (fieldType.equals(DATETIME))sheetWriter.writeAttribute("s",""+(xfsCount-1));
     else if (fieldType.equals(TIME))sheetWriter.writeAttribute("s",""+(xfsCount-2));
     sheetWriter.writeStartElement("v");
     sheetWriter.writeCharacters(""+getDoubleValue(textValue));
     }
     sheetWriter.writeEndElement();
     sheetWriter.writeEndElement();
     colCounter++;
   }
  
  public void writeDataCell(double numValue){
  try{

     sheetWriter.writeStartElement("c");
     sheetWriter.writeAttribute("r",getExcelColumn(colCounter)+rowCounter);
     if (columnTypes[colCounter].equals(DATE))sheetWriter.writeAttribute("s",""+(xfsCount-3));
     else if (columnTypes[colCounter].equals(DATETIME))sheetWriter.writeAttribute("s",""+(xfsCount-1));
     else if (columnTypes[colCounter].equals(TIME))sheetWriter.writeAttribute("s",""+(xfsCount-2));
     sheetWriter.writeStartElement("v");
     if (!Double.isNaN(numValue)) sheetWriter.writeCharacters(""+numValue);
     sheetWriter.writeEndElement();
     sheetWriter.writeEndElement();
     colCounter++;
  }    
  catch(Exception e){
      e.printStackTrace();
  }    
  }
 
  public void endSheet() throws XMLStreamException{

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
  public void init(){
      try {
      System.out.println("initializing...");
      if (textFile==null)
      input = new BufferedReader(new InputStreamReader(System.in));
      else input=new BufferedReader(new FileReader(new File(textFile)));
      if (outFile==null){
          for(String line=input.readLine();!line.trim().equalsIgnoreCase("</PARAMS>");line=input.readLine()) 
              setParams(line.split("="));
      }
      if (logFile!=null)log=new PrintWriter(new FileWriter(logFile));
      //else log=new PrintWriter(System.out);
      System.out.println("Creating work folder...");
      File workDir = new File(workLibrary+"/XlsxWriter");
      if (workDir.exists()){
        deleteFolder(workDir);
        }
      workDir.mkdirs();
      System.out.println("Unzipping...");
      unzipExcelFile(xlsxTemplate);
      xl_workbook();
      docProps_core();
      //System.out.println("Double:"+getDoubleValue("41630D54FFF68872"));
      if (!sheetOverwriting){
          Content_Types();
          xl_rels_workbook();
          docProps_app();
      }
      xl_styles();
      processSheet();
      input.close();
      if (log!=null){
          log.flush();
          log.close();
      }
      }
      catch(Exception e){
          e.printStackTrace();
      }
  }  
  private void deleteFolder(File file){
      for (File subFile : file.listFiles()) {
         if(subFile.isDirectory()) {
            deleteFolder(subFile);
         } else {
            subFile.delete();
         }
      }
      file.delete();
   }
  private String getDoubleValue(String hexString){
      // hexString = "41630D54FFF68872";
    try{  
        return ""+Double.longBitsToDouble(Long.valueOf(hexString,16));
    }
    catch(Exception e){
        return "";
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
         ioe.printStackTrace();
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
  
  private void docProps_app(){//EZ JÓ!!! Node sorrend nem számít!
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
private void xl_workbook(){ // Ez jó!
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
EZ JÓ!!!
*/
private void xl_rels_workbook(){ //EZ JÓ!!! Node sorrend nem számít!
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

private void Content_Types(){ //EZ JÓ!!! Node sorrend nem számít!
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

private void docProps_core(){ //EZ JÓ!!! Node sorrend nem számít!
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
    else log("This excel file is created by XlsxWriter!");
    System.getProperty("user.name");
    Element lastModifiedBy=root.element("lastModifiedBy");
    if (lastModifiedBy.getText().contains("XlsxWriter")) {
        generatedByThis=true;
        log("This excel file is modified by XlsxWriter!");
    }
    else{
    root.remove(lastModifiedBy); 
    Element lmb=root.addElement("cp:lastModifiedBy");
    lmb.addText(System.getProperty("user.name")+" by XlsxWriter");
    }
    Date date = new Date();  
    SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ssz");  
    root.remove(root.element("modified"));
    Element modify=root.addElement("dcterms:modified")
    .addAttribute("xsi:type","dcterms:W3CDTF");
    modify.addText(formatter.format(date).replaceAll("CET","Z"));
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
// az fxstyle index nullától indul
private void xl_styles(){
    try{
       
    File ct=new File(workLibrary+"/XlsxWriter/xl/styles.xml");
    SAXReader reader = new SAXReader();
    Document document = reader.read(ct);
    Element root = document.getRootElement();
    log("RootElement: "+root.getName());
    Element borders=root.element("borders");
    borderCount=Integer.parseInt(borders.attributeValue("count"));
    if (!generatedByThis){
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
    }
    Element fills=root.element("fills");
    //if (fills==null)fills=root.addElement("fills").addAttribute("count","1");
    fillCount=Integer.parseInt(fills.attributeValue("count"));
    if (!generatedByThis){
        fillCount++;
        fills.remove(fills.attribute("count"));
        fills.addAttribute("count",""+fillCount);
        Element fill=fills.addElement("fill");
        Element patternFill=fill.addElement("patternFill").addAttribute("patternType","none");
        patternFill.addElement("bgColor").addAttribute("indexed","23");
        }
    Element fonts=root.element("fonts");
    if (fonts==null)fonts=root.addElement("fonts").addAttribute("count","2")
            .addAttribute("x14ac:knownFonts","1");
    fontCount=Integer.parseInt(fonts.attributeValue("count"));
    if (!generatedByThis){
        fontCount=fontCount+3;
        fonts.remove(fonts.attribute("count"));
        fonts.addAttribute("count",""+fontCount);
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
    }
    Element numFmts=root.element("numFmts");
    if (numFmts==null)numFmts=root.addElement("numFmts").addAttribute("count","2");
    numFmtCount=Integer.parseInt(numFmts.attributeValue("count"));
    if (!generatedByThis){
        numFmtCount=numFmtCount+2;
        numFmts.remove(numFmts.attribute("count"));
        numFmts.addAttribute("count",""+numFmtCount);
    }
    int id=163,fmtid;
    for (Iterator<Element> it = numFmts.elementIterator("numFmt"); it.hasNext();) {
            Element fmt=it.next();
            fmtid=Integer.parseInt(fmt.attributeValue("numFmtId"));
            if (id<fmtid)id=fmtid;
            }    
    if (!generatedByThis){
        id++;
        Element numFmt=numFmts.addElement("numFmt").addAttribute("numFmtId",""+id)
            .addAttribute("formatCode","[$-F400]h:mm:ss\\ AM/PM");
        id++;
        numFmt=numFmts.addElement("numFmt").addAttribute("numFmtId",""+id)
            .addAttribute("formatCode","yyyy/mm/dd\\ hh:mm:ss");
        }
    Element xfs=root.element("cellXfs");
    xfsCount=Integer.parseInt(xfs.attributeValue("count"));
    if (!generatedByThis){
        xfsCount=xfsCount+6;
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
