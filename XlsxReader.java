package excelcasirtools;
import java.util.zip.ZipFile;
import java.util.zip.ZipEntry;
import java.util.Enumeration;
import java.io.InputStream;
import java.io.IOException;
import java.io.File;
import java.io.FileWriter;
import java.io.PrintWriter;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.HashSet;
import java.util.Set;
import java.util.LinkedHashMap;
import java.io.RandomAccessFile;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.Characters;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

/**
 *
 * @author dell
 */

public class XlsxReader {
    static String DATE="D";
    static String TEXT="C";
    static String NUM="N";
    static String TIME="T";
    static String DATETIME="DT";
    static String PERCENT="P";
    static String ABC="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    private String workpath=null;
    private String separator="¤";
    private String workingMode="ALL";
    private int startRow=1;
    private String xlsxFile="c:/tmp/excelteszt.xlsx";
    private String logFile=null;
    private PrintWriter log=null;
    private String charSet=null;
    private String sheetName="table";
    private boolean logEnabled=true;
    private int cacheFillIndex=0;
    private String [] cache;
    private boolean breakAtEmptyRow=true;
    private long mapIndex=0;
    private int cacheSize=1000,actualRow=-1;
    private boolean mapFileNeed=false;
    private String cacheFileName=null;
    private RandomAccessFile textMapFile=null,textIndexFile=null;
    private String textMapFileName=null,textIndexFileName=null;
    long mapPos, indexPos;
    int textLength,x,y,z,i;
    static int[] dateFormatIds={14,15,16,17,27,28,29,31,31,36,50,51,54,57,58,81};
    static int[] timeFormatIds={18,19,20,21,22,45,46,47,32,33,34,35,55,56};
    static int[] dateTimeFormatIds={22,52,53};
    StringBuilder word;
    StringBuilder t=new StringBuilder();
    StringBuilder n=new StringBuilder();
    LinkedHashMap<String,String> cellAttributes;
    Cell[] headerRow;
    Cell[] dataRow;
    String[] typeRow;
    int[] lengthRow;
    int cellStyleIndex=0;
    LinkedHashMap<String,String> cellStyles=new LinkedHashMap();
    LinkedHashMap<String,String> numberFormats=new LinkedHashMap();
    String sheetID="";
    char problemChars[]={'\u00f5','\u00fb','\u00db','\u00D5','\n','\r'}; 
    char correctingChars[]={'õ','û','Û','Õ',' ',' '};
    StringBuilder tmpbf=new StringBuilder();
    char betu;
    char[] dateproblemChars={'.','/','-',':'};
    boolean jobetu=true;	
    HashSet fieldNameChecker=new HashSet();
    
    public static void main(String[] args) {
        // TODO code application logic here
     XlsxReader xr=new XlsxReader();   
     for (int i=0;i<args.length;i++){
        if (args[i].toLowerCase().equals("-filename")) xr.xlsxFile=args[i+1];
	if (args[i].toLowerCase().equals("-logfile")) xr.logFile=args[i+1];
	if (args[i].toLowerCase().equals("-workingmode")) xr.workingMode=args[i+1];
	if (args[i].toLowerCase().equals("-workpath")) xr.workpath=args[i+1];
	if (args[i].toLowerCase().equals("-charset")) xr.charSet=args[i+1];
	if (args[i].toLowerCase().equals("-separator")) xr.separator=args[i+1];
    if (args[i].toLowerCase().equals("-cachesize")) xr.cacheSize=Integer.parseInt(args[i+1]);
    if (args[i].toLowerCase().equals("-startrow")) xr.startRow=Integer.parseInt(args[i+1]);
	if (args[i].toLowerCase().equals("-sheetname")) xr.sheetName=args[i+1];
	if (args[i].toLowerCase().equals("-log")) xr.logEnabled=args[i+1].toLowerCase().equals("yes") ? true:false;
	if (args[i].toLowerCase().equals("-breakatemptyrow")) xr.breakAtEmptyRow=args[i+1].toLowerCase().equals("yes") ? true:false;
     }
     xr.process();
    }
  public void log(String msg){
      if (logEnabled){
      if (log!=null)log.println(msg);
      else System.out.println(msg);
      }
  }
  
  private void process(){
      try {
      for (i=0;i<dateFormatIds.length;i++) numberFormats.put(""+dateFormatIds[i], DATE);
      for (i=0;i<timeFormatIds.length;i++) numberFormats.put(""+timeFormatIds[i], TIME);
      for (i=0;i<dateTimeFormatIds.length;i++) numberFormats.put(""+dateTimeFormatIds[i], TIME);
      if (workpath==null) workpath=new File(".").getCanonicalPath();
      if (logFile!=null)log=new PrintWriter(new FileWriter(logFile));    
      else log=new PrintWriter(new FileWriter(workpath+"/xlsxreader.log"));
      log("CharSet: "+charSet);
      log("Identifying sheet ID...");
      sheetID=getSheetID();
      log("Sheet ID:"+sheetID);
      cache=new String[cacheSize];
      log("creating text value maps...");
      log("cache size is "+cacheSize);
      log("It has "+createSharedStringFile()+" elements");
      cellAttributes=new LinkedHashMap();
      log("loading styles into memory...");
      loadStyles();
      //printHashMap(numberFormats);
      //printHashMap(cellStyles);
      log("collecting information of table's header...");
      if (startRow<0) startRow=getHeaderRowIndex();
      loadHeader();
      if (workingMode.toUpperCase().equals("HEADER") || workingMode.toUpperCase().equals("ALL")){
    	  log("printing header...");
          printHeader();
      }
      if (workingMode.toUpperCase().indexOf("COL")>-1){
    	  log("printing fields info...");
          printColumnData();
      }
      if (workingMode.toUpperCase().equals("DATA") || workingMode.toUpperCase().equals("ALL")){
          log("printing data rows...");
          printData();
      }
      long memoryUsage=(Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory())/1048576;
      long totalMemory=Runtime.getRuntime().totalMemory()/1048576;
      log("----------------------------");
      log("Total memory: "+totalMemory+" MByte.");
      log("Memory usage: "+memoryUsage+" MByte.");
      log("----------------------------");
      if (textMapFile!=null) {
      log("Closing and deleting map files");
      closeMapFiles();
      }
      if (log!=null){
          log.flush();
          log.close();
      }
      }
      catch(Exception e){
          e.printStackTrace(log);
          if (log!=null){
              log.flush();
              log.close();
          }
      }
  }  
  private void printHashMap(LinkedHashMap<String,String> lhm){
      Set<String> s=lhm.keySet();
      Iterator<String> keys = s.iterator();
      String key;
      tmpbf.setLength(0);
      while (keys.hasNext()){
          key=keys.next();
          tmpbf.append('[').append(key).append(',').append(lhm.get(key)).append("] ");
      }
      log(tmpbf.toString());  
  }
  private int findInArray(int[] array,int value){
      for (int i=0;i<array.length;i++) if (array[i]==value) return i;
      return -1;
  }
  private void createMapFiles() throws IOException{
    String unique=workpath+"/c"+System.currentTimeMillis();
    textMapFileName=unique+".dat";
    textIndexFileName=unique+".idx";
	textMapFile = new RandomAccessFile(textMapFileName, "rw");
	textIndexFile = new RandomAccessFile(textIndexFileName, "rw");
    word=new StringBuilder(" ");
	}
  
   private void insertToMapFiles(String text)throws IOException{
    long pos=textMapFile.length();
    //textMapFile.seek(pos);
    textMapFile.writeChars(text);
    //textIndexFile.seek(textIndexFile.length());
    textIndexFile.writeLong(pos);
    //textMapFile.seek(textIndexFile.length());
    textIndexFile.writeInt(text.length());
  }
  private void closeMapFiles() throws IOException{
    if (mapFileNeed){  
    textMapFile.close();
    textIndexFile.close();
    File f=new File(textMapFileName);
    f.delete();
    f=new File(textIndexFileName);
    f.delete();
    }
  }
  
  private String getFromMap(int id){
    try{  
    mapPos=-1;
    indexPos=-1;
    textLength=-1;
    if (id<cache.length) return cache[id];
    id=id-cache.length;
    indexPos=12*id;
    textIndexFile.seek(indexPos);
    mapPos=textIndexFile.readLong();
    textLength=textIndexFile.readInt();
    textMapFile.seek(mapPos);
    word.delete(0, word.length());
    for (z=0;z<textLength;z++)word.append(textMapFile.readChar());
    }
    catch(Exception e){
        e.printStackTrace();
        log("id: "+id);
        log("cache size: "+cache.length);
        log("indexPos: "+indexPos);
        log("mapPos: "+mapPos);
        log("textLength: "+textLength);
    }
    return word.toString();
  }
  
  private long createSharedStringFile(){
  boolean tin=false;
  StartElement startElement;
  EndElement endElement;
  String qName;
  long counter=0;
  Characters characters;
  StringBuilder collector=new StringBuilder();
  try{      
    ZipFile zipFile = new ZipFile(xlsxFile);
    Enumeration<? extends ZipEntry> entries = zipFile.entries();
    while(entries.hasMoreElements()){
        ZipEntry entry = entries.nextElement();
        if (entry.getName().equals("xl/sharedStrings.xml")){
            InputStream stream = zipFile.getInputStream(entry);
            XMLInputFactory factory = XMLInputFactory.newInstance();
            XMLEventReader eventReader;
            if (charSet==null)
             eventReader=factory.createXMLEventReader(stream);
            else
             eventReader =factory.createXMLEventReader(stream,charSet);	
            while(eventReader.hasNext()) {
                XMLEvent event = eventReader.nextEvent();
                if (event.getEventType()==XMLStreamConstants.START_ELEMENT) {
                    startElement = event.asStartElement();
                    qName = startElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("t")) {
                        tin=true;
                        collector.setLength(0);
                    }
                    else{
                        if (tin)collector.append('<').append(qName).append('>');
                        }
                    }
                if (event.getEventType()==XMLStreamConstants.END_ELEMENT) {
                    endElement = event.asEndElement();
                    qName = endElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("t")) {
                        counter++;
                        tin=false;
                        if (cacheFillIndex<cache.length){
                        cache[cacheFillIndex]=collector.toString();
                        cacheFillIndex++;
                        }
                        else{
                        if (!mapFileNeed){ 
                            createMapFiles();
                            mapFileNeed=true;
                            log("Cache is full, further values will be written to file.");
                            }
                        insertToMapFiles(collector.toString());
                        }
                    }
                    else{
                        if (tin)collector.append('<').append(qName).append("/>");
                    }
                    }
                if (event.getEventType()==XMLStreamConstants.CHARACTERS){
                  if (tin){  
                    characters = event.asCharacters();
                    collector.append(characters.getData());
                  }
                } 
            }
            stream.close();
        }
    }
    zipFile.close();
    }
    catch(Exception e){
          e.printStackTrace(log);
      }
  return counter;
  }
  
String replaceProblemChars(String miben){
    if (problemChars.length<1) return miben;
   	tmpbf.setLength(0);
   	int meddig=miben.length();
        for (int i=0;i<meddig;i++){
		betu=miben.charAt(i);
        belsoloop:
		for (int j=0;j<problemChars.length;j++){
			if (betu==problemChars[j]){
                            betu=correctingChars[j];
                            break belsoloop;
                            }
			}
		tmpbf.append(betu);		
		}
	return tmpbf.toString();
   }
  	
  	
 String numberDate(String miben){
  if (problemChars.length<1) return miben;
  tmpbf.setLength(0);
  int meddig=miben.length();
  for (int i=0;i<meddig;i++){
        jobetu=true;
        betu=miben.charAt(i);
        innerloop:
        for (int j=0;j<dateproblemChars.length;j++){
            if (betu==dateproblemChars[j]){
		jobetu=false;
		break innerloop;
		}
            }
        if (jobetu)tmpbf.append(betu);		
        }
    return tmpbf.toString();
   }

  
  public String getConsolidatedFieldName(String text){
  	text=replaceProblemChars(text.toLowerCase());
   	text=text.replace('ö','o').replace('í','i').replace('á','a').replace('é','e').replace('û','u');
   	text=text.replace('õ','o').replace(' ','_').replace('ü','u').replace('ú','u').replace('ó','o');
   	text=text.replace('á','a').replace(':','_').replace('(','_').replace(')','_').replace('?','_');
   	StringBuffer seq=new StringBuffer("");
   	char[] stock="abcdefghijklmnopqrstuw_vzxy1234567089".toCharArray();
   	char[] karakterek=text.toCharArray();
   	for (int i=0;i<karakterek.length;i++){
   		inner:
   		for (int j=0;j<stock.length;j++){
   			if (karakterek[i]==stock[j])seq.append(karakterek[i]);
   			if (karakterek[i]==stock[j]) break inner;
   		}
   	}
   	String txt=seq.toString();
   	if (fieldNameChecker.contains(txt)){
   		txt="_"+txt;
   		fieldNameChecker.add(txt);
	   	}
	else fieldNameChecker.add(txt);   	
   	return txt;
   }

  
  private LinkedHashMap<String,String> getElementAttributes(StartElement startElement){
  LinkedHashMap<String,String> attrs=new LinkedHashMap();    
  Iterator<Attribute> attributes = startElement.getAttributes();
  while (attributes.hasNext()){
        Attribute a = attributes.next();
        attrs.put(a.getName().getLocalPart().toLowerCase(),a.getValue());
        }
  return attrs;
  }
  private String getElementAttribute(StartElement startElement,String attributeName){
  Iterator<Attribute> attributes = startElement.getAttributes();
    while (attributes.hasNext()){
        Attribute a = attributes.next();
        if (a.getName().getLocalPart().equalsIgnoreCase(attributeName))
        return a.getValue();
        }
    return null;
  }

  private void loadCellAttributes(StartElement startElement){
  Iterator<Attribute> attributes = startElement.getAttributes();
  cellAttributes.clear();
  while (attributes.hasNext()){
        Attribute a = attributes.next();
        cellAttributes.put(a.getName().getLocalPart(),a.getValue());
        }
  }
  private LinkedList<Cell> getRow(XMLEventReader eventReader,StartElement startElement) 
          throws XMLStreamException,IOException{
    actualRow=Integer.parseInt(getElementAttribute(startElement,"r"));  
    LinkedList<Cell> sor=new LinkedList();
    EndElement endElement;
    String qName;
    boolean rowEnd=false;
    boolean rin=true;
      while(!rowEnd){
                XMLEvent event = eventReader.nextEvent();
                if (event.getEventType()==XMLStreamConstants.START_ELEMENT) {
                    startElement = event.asStartElement();
                    qName = startElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("c")) sor.add(getCell(eventReader,startElement));
                    }
                if (event.getEventType()==XMLStreamConstants.END_ELEMENT) {
                    endElement = event.asEndElement();
                    qName = endElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("row")) rowEnd=true;
                    }
                }
    return sor;
  }
  private boolean loadDataRow(XMLEventReader eventReader,StartElement startElement) 
          throws XMLStreamException,IOException{
    actualRow=Integer.parseInt(getElementAttribute(startElement,"r"));  
    for (i=0;i<dataRow.length;i++) dataRow[i]=null;  
    EndElement endElement;
    String qName;
    Cell c;
    boolean rowEnd=false;
    boolean rin=true;
    boolean vissza=false;
      while(!rowEnd){
                XMLEvent event = eventReader.nextEvent();
                if (event.getEventType()==XMLStreamConstants.START_ELEMENT) {
                    startElement = event.asStartElement();
                    qName = startElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("c")){
                        c=getCell(eventReader,startElement);
                        dataRow[c.col]=c;
                        if ((c.numValue!=null) || (c.textValue!=null)) vissza=true;
                    }
                    }
                if (event.getEventType()==XMLStreamConstants.END_ELEMENT) {
                    endElement = event.asEndElement();
                    qName = endElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("row")) rowEnd=true;
                    }
                }
      return vissza;
  }
  
  private Cell getCell(XMLEventReader eventReader,StartElement startElement) 
          throws XMLStreamException,IOException{
  EndElement endElement;
  String qName;
  int mapPosition;
  String txtValue="";
  Double numValue=null;
  String cellType=NUM;
  Characters characters;
  String styleId;
  boolean cellEnd=false,vin=false,tin=false,inlineStr=false;
  loadCellAttributes(startElement);
  int[] coords=getPositions(cellAttributes.get("r"));
  if (cellAttributes.containsKey("t")){
      if (cellAttributes.get("t").equalsIgnoreCase("s")) cellType=TEXT;
      if (cellAttributes.get("t").equalsIgnoreCase("inlineStr")) {
    	  cellType=TEXT;
    	  inlineStr=true;
      }
  }
  else if (cellAttributes.containsKey("s")){
      styleId=cellAttributes.get("s");
      cellType=cellStyles.get(styleId);
  }
  Cell cell=new Cell(coords[0],coords[1],cellType);
  while(!cellEnd){
                txtValue="";
                numValue=null;
                XMLEvent event = eventReader.nextEvent();
                if (event.getEventType()==XMLStreamConstants.START_ELEMENT) {
                    startElement = event.asStartElement();
                    qName = startElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("v")) vin=true;
                    else vin=false;
                    if (qName.equalsIgnoreCase("t")) tin=true;
                    else tin=false;
                    }
                if (event.getEventType()==XMLStreamConstants.END_ELEMENT) {
                    endElement = event.asEndElement();
                    qName = endElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("c")) {
                        cellEnd=true;
                        //cell.setValue(txtValue);
                        //cell.setValue(numValue);
                    }
                    if (qName.equalsIgnoreCase("v")) vin=false;
                    if (qName.equalsIgnoreCase("t")) tin=false;
                    }
                if (event.getEventType()==XMLStreamConstants.CHARACTERS){
                  if (vin || tin){  
                    characters = event.asCharacters();
                    if (!cellType.equals(TEXT)) {
                        numValue=new Double(Double.parseDouble(characters.getData()));
                        txtValue=characters.getData();
                    }
                    else{
                    	if (tin) {
                    		txtValue=characters.getData();
                    		numValue=null;
                    	}
                    	else {
                        mapPosition=Integer.parseInt(characters.getData());
                        numValue=new Double(mapPosition);
                        txtValue=getFromMap(mapPosition);
                    	}
                    }
                    vin=false;
                    tin=false;
                    cell.setValue(txtValue);
                    cell.setValue(numValue);
                    }
                }  
            }
    return cell;
  }
  
  private void createHeader(LinkedList<Cell> row){
    int colCount=-1;
    Cell c;
    for (i=0;i<row.size();i++){
        c=row.get(i);
        if (colCount<c.col)colCount=c.col;
    }
    log("the header has "+colCount+" column.");
    colCount++;
    headerRow=new Cell[colCount];
    dataRow=new Cell[colCount];
    typeRow=new String[colCount];
    lengthRow=new int[colCount];
    for (i=0;i<headerRow.length;i++){
        headerRow[i]=null;
        dataRow[i]=null;
        typeRow[i]=NUM;
        lengthRow[i]=1;
    }

    for (i=0;i<row.size();i++){
        c=row.get(i);
        headerRow[c.col]=c;
    }
    row.clear();
    row=null;
  }

  private int getHeaderRowIndex(){
  log("Exploring the sheet to find the header row...");    
  int maxLength=0,maxLengthRowIndex=0;
  try{      
    int rowIndex=0;
    StartElement startElement;
    EndElement endElement;
    String qName,name;
    LinkedList <Cell> sor;
    boolean sheetEnd=false;
    ZipFile zipFile = new ZipFile(xlsxFile);
    Enumeration<? extends ZipEntry> entries = zipFile.entries();
    while(entries.hasMoreElements()){
        ZipEntry entry = entries.nextElement();
        if (entry.getName().equals("xl/worksheets/sheet"+sheetID+".xml")){
            InputStream stream = zipFile.getInputStream(entry);
            XMLInputFactory factory = XMLInputFactory.newInstance();
            XMLEventReader eventReader;
            if (charSet==null)
                eventReader=factory.createXMLEventReader(stream);
               else
                eventReader =factory.createXMLEventReader(stream,charSet);	
                while(!sheetEnd) {
                XMLEvent event = eventReader.nextEvent();
                if (event.getEventType()==XMLStreamConstants.START_ELEMENT) {                
                  startElement = event.asStartElement();
                  qName = startElement.getName().getLocalPart();
                  if (qName.equalsIgnoreCase("row")) {
                     rowIndex=Integer.parseInt(getElementAttribute(startElement,"r"));
                     if (rowIndex>100) sheetEnd=true;
                     else {
                         sor=getRow(eventReader,startElement);  
                         if (maxLength<sor.size()) {
                         maxLength=sor.size();
                         maxLengthRowIndex=rowIndex;
                         }
                       }
                    }
                } 
                if (event.getEventType()==XMLStreamConstants.END_ELEMENT) {
                    endElement = event.asEndElement();
                    qName = endElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("sheetData")) sheetEnd=true;
                    }
            }
            stream.close();
        }
    }
    zipFile.close();
    }
    catch(Exception e){
          e.printStackTrace(log);
      }
  log("Header row index is "+maxLengthRowIndex);
  return maxLengthRowIndex;  
  }
    
  
  private void loadHeader(){
  try{      
    int rowIndex=0;
    StartElement startElement;
    EndElement endElement;
    String qName,name;
    boolean sheetEnd=false;
    ZipFile zipFile = new ZipFile(xlsxFile);
    Enumeration<? extends ZipEntry> entries = zipFile.entries();
    while(entries.hasMoreElements()){
        ZipEntry entry = entries.nextElement();
        if (entry.getName().equals("xl/worksheets/sheet"+sheetID+".xml")){
            InputStream stream = zipFile.getInputStream(entry);
            XMLInputFactory factory = XMLInputFactory.newInstance();
            XMLEventReader eventReader;
            if (charSet==null)
                eventReader=factory.createXMLEventReader(stream);
               else
                eventReader =factory.createXMLEventReader(stream,charSet);	
 sheet:     while(!sheetEnd) {
                XMLEvent event = eventReader.nextEvent();
                if (event.getEventType()==XMLStreamConstants.START_ELEMENT) {                
                  startElement = event.asStartElement();
                  qName = startElement.getName().getLocalPart();
                  if (qName.equalsIgnoreCase("row")) {
                     rowIndex=Integer.parseInt(getElementAttribute(startElement,"r"));
                     if (rowIndex==startRow)
                       createHeader(getRow(eventReader,startElement));  
                     if (rowIndex>startRow){
                        if (!loadDataRow(eventReader,startElement)) break sheet;
                        for (i=0;i<headerRow.length;i++){
                           if ((headerRow[i]!=null) && (dataRow[i]!=null)){
                        	   if (dataRow[i].getTextValue().length()>lengthRow[i])
                                   lengthRow[i]=dataRow[i].getTextValue().length();
                               if (dataRow[i].cellType.equals(TEXT))
                                   typeRow[i]=TEXT;
                               if (typeRow[i].equals(NUM)) {
                                if (dataRow[i].cellType.equals(DATE))
                                   typeRow[i]=DATE;
                                if (dataRow[i].cellType.equals(DATETIME))
                                   typeRow[i]=DATETIME;
                                if (dataRow[i].cellType.equals(TIME))
                                   typeRow[i]=TIME;
                               }
                           } 
                        }
                     }
                    }
                } 
                if (event.getEventType()==XMLStreamConstants.END_ELEMENT) {
                    endElement = event.asEndElement();
                    qName = endElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("sheetData")) sheetEnd=true;
                    }
                //if ((rowIndex-startRow)>200) sheetEnd=true;
            }
            stream.close();
        }
    }
    zipFile.close();
    }
    catch(Exception e){
          e.printStackTrace(log);
      }
      
  }
  private void printColumnData(){
      for (i=0;i<headerRow.length;i++){
          if (headerRow[i]!=null){
        	  System.out.println(getConsolidatedFieldName(headerRow[i].textValue));
        	  System.out.println(typeRow[i]);
        	  System.out.println(lengthRow[i]);
              }
          }
      }
  
  private void printHeader(){
      StringBuilder names=new StringBuilder();
      StringBuilder types=new StringBuilder();
      StringBuilder lengths=new StringBuilder();
      boolean elso=true;
      for (i=0;i<headerRow.length;i++){
          if (headerRow[i]!=null){
              if (!elso) {
                  names.append(separator);
                  types.append(separator);
                  lengths.append(separator);
              }
              names.append(getConsolidatedFieldName(headerRow[i].textValue));
              types.append(typeRow[i]);
              lengths.append(lengthRow[i]);
              elso=false;
          }
      }
      System.out.println(names.toString());
      System.out.println(types.toString());
      System.out.println(lengths.toString());
  }
  
  private void printData(){
    boolean elso;
    StartElement startElement;
    EndElement endElement;
    String qName,name;
    int rowIndex;
    boolean sheetEnd=false;
    try{
    ZipFile zipFile = new ZipFile(xlsxFile);
    Enumeration<? extends ZipEntry> entries = zipFile.entries();
    while(entries.hasMoreElements()){
        ZipEntry entry = entries.nextElement();
        if (entry.getName().equals("xl/worksheets/sheet"+sheetID+".xml")){
            InputStream stream = zipFile.getInputStream(entry);
            XMLInputFactory factory = XMLInputFactory.newInstance();
            XMLEventReader eventReader;
            if (charSet==null)
                eventReader=factory.createXMLEventReader(stream);
               else
                eventReader =factory.createXMLEventReader(stream,charSet);	
torzs:      while(!sheetEnd) {
                XMLEvent event = eventReader.nextEvent();
                if (event.getEventType()==XMLStreamConstants.START_ELEMENT) {                
                  startElement = event.asStartElement();
                  qName = startElement.getName().getLocalPart();
                  if (qName.equalsIgnoreCase("row")) {
                     rowIndex=Integer.parseInt(getElementAttribute(startElement,"r"));
                     if (rowIndex>startRow){
                        if (!loadDataRow(eventReader,startElement) && breakAtEmptyRow) break torzs;
                        tmpbf.setLength(0);
                        elso=true;
                        for (i=0;i<headerRow.length;i++){
                           if (headerRow[i]!=null){
                               if (!elso) tmpbf.append(separator);
                               else elso=false;
                               if (dataRow[i]!=null) {
                            	   if (typeRow[i].equals(TEXT)) tmpbf.append(dataRow[i].getTextValue());
                            	   else tmpbf.append(dataRow[i].getNumValue());
                               }
                               else tmpbf.append("");
                           } 
                        }
                        System.out.println(replaceProblemChars(tmpbf.toString()));
                     }
                    }
                } 
                if (event.getEventType()==XMLStreamConstants.END_ELEMENT) {
                    endElement = event.asEndElement();
                    qName = endElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("sheetData")) sheetEnd=true;
                    }
            }
            stream.close();
        }
    }
    zipFile.close();
    }
    catch(Exception e){
          e.printStackTrace(log);
      }
      
      
  }
  
  private String getSheetID(){
  boolean found ;
  String ID=null, tempID;
  StartElement startElement;
  String qName,name;
  try{      
    ZipFile zipFile = new ZipFile(xlsxFile);
    Enumeration<? extends ZipEntry> entries = zipFile.entries();
    while(entries.hasMoreElements()){
        ZipEntry entry = entries.nextElement();
        if (entry.getName().equals("xl/workbook.xml")){
            InputStream stream = zipFile.getInputStream(entry);
            XMLInputFactory factory = XMLInputFactory.newInstance();
            XMLEventReader eventReader;
            if (charSet==null)
                eventReader=factory.createXMLEventReader(stream);
               else
                eventReader =factory.createXMLEventReader(stream,charSet);	
            while(eventReader.hasNext()) {
                XMLEvent event = eventReader.nextEvent();
                if (event.getEventType()==XMLStreamConstants.START_ELEMENT) {                
                  found=false;
                  tempID=null;
                  startElement = event.asStartElement();
                  qName = startElement.getName().getLocalPart();
                  if (qName.equalsIgnoreCase("sheet")) {
                     log("Start Element : sheet");
                     name=getElementAttribute(startElement,"name");
                     if (name!=null){
                        if (name.equalsIgnoreCase(sheetName)) 
                            ID=getElementAttribute(startElement,"sheetId");
                        }
                    }
                } 
                if (ID!=null) break;
            }
            stream.close();
        }
    }
    zipFile.close();
    }
    catch(Exception e){
          e.printStackTrace(log);
      }
   return ID;
  } 
  private void loadNumberFormats(XMLEventReader eventReader,StartElement startElement)
          throws XMLStreamException{
    boolean fEnd=false;
    String qName;
    String id;
    String code;
    int index=0;
    String type="N";
    EndElement endElement;
    while(!fEnd){
                XMLEvent event = eventReader.nextEvent();
                if (event.getEventType()==XMLStreamConstants.START_ELEMENT) {
                    startElement = event.asStartElement();
                    qName = startElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("numFmt")) {
                        id=getElementAttribute(startElement,"numFmtId");
                        code=getElementAttribute(startElement,"formatCode");
                        if (code.indexOf('d')>-1){
                            type=DATE;
                            if (code.indexOf("h:")>-1)type=DATETIME;
                            }
                        else if (code.indexOf("h:")>-1)type=TIME;
                        else if (code.indexOf("yyyy")>-1) type=DATE;
                        if (!type.equals("N")) {
                            numberFormats.put(id,type);
                            index++;
                        }
                        }
                    
                    }
                if (event.getEventType()==XMLStreamConstants.END_ELEMENT) {
                    endElement = event.asEndElement();
                    qName = endElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("numFmts")) {
                        fEnd=true;
                    }
                }
    }
    log(""+index+" special formats loaded into memory.");
  }
  private void loadCellStyles(XMLEventReader eventReader,StartElement startElement)
  throws XMLStreamException {
    boolean xEnd=false;
    String qName;
    String id;
    String code;
    String type="N";
    int index=0;
    EndElement endElement;
    while(!xEnd){
                XMLEvent event = eventReader.nextEvent();
                if (event.getEventType()==XMLStreamConstants.START_ELEMENT) {
                    startElement = event.asStartElement();
                    qName = startElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("xf")) {
                        id=getElementAttribute(startElement,"numFmtId");
                        if (numberFormats.containsKey(id)) 
                            cellStyles.put(""+index,numberFormats.get(id));
                        else cellStyles.put(""+index,NUM);
                        index++;
                        }
                      }
                if (event.getEventType()==XMLStreamConstants.END_ELEMENT) {
                    endElement = event.asEndElement();
                    qName = endElement.getName().getLocalPart();
                    if (qName.equalsIgnoreCase("cellXfs")) {
                        xEnd=true;
                    }
                }
    }
    log(""+index+" cell style loaded into memory...");
  }
  
  private void loadStyles(){
  boolean found ;
  String ID=null, tempID;
  try{      
    ZipFile zipFile = new ZipFile(xlsxFile);
    Enumeration<? extends ZipEntry> entries = zipFile.entries();
    while(entries.hasMoreElements()){
        ZipEntry entry = entries.nextElement();
        if (entry.getName().equals("xl/styles.xml")){
            InputStream stream = zipFile.getInputStream(entry);
            XMLInputFactory factory = XMLInputFactory.newInstance();
            XMLEventReader eventReader;
            if (charSet==null)
                eventReader=factory.createXMLEventReader(stream);
               else
                eventReader =factory.createXMLEventReader(stream,charSet);	
            while(eventReader.hasNext()) {
                XMLEvent event = eventReader.nextEvent();
                if (event.getEventType()==XMLStreamConstants.START_ELEMENT) {                
                  StartElement startElement = event.asStartElement();
                  String qName = startElement.getName().getLocalPart();
                  if (qName.equalsIgnoreCase("numFmts")) loadNumberFormats(eventReader,startElement);
                  if (qName.equalsIgnoreCase("cellXfs")) loadCellStyles(eventReader,startElement);
                } 
            }
            stream.close();
        }
    }
    zipFile.close();
    }
    catch(Exception e){
          e.printStackTrace(log);
      }
  } 

  
  int[] getPositions(String excelCoords){
          int[] vissza=new int[2];
          if (t.length()>0) t.delete(0, t.length());
          if (n.length()>0) n.delete(0, n.length());
          for(z=0;z<excelCoords.length();z++){
              if (ABC.indexOf(excelCoords.charAt(z))>-1)t.append(excelCoords.charAt(z));
              else n.append(excelCoords.charAt(z));
          }
          //log(t.toString());
          //log(n.toString());
          vissza[1]=Integer.parseInt(n.toString())-1;
          vissza[0]=-1;
          for (z=0;z<t.length();z++){
              vissza[0]=vissza[0]+(ABC.indexOf(t.charAt(z))+1)*power(ABC.length(),t.length()-(z+1));
          }
          return vissza;     
      }
      
      int power(int mit,int mire){
          int value=1;
          for (x=0;x<mire;x++)value=value*mit;
          //log("mit: "+mit+" mire: "+mire+" value:"+value);
          return value;
      }
      
  class Cell{
      String textValue=null;
      Double numValue=null;
      String cellType=NUM;
      int col=0,row=0;
      String excelPos="--";
      int sasDateValue=0;
      float sasTimeValue=0;
      
      public Cell(){
          
      }
      public Cell(int col,int row,String cellType,String textValue,Double numValue){
          this.col=col;
          this.row=row;
          this.cellType=cellType;
          this.textValue=textValue;
          this.numValue=numValue;
          }
      public Cell(int col,int row,String cellType){
          this.col=col;
          this.row=row;
          this.cellType=cellType;
          }
      public void setValues(int col,int row,String cellType,String textValue,Double numValue){
          this.col=col;
          this.row=row;
          this.cellType=cellType;
          this.textValue=textValue;
          this.numValue=numValue;
          }
      public String getNumValue() {
    	  if (numValue==null) return "";
    	  else return numValue.toString();
      }
      public String getTextValue() {
    	  if (textValue==null) return "";
    	  else return textValue;
      }
      public void setType(String cellType){
          this.cellType=cellType;
      }
      public void setValue(String textValue){
          this.textValue=textValue;
      }
      public void setValue(Double numValue){
          this.numValue=numValue;
      }
  }
  }


