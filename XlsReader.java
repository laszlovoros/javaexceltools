package excelcasirtools;


import java.io.File;
import java.io.FileWriter;
import java.io.FileInputStream;
import java.io.PrintWriter;
import java.util.Date;
import java.util.HashSet;
import java.util.LinkedHashMap;

import jxl.Cell;
import jxl.DateCell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;

/**
 * Demo class which uses the api to read in a spreadsheet and generate a clone
 * of that spreadsheet which contains the same data.  If the spreadsheet read
 * in is the spreadsheet called jxlrwtest.xls (provided with the distribution)
 * then this class will modify certain fields in the copy of that spreadsheet.
 * This is illustrating that it is possible to read in a spreadsheet, modify
 * a few values, and write it under a new name.
 */
public class XlsReader
{

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
  private String xlsFile=null;
  private String logFile=null;
  private PrintWriter log=null;
  private String charSet="Cp1252";
  private String sheetName="table";
  private boolean logEnabled=true;
  private boolean breakAtEmptyRow=true;
  int textLength,x,y,z,i;
  static int[] dateFormatIds={14,15,16,17,27,28,29,31,31,36,50,51,54,57,58,81};
  static int[] timeFormatIds={18,19,20,21,22,45,46,47,32,33,34,35,55,56};
  static int[] dateTimeFormatIds={22,52,53};
  LinkedHashMap<String,String> numberFormats=new LinkedHashMap();
  StringBuilder t=new StringBuilder();
  StringBuilder n=new StringBuilder();
  Cell[] headerRow;
  Cell[] dataRow;
  String[] typeRow;
  int[] lengthRow;
  String[] formatRow;
  int cellStyleIndex=0;
  String sheetID="";
  char problemChars[]={'\u00f5','\u00fb','\u00db','\u00D5','\n'}; 
  char correctingChars[]={'õ','û','Û','Õ',' '};
  StringBuilder tmpbf=new StringBuilder();
  char betu;
  char[] dateproblemChars={'.','/','-',':'};
  boolean jobetu=true;	
  HashSet fieldNameChecker=new HashSet();
  Workbook workbook=null;
  Sheet sheet=null;
  DateCell datumcella;
  Date datum;
  long javaTime;
  final long exceljavasecdiff=2209161600l;
  double excelTime;
  Double D;
  
  public static void main(String[] args) {
      // TODO code application logic here
   XlsReader xr=new XlsReader();   
   for (int i=0;i<args.length;i++){
    if (args[i].toLowerCase().equals("-filename")) xr.xlsFile=args[i+1];
	if (args[i].toLowerCase().equals("-logfile")) xr.logFile=args[i+1];
    if (args[i].toLowerCase().equals("-startrow")) xr.startRow=Integer.parseInt(args[i+1]);
	if (args[i].toLowerCase().equals("-workingmode")) xr.workingMode=args[i+1];
	if (args[i].toLowerCase().equals("-workpath")) xr.workpath=args[i+1];
	if (args[i].toLowerCase().equals("-charset")) xr.charSet=args[i+1];
	if (args[i].toLowerCase().equals("-separator")) xr.separator=args[i+1];
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
    log("Identifying sheet ID...");
    WorkbookSettings ws = new WorkbookSettings(); 
    ws.setEncoding(charSet);
    FileInputStream fis = new FileInputStream(new File(xlsFile));
    Workbook workbook = Workbook.getWorkbook(fis,ws);
    log(workbook.toString());
    sheet=workbook.getSheet(sheetName);
    log(sheet.toString());
    if (startRow<1) startRow=getHeaderRowIndex();
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

  private int countTextColumns(Cell[] row) {
	  int count=0;
	  for (int i=0;i<row.length;i++) {
		  if (row[i]!=null) {
			  if (row[i].getType()==CellType.LABEL) {
				  if (!row[i].getContents().replaceAll("\\s+","").trim().equals(""))count++;
			  }
		  }
	  }
	  return count;
  }
  private int getHeaderRowIndex(){
    log("Exploring the sheet to find the header row...");    
    int maxLength=0,maxLengthRowIndex=0,hossz;
    for (int r=0;r<sheet.getRows();r++) {
    	hossz=countTextColumns(sheet.getRow(r));
        if (hossz>maxLength){
            maxLength=hossz;
            maxLengthRowIndex=r;
            }
        
        if (r>100) r=sheet.getRows();
        }
    log("Header row index is "+maxLengthRowIndex);
    return maxLengthRowIndex;
   }
  private void loadHeader() {
	  Cell[] row =sheet.getRow(startRow);
	  int maxcol=0;
	  for (int j=0;j<row.length;j++) if (maxcol<row[j].getColumn()) maxcol=row[j].getColumn();
	  maxcol++;
	  headerRow=new Cell[maxcol];
	  typeRow=new String[maxcol];
	  lengthRow=new int[maxcol];
	  formatRow=new String[maxcol];
	  for (int i=0;i<maxcol;i++) { 
		  headerRow[i]=null;
		  typeRow[i]=NUM;
		  lengthRow[i]=0;
	  }
	  String formatString;
	  int c,col;
	  Double numValue;
	  String textValue;
	  for (int j=0;j<row.length;j++) {
		  if (row[j].getType()==CellType.LABEL) {
			  if (!row[j].getContents().trim().equals(""))headerRow[row[j].getColumn()]=row[j];
			  
		  }
	  }
	  for (int r=startRow+1;r<sheet.getRows();r++) {
		  dataRow=sheet.getRow(r);
		  for (c=0;c<dataRow.length;c++) {
			  col=dataRow[c].getColumn();
			  if (col<headerRow.length) {
				  if (!typeRow[col].equals(TEXT)) {
					if (dataRow[c].getType()==CellType.LABEL) typeRow[col]=TEXT;
				    else if ((dataRow[c].getType()==CellType.DATE) && (typeRow[col].equals(NUM))) {
					  typeRow[col]=DATE;
					  formatString=dataRow[c].getCellFormat().getFormat().getFormatString();
                      formatRow[col]=formatString;
					  if (formatString.toLowerCase().indexOf("mm:")>-1) typeRow[col]=TIME;
					  else if (formatString.toLowerCase().indexOf("ss")>-1) typeRow[col]=TIME;
					  else if (formatString.toLowerCase().indexOf(":mm")>-1) typeRow[col]=TIME;
					  if (typeRow[col].equals(TIME)) {
						  if (formatString.toLowerCase().indexOf("yy")>-1) typeRow[col]=DATETIME;
						  else if (formatString.toLowerCase().indexOf("yyyy")>-1) typeRow[col]=DATETIME;
						  else if (formatString.toLowerCase().indexOf("dd")>-1) typeRow[col]=DATETIME;
						  else if (formatString.toLowerCase().indexOf("dd")>-1) typeRow[col]=DATETIME;
					  	}
					  else {
						  if (formatString.toLowerCase().indexOf("yy")>-1) typeRow[col]=DATE;
						  else if (formatString.toLowerCase().indexOf("yyyy")>-1) typeRow[col]=DATE;
						  else if (formatString.toLowerCase().indexOf("dd")>-1) typeRow[col]=DATE;
						  else if (formatString.toLowerCase().indexOf("dd")>-1) typeRow[col]=DATE;
					  }
				    }
				  }
				  if (dataRow[c].getType()==CellType.LABEL) {
					  textValue=dataRow[c].getContents();
					  if (lengthRow[col]<textValue.length()) lengthRow[col]=textValue.length();
				  }
			  }
		  }
	  }
  }
  private void printColumnData(){
      for (i=0;i<headerRow.length;i++){
          if (headerRow[i]!=null){
        	  System.out.println(getConsolidatedFieldName(headerRow[i].getContents()));
        	  System.out.println(typeRow[i]);
        	  System.out.println(lengthRow[i]);
        	  System.out.println(formatRow[i]);
              }
          }
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
              names.append(getConsolidatedFieldName(headerRow[i].getContents()));
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
      Cell[] row;
      int c,col;
      StringBuilder sb=new StringBuilder();
      boolean notFirst,ures;
      for (c=0;c<dataRow.length;c++) dataRow[c]=null;
      for (int i=startRow+1;i<sheet.getRows();i++){
              row=sheet.getRow(i);
              for (c=0;c<row.length;c++){
                  col=row[c].getColumn();
                  if (col<dataRow.length) dataRow[col]=row[c];
              }
              sb.setLength(0);
              ures=true;
              notFirst=false;
              for (c=0;c<dataRow.length;c++){
                  if (headerRow[c]!=null){
                      if (notFirst) sb.append(separator);
                      notFirst=true;
                      if (dataRow[c]!=null){
                          if (typeRow[c].equals(TEXT)||typeRow[c].equals(NUM))
                          sb.append(replaceProblemChars(dataRow[c].getContents()));
                          else {
                              sb.append(getDateValue(dataRow[c]));
                          }
                          if (dataRow[c].getType()!=CellType.EMPTY) ures=false;
                      }
                      else sb.append(' ');
                  }
              }
              System.out.println(sb.toString());
              //log(sb.toString());
              if (ures && breakAtEmptyRow)i=sheet.getRows();
          }
      
  }
  private String getDateValue(Cell c){
      String vissza;
      try{
        datumcella=(DateCell)c;
        datum=datumcella.getDate();
        excelTime=((datum.getTime()/1000)+exceljavasecdiff)/86400;
        }
      catch(Exception e){
          return "";
      }
      return new Double(excelTime).toString();
  }
}
