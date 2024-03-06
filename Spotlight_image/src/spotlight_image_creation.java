import java.awt.Graphics;
import java.awt.GraphicsEnvironment;
import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import javax.imageio.ImageIO;
import javax.swing.JEditorPane;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

import ru.yandex.qatools.ashot.AShot;
import ru.yandex.qatools.ashot.Screenshot;
import ru.yandex.qatools.ashot.shooting.ShootingStrategies;

public class spotlight_image_creation {
	
	static String Nominee_emp;
	static String Nominee_name;
	static String reason;
	static String gp;
	static String nm_by;
	static String filecount;
	static String SL_date;
	static String imagename;
	static String inputfilelocation;
	static String finanlstr;
    static String htmlend = "</div></div><div class=\"footer\">www.nallas.com</div></div></body></html>";
    static String htmlstring;
    static StringBuilder htmlstr = new StringBuilder();
    static String currentdate;
    static String finaloutput_html;
    static String defaultvalue="test";
    static WebDriver driver;
    static String group;
    static List<String> groupvalue = new ArrayList<String>();
    static String[] uniquevalue;
	public static void main(String[] args) throws IOException, FilloException, InterruptedException 
	
	{
		
		String inputfilelocation = System.getProperty("user.dir")+"\\input\\excel\\input.xlsx";
		Fillo fillo = new Fillo();
		Connection connection = fillo.getConnection(inputfilelocation);
		Recordset recordset1 = connection.executeQuery("SELECT * FROM sheet1");
		int numberOfRows = recordset1.getCount();
	
		System.out.println("Execution Started");
		
		try
		{
		 File excel = new File(inputfilelocation);
	        FileInputStream fis = new FileInputStream(excel);

	        XSSFWorkbook wb = new XSSFWorkbook(fis);
	        XSSFSheet ws = wb.getSheetAt(0);
	        System.out.println("Size:");
	        for(int a=1;a<numberOfRows+1;a++)
	        {
	        XSSFRow rowHeader = ws.getRow(a);
	        String countfile;
	        XSSFCell cell = rowHeader.getCell(5);
	        try
	        {
	         countfile = cell.getRawValue();
	        }
	        catch(Exception e)
	        {
	        	countfile = "";
	        }
	        if(countfile==null||countfile=="")
	        {
	        	defaultvalue ="empty";
	        	
	        }
	        System.out.println("Size:"+countfile);
	        groupvalue.add(countfile);
	        }
		}
		catch(Exception e)
		{
			System.out.println(e.toString());
		}
	        
		uniquevalue=  findunique();
		
		
		for(String val:uniquevalue)
		{
			Recordset recordset;
			System.out.println("setup");
			if(defaultvalue.equalsIgnoreCase("empty"))
			{
		     recordset = connection.executeQuery("SELECT * FROM sheet1");
			}
			else
			{
				recordset = connection.executeQuery("SELECT * FROM sheet1 where file_group ='"+val+"'");
			}
		
		int count = recordset.getCount();
		System.out.println("Count of unique list"+count);
		
		int i = 0;
		
		while (recordset.next()) 
		{
			Nominee_emp = recordset.getField("Nominee");
			reason = recordset.getField("Reason");
			gp = recordset.getField("GP");
			nm_by = recordset.getField("Nominated_By");
			SL_date = recordset.getField("date");
			group = recordset.getField("file_group");
			find_image(Nominee_emp);
			contentbox();
		}
		
		spotlightprocess();
		 finanlstr = "";
		}
		driver.quit();
		
		System.out.println("SpotLight Automation execution Completed");
	}
	
	@SuppressWarnings("unchecked")
	public static String[] findunique() 
	{
		
		@SuppressWarnings("rawtypes")
		Set distinct = new HashSet();
	    for(String element : groupvalue) 
	    {
	        distinct.add(element);
	    }

	    return (String[]) distinct.toArray(new String[0]);
		
	}
	
	public static void spotlightprocess() throws IOException, InterruptedException {
		currentdate();
		String filename = System.getProperty("user.dir")+"\\input\\html\\htmldata1.html";
		
		File file = new File(filename);
        FileReader fr = new FileReader(file);
        BufferedReader br = new BufferedReader(fr);
        String line;
        while ((line = br.readLine()) != null) {
            // process the line
        	htmlstr.append(line);
        }
        String dateval =  date_input();
        htmlstr=htmlstr.append(dateval+finanlstr+htmlend);
        br.close();
        fr.close();
        
        createhtml(htmlstr.toString());
        
        openhtmlfile();
        htmltoimage();
        htmlstr.setLength(0);
	}
	
	public static void find_image(String empid)
	{
		
		String[] emp_split = empid.split("-");
		String em_id = emp_split[0];
		Nominee_name =  emp_split[1];
		
		inputfilelocation = System.getProperty("user.dir")+"\\input\\image\\";
		
		 File directory = new File(inputfilelocation); 
		  
	        // store all names with same name 
	        // with/without extension 
	        String[] flist = directory.list(); 
	        int flag = 0; 
	        if (flist == null) { 
	            System.out.println("Empty directory."); 
	        } 
	        else { 
	  
	            // Linear search in the array 
	            for (int i = 0; i < flist.length; i++) { 
	                String filename = flist[i]; 
	                //System.out.println("Empty directory."+filename); 
	                if (filename.contains(em_id)) { 
	                	imagename = filename; 
	                    //System.out.println(filename + " found"); 
	                    flag = 1; 
	                } 
	            } 
	        } 
		
	}
	
	
	public static void contentbox()
	{
		String content1 = 
				"<div class=\"content-box\">\r\n"
				+ " <div class=\"content-box-left\">\r\n"
				+ "<div class=\"content-thumb\">";
		String image =  "<img src=\""+inputfilelocation +imagename +"\"/>"+"</div>";
		String empname = "<div class=\"content-name\">"+Nominee_name+"</div></div>";
		String reason_html = "<div class=\"content-box-right\"><div class=\"content-box-info\"><div class=\"content-box-label\">Reason</div><div class=\"content-box-text\">";
		String reason_content = reason+"</div></div>";
		String gp_html = "<div class=\"content-box-info\"><div class=\"content-box-label\">Guiding Principles</div><div class=\"content-box-text\"><b>";
		String gp_content = gp+ "</b></div></div>";
		String nomi_html = "<div class=\"content-box-info\"><div class=\"content-box-label\">Nominated By</div><div class=\"content-box-text\">";
		String nomi = nm_by + "</div></div></div></div>";
		finanlstr += content1+image+empname+reason_html+reason_content+gp_html+gp_content+nomi_html+nomi;
	    
	}
	
	public static String  date_input()
	{		
		String dateselect = "<div class=\"date\">"+SL_date+"</div><div class=\"content\">\r\n";
	   return dateselect;
	}
	
	
	public static void createhtml(String data) throws IOException
	{
		String location = System.getProperty("user.dir")+"\\output\\spotlight"+currentdate+"--"+group+".html";
		finaloutput_html = location;
		FileWriter fw = new FileWriter(location);
		BufferedWriter writer = new BufferedWriter(fw);
	    writer.write(data);
	    writer.newLine(); //this is not actually needed for html files - can make your code more readable though 
	    writer.close();
	    fw.close();
	    System.out.println("Html file created. Folder location: "+location);
	}
	
	public static void openhtmlfile() throws InterruptedException
	{
		String location = System.getProperty("user.dir")+"\\Drivers\\chromedriver.exe";
		
		System.setProperty("webdriver.chrome.driver", location);
    	ChromeOptions option = new ChromeOptions();
    	option.addArguments("--headless", "--window-size=1920,1200","--start-fullscreen");
    	System.out.println(finaloutput_html);
    	driver = new ChromeDriver(option);
    	driver.get(finaloutput_html);
    	driver.manage().window().fullscreen();
        Thread.sleep(3000);
    	
	}
	
	public static void currentdate() 
	{
		String date = new SimpleDateFormat("dd-MM-yyyy").format(new Date());
		currentdate = date;
		System.out.println("current date"+currentdate);
	}
	
	public static void htmltoimage() throws IOException, InterruptedException
	{
		 String fileWithPath = System.getProperty("user.dir")+"\\output\\spotlight"+currentdate+"--"+group+".jpeg";
		//Screenshot screenshot = new AShot().takeScreenshot(driver);
		Screenshot s=new AShot().shootingStrategy(ShootingStrategies.viewportPasting(2000)).takeScreenshot(driver);
		ImageIO.write(s.getImage(),"png",new File(fileWithPath));
		System.out.println("Image file created. Folder location: "+fileWithPath);
	}
	

}
