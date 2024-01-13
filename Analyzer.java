import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.temporal.ChronoUnit;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet; 

public class Analyzer
{
    int j=1,rowCount,flg=0,i=0,m=0,n=0;
    XSSFSheet sheet;
    String path="./AssignmentTimecard.xlsx",id; //Path Excel Sheet
    XSSFWorkbook book;

    //Declaration of arrays
    String consDay[] = new String[50];      
    String ShiftDiff[] = new String[50];
    String ShiftWork[] = new String[50];

    Analyzer()
    {
        int k=0,temp;        
        try{
            book = new XSSFWorkbook(path);
            sheet = book.getSheet("Sheet1"); 

            //getting count of No. of rows in excel sheet
            rowCount = sheet.getPhysicalNumberOfRows(); 
            
            // while loop to scan whole excel sheet
            while(j<rowCount-1)                         
            { 
                //taking position ID to check similarity
                id = sheet.getRow(j).getCell(0).getStringCellValue();                
                temp=j;

                //while loop to count how many similar ID are there
                while(id.equals(sheet.getRow(j).getCell(0).getStringCellValue()))   
                {                  
                    k++;
                    j++;
                }

                //passing batches of data to different funtion to analyze sheet 
                DateDiff(temp,temp+(k-1));                                      
                TimeDiff(temp,temp+(k-1));
                WorkHours(temp,temp+(k-1));

                k=0;
            }
        }catch(Exception ep){
            System.out.println(ep);  }

        //three for loops to print analyzed data to console

        System.out.println("\nEmployees Details who has worked for 7 consecutive days.\n\nPosition ID   Employe Name\n");    
        for(int p=0;p<i;p++)
            System.out.println(consDay[p]);
        
        System.out.println("\n********************************************************************************************");
        
        System.out.println("\n\nEmployees Details who have less than 10 hours of time between shifts but greater than 1 hour.\n\nPosition ID   Employe Name\n");    
        for(i=0;i<m;i++)
            System.out.println(ShiftDiff[i]);    

        System.out.println("\n********************************************************************************************");
        
        System.out.println("\n\nEmployees Details Who has worked for more than 14 hours in a single shift.\n\nPosition ID   Employe Name\n");    
        for(i=0;i<n;i++)
            System.out.println(ShiftWork[i]);        

    }

    //funtion to anaylze who has worked for 7 consecutive days
    void DateDiff(int j,int k)  
    {
        //storing Name and position ID of Employe
        id=sheet.getRow(j).getCell(0).getStringCellValue();
        id = id +"     "+sheet.getRow(j).getCell(7).getStringCellValue();   
        int cnt=0,flg=0;

        // loop to scan dATA 
        while(j<k-1)       
        {
            // to CHECK that fetch value is numeric
            if (sheet.getRow(j).getCell(2).getCellType() == CellType.NUMERIC  && sheet.getRow(j+1).getCell(2).getCellType() == CellType.NUMERIC) 
            { 
                //converting fetch value to LocalDateTime        
                LocalDate startDate =  (sheet.getRow(j).getCell(2).getDateCellValue()).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();    
                LocalDate endDate =  (sheet.getRow(j).getCell(2).getDateCellValue()).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();

               //to scan the next date and to avoid ambiguity 
               while(startDate.equals(endDate))     
               {
                    endDate =  (sheet.getRow(++j).getCell(2).getDateCellValue()).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                    flg=1;
                }
                //calculating duration of two dates 
                long daysDifference = ChronoUnit.DAYS.between(startDate, endDate);  

                //difference is 1 means the date are consecutive
                if(daysDifference==1)   
                    cnt++;
                else
                    cnt=0;
             }
            if(flg==0)
                j++;

            flg=0;     
        }

        //storing the the employe details in array who has worked for 7 consecutive days
        if(cnt==7)            
            consDay[i++]=id;   
    }

    // funtion to analyze who have less than 10 hours of time between shifts but greater than 1 hour
    void TimeDiff(int j,int k)  
    {
        //storing Name and position ID of Employe
        id=sheet.getRow(j).getCell(0).getStringCellValue();
        id = id +"     "+sheet.getRow(j).getCell(7).getStringCellValue();   
        int flg=0;
     
        // loop to scan dATA 
        while(j<k)           
        {
            // to CHECK that fetch value is numeric
            if (sheet.getRow(j).getCell(3).getCellType() == CellType.NUMERIC  && sheet.getRow(j+1).getCell(2).getCellType() == CellType.NUMERIC) 
            {    
                //converting fetch value to LocalDateTime     
                LocalDateTime startTime =  (sheet.getRow(j).getCell(3).getDateCellValue()).toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();    
                LocalDateTime endTime =  (sheet.getRow(j+1).getCell(2).getDateCellValue()).toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
                
                //calculating duration between two time 
                Duration duration = Duration.between(startTime, endTime);   

                //converting hours to miniute
                long minutes = duration.toMinutes() ;   
                    
                //condition for hours
                if(minutes>60 && minutes<600)   
                    flg=1;      
                   
                j++;
              }
              else
                j++;
        }

        //storing the the employe details in array who have less than 10 hours of time between shifts but greater than 1 hour
        if(flg==1)                
            ShiftDiff[m++]=id;   
    }

    // funtion to analyze Who has worked for more than 14 hours in a single shift
    void WorkHours(int j,int k)     
    {
        //storing Name and position ID of Employe
        id=sheet.getRow(j).getCell(0).getStringCellValue();
        id = id +"     "+sheet.getRow(j).getCell(7).getStringCellValue();   
        int flg=0;
     
        // loop to scan dATA 
        while(j<k)           
        {
            // to CHECK that fetch value is String
            if (sheet.getRow(j).getCell(4).getCellType() == CellType.STRING)        
            { 
                //taking working hours 
                String hourString=sheet.getRow(j).getCell(4).getStringCellValue(); 

                //splitting value in hours and miniutes
                String[] hourParts = hourString.split(":");                     
        
                // Extract hours and minutes from the array
                int hours = Integer.parseInt(hourParts[0]);
                int minutes = Integer.parseInt(hourParts[1]);

                //calculating total miniutes
                int totalMinutes = hours * 60 + minutes;            

                //condition for working hours
                if(totalMinutes>840)                   
                    flg=1;

                j++;
              }
              else
                j++;
        }

        //storing the the employe details in array Who has worked for more than 14 hours in a single shift
        if(flg==1)                
            ShiftWork[n++]=id;   
    }

       public static void main(String args[])
    {
        Analyzer anz = new Analyzer();
    }
}