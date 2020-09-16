# utl-adding-a-second-ods-excel-created-sheet-to-a-closed-ods-excel-workbook
Adding a second ods excel created sheet to a closed ods excel workbook 
    Adding a second ods excel created sheet to a closed ods excel workbook                                                           
                                                                                                                                     
       Method                                                                                                                        
          1.  Create a sheet, 'females', using ods excelin workbook d:/xls/odsXlsAddReport.xlsx                                      
          2.  Create a second ods excel sheet, 'males' in a temp workbook d:/xls/temp.xlsx                                           
          3.  Copy the sheet 'males' from workbook d:/xls/temp.xlsx to workbook d:/xls/odsXlsAddReport.xlsx                          
                                                                                                                                     
      MACRO COPY SHEET ON END AND IN LINK                                                                                            
                                                                                                                                     
    github                                                                                                                           
    https://tinyurl.com/y3vglcab                                                                                                     
    https://github.com/rogerjdeangelis/utl-adding-a-second-ods-excel-created-sheet-to-a-closed-ods-excel-workbook                    
                                                                                                                                     
    macros                                                                                                                           
    https://tinyurl.com/y9nfugth                                                                                                     
    https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories/tree/master                           
                                                                                                                                     
    SAS Forum                                                                                                                        
    https://tinyurl.com/y3rrxz5o                                                                                                     
    https://communities.sas.com/t5/SAS-Enterprise-Guide/Is-it-possible-to-append-data-using-ods-excel/m-p/684247                     
                                                                                                                                     
    related repos                                                                                                                    
    https://tinyurl.com/y3j8h43z                                                                                                     
    https://github.com/rogerjdeangelis?tab=repositories&q=copysheet+in%3Areadme&type=&language=                                      
                                                                                                                                     
                                                                                                                                     
    /*                   _                                                                                                           
    (_)_ __  _ __  _   _| |_                                                                                                         
    | | `_ \| `_ \| | | | __|                                                                                                        
    | | | | | |_) | |_| | |_                                                                                                         
    |_|_| |_| .__/ \__,_|\__|                                                                                                        
            |_|                                                                                                                      
    */                                                                                                                               
                                                                                                                                     
    TWO SEPARATE WORKBOOKS                                                                                                           
                                                                                                                                     
      d:/xls/odsXlsAddReport.xlsx                                                                                                    
                                                                                                                                     
          +---------------------------------------------------+                                                                      
          |     A      |    B       |     C      |    D       |                                                                      
          +---------------------------------------------------+                                                                      
       1  | NAME       |   SEX      |    AGE     |  HEIGHT    |                                                                      
          +------------+------------+------------+------------+                                                                      
       2  | ALICE      |    F       |    13      |   66       |                                                                      
          +------------+------------+------------+------------+                                                                      
           ...                                                                                                                       
          +------------+------------+------------+------------+                                                                      
       9  | JANE       |    F       |    15      |   68.5     |                                                                      
          +------------+------------+------------+------------+                                                                      
                                                                                                                                     
          [FEMALES]                                                                                                                  
                                                                                                                                     
                                                                                                                                     
      d:/xls/temp.xlsx                                                                                                               
                                                                                                                                     
          +---------------------------------------------------+                                                                      
          |     A      |    B       |     C      |    D       |                                                                      
          +---------------------------------------------------+                                                                      
       1  | NAME       |   SEX      |    AGE     |  HEIGHT    |                                                                      
          +------------+------------+------------+------------+                                                                      
       2  | ALFRED     |    M       |    14      |    69      |                                                                      
          +------------+------------+------------+------------+                                                                      
           ...                                                                                                                       
          +------------+------------+------------+------------+                                                                      
       11 | WILLIAM    |    M       |    15      |   66.5     |                                                                      
          +------------+------------+------------+------------+                                                                      
                                                                                                                                     
          [MALES]                                                                                                                    
                                                                                                                                     
    /*           _               _                                                                                                   
      ___  _   _| |_ _ __  _   _| |_                                                                                                 
     / _ \| | | | __| `_ \| | | | __|                                                                                                
    | (_) | |_| | |_| |_) | |_| | |_                                                                                                 
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                                                
                    |_|                                                                                                              
    */                                                                                                                               
                                                                                                                                     
    COPY 'MALES' sheet in temp worbook to workbook d:/xls/odsXlsAddReport.xlsx                                                       
                                                                                                                                     
     SHEET FEMALES                                                                                                                   
                                                                                                                                     
      d:/xls/odsXlsAddReport.xlsx                                                                                                    
                                                                                                                                     
                                                                                                                                     
          +---------------------------------------------------+                                                                      
          |     A      |    B       |     C      |    D       |                                                                      
          +---------------------------------------------------+                                                                      
       1  | NAME       |   SEX      |    AGE     |  HEIGHT    |                                                                      
          +------------+------------+------------+------------+                                                                      
       2  | ALICE      |    F       |    13      |   66       |                                                                      
          +------------+------------+------------+------------+                                                                      
           ...                                                                                                                       
          +------------+------------+------------+------------+                                                                      
       9  | JANE       |    F       |    15      |   68.5     |                                                                      
          +------------+------------+------------+------------+                                                                      
                                                                                                                                     
          [FEMALES]                                                                                                                  
                                                                                                                                     
     SHEET MALES (IN SAME WORKBOOK)                                                                                                  
                                                                                                                                     
          +---------------------------------------------------+                                                                      
          |     A      |    B       |     C      |    D       |                                                                      
          +---------------------------------------------------+                                                                      
       1  | NAME       |   SEX      |    AGE     |  HEIGHT    |                                                                      
          +------------+------------+------------+------------+                                                                      
       2  | ALFRED     |    M       |    14      |    69      |                                                                      
          +------------+------------+------------+------------+                                                                      
           ...                                                                                                                       
          +------------+------------+------------+------------+                                                                      
       11 | WILLIAM    |    M       |    15      |   66.5     |                                                                      
          +------------+------------+------------+------------+                                                                      
                                                                                                                                     
          [MALES]                                                                                                                    
                                                                                                                                     
    /*         _       _   _                                                                                                         
     ___  ___ | |_   _| |_(_) ___  _ __                                                                                              
    / __|/ _ \| | | | | __| |/ _ \| `_ \                                                                                             
    \__ \ (_) | | |_| | |_| | (_) | | | |                                                                                            
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|                                                                                            
                                                                                                                                     
    */                                                                                                                               
                                                                                                                                     
    * just in case they exist then delete them;                                                                                      
                                                                                                                                     
    %utlfkil(d:/xls/odsXlsAddReport.xlsx);                                                                                           
    %utlfkil(d:/xls/temp.xlsx);                                                                                                      
                                                                                                                                     
    * females sheet;                                                                                                                 
    ods excel file="d:/xls/odsXlsAddReport.xlsx" options(sheet_name="female");                                                       
    proc report data=sashelp.class(obs=3 where=(sex="F"));                                                                           
    run;quit;                                                                                                                        
    ods excel close;                                                                                                                 
                                                                                                                                     
    * males sheet;                                                                                                                   
    ods excel file="d:/xls/temp.xlsx"  options(sheet_name="male");   ;                                                               
    proc report data=sashelp.class(obs=3 where=(sex="M"));                                                                           
    run;quit;                                                                                                                        
    ods excel close;                                                                                                                 
                                                                                                                                     
    * males shhet to females workbook;                                                                                               
    %utl_copysheet(                                                                                                                  
     frombook  = d:/xls/temp.xlsx                                                                                                    
    ,fromsheet = male                                                                                                                
    ,tobook    = d:/xls/odsXlsAddReport.xlsx                                                                                         
    );                                                                                                                               
                                                                                                                                     
    /*                                                                                                                               
    | | ___   __ _                                                                                                                   
    | |/ _ \ / _` |                                                                                                                  
    | | (_) | (_| |                                                                                                                  
    |_|\___/ \__, |                                                                                                                  
             |___/                                                                                                                   
    */                                                                                                                               
                                                                                                                                     
    3060   %utlfkil(d:/xls/odsXlsAddReport.xlsx);                                                                                    
    MLOGIC(UTLFKIL):  Beginning execution.                                                                                           
    MLOGIC(UTLFKIL):  Parameter UTLFKIL has value d:/xls/odsXlsAddReport.xlsx                                                        
    MLOGIC(UTLFKIL):  %LOCAL  URC                                                                                                    
    MLOGIC(UTLFKIL):  %LET (variable name is URC)                                                                                    
    SYMBOLGEN:  Macro variable UTLFKIL resolves to d:/xls/odsXlsAddReport.xlsx                                                       
    SYMBOLGEN:  Macro variable URC resolves to 0                                                                                     
    SYMBOLGEN:  Macro variable FNAME resolves to #LN04232                                                                            
    MLOGIC(UTLFKIL):  %IF condition &urc = 0 and %sysfunc(fexist(&fname)) is TRUE                                                    
    MLOGIC(UTLFKIL):  %LET (variable name is URC)                                                                                    
    SYMBOLGEN:  Macro variable FNAME resolves to #LN04232                                                                            
    MLOGIC(UTLFKIL):  %LET (variable name is URC)                                                                                    
    MPRINT(UTLFKIL):   run;                                                                                                          
    MLOGIC(UTLFKIL):  Ending execution.                                                                                              
    3061   %utlfkil(d:/xls/temp.xlsx);                                                                                               
    MLOGIC(UTLFKIL):  Beginning execution.                                                                                           
    MLOGIC(UTLFKIL):  Parameter UTLFKIL has value d:/xls/temp.xlsx                                                                   
    MLOGIC(UTLFKIL):  %LOCAL  URC                                                                                                    
    MLOGIC(UTLFKIL):  %LET (variable name is URC)                                                                                    
    SYMBOLGEN:  Macro variable UTLFKIL resolves to d:/xls/temp.xlsx                                                                  
    SYMBOLGEN:  Macro variable URC resolves to 0                                                                                     
    SYMBOLGEN:  Macro variable FNAME resolves to #LN04233                                                                            
    MLOGIC(UTLFKIL):  %IF condition &urc = 0 and %sysfunc(fexist(&fname)) is TRUE                                                    
    MLOGIC(UTLFKIL):  %LET (variable name is URC)                                                                                    
    SYMBOLGEN:  Macro variable FNAME resolves to #LN04233                                                                            
    MLOGIC(UTLFKIL):  %LET (variable name is URC)                                                                                    
    MPRINT(UTLFKIL):   run;                                                                                                          
    MLOGIC(UTLFKIL):  Ending execution.                                                                                              
    3062   * females sheet;                                                                                                          
    3063   ods excel file="d:/xls/odsXlsAddReport.xlsx" options(sheet_name="female");                                                
    3064   proc report data=sashelp.class(obs=3 where=(sex="F"));                                                                    
    3065   run;                                                                                                                      
                                                                                                                                     
    NOTE: Multiple concurrent threads will be used to summarize data.                                                                
    NOTE: There were 3 observations read from the data set SASHELP.CLASS.                                                            
          WHERE sex='F';                                                                                                             
    NOTE: PROCEDURE REPORT used (Total process time):                                                                                
          real time           0.11 seconds                                                                                           
          user cpu time       0.04 seconds                                                                                           
          system cpu time     0.06 seconds                                                                                           
          memory              11986.48k                                                                                              
          OS Memory           45640.00k                                                                                              
          Timestamp           09/16/2020 05:18:16 PM                                                                                 
          Step Count                        1095  Switch Count  0                                                                    
                                                                                                                                     
                                                                                                                                     
    3065 !     quit;                                                                                                                 
    3066   ods excel close;                                                                                                          
    NOTE: Writing EXCEL file: d:/xls\odsXlsAddReport.xlsx                                                                            
    3067   * males sheet;                                                                                                            
    3068   ods excel file="d:/xls/temp.xlsx"  options(sheet_name="male");   ;                                                        
    3069   proc report data=sashelp.class(obs=3 where=(sex="M"));                                                                    
    3070   run;                                                                                                                      
                                                                                                                                     
    NOTE: Multiple concurrent threads will be used to summarize data.                                                                
    NOTE: There were 3 observations read from the data set SASHELP.CLASS.                                                            
          WHERE sex='M';                                                                                                             
    NOTE: PROCEDURE REPORT used (Total process time):                                                                                
          real time           0.05 seconds                                                                                           
          user cpu time       0.00 seconds                                                                                           
          system cpu time     0.03 seconds                                                                                           
          memory              12017.00k                                                                                              
          OS Memory           45384.00k                                                                                              
          Timestamp           09/16/2020 05:18:17 PM                                                                                 
          Step Count                        1096  Switch Count  0                                                                    
                                                                                                                                     
                                                                                                                                     
    3070 !     quit;                                                                                                                 
    3071   ods excel close;                                                                                                          
    NOTE: Writing EXCEL file: d:/xls\temp.xlsx                                                                                       
    3072   * males shhet to females workbook;                                                                                        
    3073   %utl_copysheet(                                                                                                           
    MLOGIC(UTL_COPYSHEET):  Beginning execution.                                                                                     
    3074    frombook  = d:/xls/temp.xlsx                                                                                             
    3075   ,fromsheet = male                                                                                                         
    3076   ,tobook    = d:/xls/odsXlsAddReport.xlsx                                                                                  
    3077   );                                                                                                                        
    MLOGIC(UTL_COPYSHEET):  Parameter FROMBOOK has value d:/xls/temp.xlsx                                                            
    MLOGIC(UTL_COPYSHEET):  Parameter FROMSHEET has value male                                                                       
    MLOGIC(UTL_COPYSHEET):  Parameter TOBOOK has value d:/xls/odsXlsAddReport.xlsx                                                   
    MLOGIC(UTL_COPYSHEET):  %LOCAL  __CMD                                                                                            
    MPRINT(UTL_COPYSHEET):   proc sql;                                                                                               
    MPRINT(UTL_COPYSHEET):   create table __utl_copysheet (chr char(80));                                                            
    NOTE: Table WORK.__UTL_COPYSHEET created, with 0 rows and 1 columns.                                                             
    SYMBOLGEN:  Macro variable FROMBOOK resolves to d:/xls/temp.xlsx                                                                 
    SYMBOLGEN:  Macro variable TOBOOK resolves to d:/xls/odsXlsAddReport.xlsx                                                        
    SYMBOLGEN:  Macro variable FROMSHEET resolves to male                                                                            
    MPRINT(UTL_COPYSHEET):  insert into __utl_copysheet VALUES("$xl = new-object -c excel.application                                
    'd:/xls/odsXlsAddReport.xlsx' # destination's fullpath                             ") VALUES("$xl = new-object -c excel.applicati
    VALUES("$wb1 = $xl.workbooks.open($file1, $null, $true) # open source, readonly ") VALUES("$wb2 = $xl.workbooks.open($file2) # op
    VALUES("$sheetToCopy = $wb1.sheets.item('male') # source sheet to copy") VALUES("$sheetToCopy.copy($sh1_wb2) # copy source sheet 
    VALUES("$wb2.close($true) # close and save destination workbook                 ") VALUES("$xl.quit()                            
    WARNING: Character expression will be truncated when assigned to character column chr.                                           
    NOTE: 14 rows were inserted into WORK.__UTL_COPYSHEET.                                                                           
                                                                                                                                     
    MPRINT(UTL_COPYSHEET):  quit;                                                                                                    
    NOTE: PROCEDURE SQL used (Total process time):                                                                                   
          real time           0.03 seconds                                                                                           
          user cpu time       0.00 seconds                                                                                           
          system cpu time     0.03 seconds                                                                                           
          memory              5386.87k                                                                                               
          OS Memory           40212.00k                                                                                              
          Timestamp           09/16/2020 05:18:17 PM                                                                                 
          Step Count                        1097  Switch Count  0                                                                    
                                                                                                                                     
                                                                                                                                     
    MLOGIC(UTLFKIL):  Beginning execution.                                                                                           
    MLOGIC(UTLFKIL):  Parameter UTLFKIL has value f:\wrk\_TD8996_T7610_\ps1.ps1                                                      
    MLOGIC(UTLFKIL):  %LOCAL  URC                                                                                                    
    MLOGIC(UTLFKIL):  %LET (variable name is URC)                                                                                    
    SYMBOLGEN:  Macro variable UTLFKIL resolves to f:\wrk\_TD8996_T7610_\ps1.ps1                                                     
    SYMBOLGEN:  Macro variable URC resolves to 0                                                                                     
    SYMBOLGEN:  Macro variable FNAME resolves to #LN04370                                                                            
    MLOGIC(UTLFKIL):  %IF condition &urc = 0 and %sysfunc(fexist(&fname)) is TRUE                                                    
    MLOGIC(UTLFKIL):  %LET (variable name is URC)                                                                                    
    SYMBOLGEN:  Macro variable FNAME resolves to #LN04370                                                                            
    MLOGIC(UTLFKIL):  %LET (variable name is URC)                                                                                    
    MPRINT(UTLFKIL):   run;                                                                                                          
    MLOGIC(UTLFKIL):  Ending execution.                                                                                              
    MPRINT(UTL_COPYSHEET):  ;                                                                                                        
    MPRINT(UTL_COPYSHEET):   filename _ps1 "f:\wrk\_TD8996_T7610_\ps1.ps1";                                                          
    MPRINT(UTL_COPYSHEET):   data _null_;                                                                                            
    MPRINT(UTL_COPYSHEET):   length cmd $4096;                                                                                       
    MPRINT(UTL_COPYSHEET):   file _ps1;                                                                                              
    MPRINT(UTL_COPYSHEET):   set __utl_copysheet;                                                                                    
    MPRINT(UTL_COPYSHEET):   put chr;                                                                                                
    MPRINT(UTL_COPYSHEET):   putlog chr;                                                                                             
    MPRINT(UTL_COPYSHEET):   if _n_=1 then do;                                                                                       
    MPRINT(UTL_COPYSHEET):   cmd=catx(' ',"'powershell -Command",cats('"',"f:\wrk\_TD8996_T7610_\ps1.ps1",cats('"',"'")));           
    MPRINT(UTL_COPYSHEET):   putlog cmd=;                                                                                            
    MPRINT(UTL_COPYSHEET):   call symputx('__cmd',cmd);                                                                              
    MPRINT(UTL_COPYSHEET):   end;                                                                                                    
    MPRINT(UTL_COPYSHEET):   run;                                                                                                    
                                                                                                                                     
    NOTE: The file _PS1 is:                                                                                                          
          Filename=f:\wrk\_TD8996_T7610_\ps1.ps1,                                                                                    
          RECFM=V,LRECL=384,File Size (bytes)=0,                                                                                     
          Last Modified=16Sep2020:17:18:17,                                                                                          
          Create Time=16Sep2020:16:39:56                                                                                             
                                                                                                                                     
    $xl = new-object -c excel.application                                                                                            
    cmd='powershell -Command "f:\wrk\_TD8996_T7610_\ps1.ps1"'                                                                        
    $file1 = 'd:/xls/temp.xlsx' # source's fullpath                                                                                  
    $file2 = 'd:/xls/odsXlsAddReport.xlsx' # destination's fullpath                                                                  
    $xl = new-object -c excel.application                                                                                            
    $xl.displayAlerts = $false # don't prompt the user                                                                               
    $wb1 = $xl.workbooks.open($file1, $null, $true) # open source, readonly                                                          
    $wb2 = $xl.workbooks.open($file2) # open target                                                                                  
    $sh1_wb2 = $wb2.sheets.item(1) # first sheet in destination workbook                                                             
    $sheetToCopy = $wb1.sheets.item('male') # source sheet to copy                                                                   
    $sheetToCopy.copy($sh1_wb2) # copy source sheet to destination workbook                                                          
    $wb1.close($false) # close source workbook w/o saving                                                                            
    $wb2.close($true) # close and save destination workbook                                                                          
    $xl.quit()                                                                                                                       
    spps -n excel                                                                                                                    
    NOTE: 14 records were written to the file _PS1.                                                                                  
          The minimum record length was 10.                                                                                          
          The maximum record length was 71.                                                                                          
    NOTE: There were 14 observations read from the data set WORK.__UTL_COPYSHEET.                                                    
    NOTE: DATA statement used (Total process time):                                                                                  
          real time           0.06 seconds                                                                                           
          user cpu time       0.03 seconds                                                                                           
          system cpu time     0.03 seconds                                                                                           
          memory              595.21k                                                                                                
          OS Memory           35088.00k                                                                                              
          Timestamp           09/16/2020 05:18:18 PM                                                                                 
          Step Count                        1098  Switch Count  0                                                                    
                                                                                                                                     
                                                                                                                                     
    MPRINT(UTL_COPYSHEET):  quit;                                                                                                    
    MPRINT(UTL_COPYSHEET):   options xwait xsync;                                                                                    
    NOTE: There are no active tasks/transactions.                                                                                    
    MPRINT(UTL_COPYSHEET):  run;                                                                                                     
    MPRINT(UTL_COPYSHEET):  quit;                                                                                                    
    MPRINT(UTL_COPYSHEET):   systask kill _ps1;                                                                                      
    SYMBOLGEN:  Macro variable __CMD resolves to 'powershell -Command "f:\wrk\_TD8996_T7610_\ps1.ps1"'                               
    MPRINT(UTL_COPYSHEET):   systask command 'powershell -Command "f:\wrk\_TD8996_T7610_\ps1.ps1"' taskname=_ps1;                    
    MPRINT(UTL_COPYSHEET):   waitfor _ps1;                                                                                           
    NOTE: Task "_ps1" produced no LOG/Output.                                                                                        
    MLOGIC(UTL_COPYSHEET):  Ending execution.                                                                                        
                                                                                                                                     
    /*                                                                                                                               
     _ __ ___   __ _  ___ _ __ ___                                                                                                   
    | `_ ` _ \ / _` |/ __| `__/ _ \                                                                                                  
    | | | | | | (_| | (__| | | (_) |                                                                                                 
    |_| |_| |_|\__,_|\___|_|  \___/                                                                                                  
                                                                                                                                     
    */                                                                                                                               
                                                                                                                                     
    %macro utl_copysheet(                                                                                                            
    frombook=c:\top\xls&pgm._100rpt.xlsx                                                                                             
    ,tobook=c:\top\xls&pgm._200rpt.xlsx                                                                                              
    ,fromsheet=utl_200rpt                                                                                                            
    )/ des="Copy a sheet from one workbook to the another workbook";                                                                 
                                                                                                                                     
    %local __cmd;                                                                                                                    
                                                                                                                                     
    /*                                                                                                                               
      For testing without macro call                                                                                                 
      %let frombook=c:\top\xls\&pgm._100rpt.xlsx;                                                                                    
      %let tobook=c:\top\xls\&pgm._200rpt.xlsx;                                                                                      
      %let fromsheet=utl_200rpt;                                                                                                     
      %let tosheet=utl_100rpt;                                                                                                       
    */                                                                                                                               
                                                                                                                                     
    proc sql;                                                                                                                        
     create                                                                                                                          
       table __utl_copysheet (chr char(80));insert into __utl_copysheet                                                              
    VALUES("$xl = new-object -c excel.application                                   ")                                               
    VALUES("$file1 = '&frombook' # source's fullpath                                ")                                               
    VALUES("$file2 = '&tobook' # destination's fullpath                             ")                                               
    VALUES("$xl = new-object -c excel.application                                   ")                                               
    VALUES("$xl.displayAlerts = $false # don't prompt the user                      ")                                               
    VALUES("$wb1 = $xl.workbooks.open($file1, $null, $true) # open source, readonly ")                                               
    VALUES("$wb2 = $xl.workbooks.open($file2) # open target                         ")                                               
    VALUES("$sh1_wb2 = $wb2.sheets.item(1) # first sheet in destination workbook    ")                                               
    VALUES("$sheetToCopy = $wb1.sheets.item('&fromsheet') # source sheet to copy    ")                                               
    VALUES("$sheetToCopy.copy($sh1_wb2) # copy source sheet to destination workbook ")                                               
    VALUES("$wb1.close($false) # close source workbook w/o saving                   ")                                               
    VALUES("$wb2.close($true) # close and save destination workbook                 ")                                               
    VALUES("$xl.quit()                                                              ")                                               
    VALUES("spps -n excel                                                           ")                                               
    ;quit;                                                                                                                           
                                                                                                                                     
    %utlfkil(%sysfunc(pathname(work))\ps1.ps1);                                                                                      
                                                                                                                                     
    filename _ps1 "%sysfunc(pathname(work))\ps1.ps1";                                                                                
    data _null_;                                                                                                                     
      length cmd $4096;                                                                                                              
      file _ps1;                                                                                                                     
      set __utl_copysheet;                                                                                                           
      put chr;                                                                                                                       
      putlog chr;                                                                                                                    
      if _n_=1 then do;                                                                                                              
        cmd=catx(' ',"'powershell -Command",cats('"',"%sysfunc(pathname(work))\ps1.ps1",cats('"',"'")));                             
        putlog cmd=;                                                                                                                 
        call symputx('__cmd',cmd);                                                                                                   
      end;                                                                                                                           
    run;quit;                                                                                                                        
                                                                                                                                     
    options xwait xsync;run;quit;                                                                                                    
    systask kill _ps1;                                                                                                               
    systask command &__cmd taskname=_ps1;                                                                                            
    waitfor _ps1;                                                                                                                    
    %mend utl_copysheet;                                                                                                             
                                                                                                                                     
