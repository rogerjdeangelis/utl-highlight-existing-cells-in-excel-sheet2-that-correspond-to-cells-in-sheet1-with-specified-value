# utl-highlight-existing-cells-in-excel-sheet2-that-correspond-to-cells-in-sheet1-with-specified-value
Highlight existing cells in sheet2 that corresponding cells in sheet1 have specified values

    Highlight existing cells in sheet2 that corresponding cells in sheet1 have specified values                                                                
                                                                                                                                                               
    Output workbook                                                                                                                                            
    https://tinyurl.com/rgz5xcx                                                                                                                                
    https://github.com/rogerjdeangelis/utl-highlight-existing-cells-in-excel-sheet2-that-correspond-to-cells-in-sheet1-with-specified-value/blob/master/output.
                                                                                                                                                               
    Input workbook                                                                                                                                             
    https://tinyurl.com/uupqn8w                                                                                                                                
    https://github.com/rogerjdeangelis/utl-highlight-existing-cells-in-excel-sheet2-that-correspond-to-cells-in-sheet1-with-specified-value/blob/master/have.xl
                                                                                                                                                               
    githib  program                                                                                                                                            
    https://tinyurl.com/t3txw3e                                                                                                                                
    https://github.com/rogerjdeangelis/utl-highlight-existing-cells-in-excel-sheet2-that-correspond-to-cells-in-sheet1-with-specified-value                    
                                                                                                                                                               
    Cells in sheet2 that correspond to ones in sheet1 will be hilighted (numbers in red)                                                                       
                                                                                                                                                               
    *_                   _                                                                                                                                     
    (_)_ __  _ __  _   _| |_                                                                                                                                   
    | | '_ \| '_ \| | | | __|                                                                                                                                  
    | | | | | |_) | |_| | |_                                                                                                                                   
    |_|_| |_| .__/ \__,_|\__|                                                                                                                                  
            |_|                                                                                                                                                
    ;                                                                                                                                                          
                                                                                                                                                               
    * create input;                                                                                                                                            
    * delete input workbook if it exists;                                                                                                                      
    %utlfkil(e:/xls/have.xlsx); * delete if exists;                                                                                                            
                                                                                                                                                               
    data template(keep=rec t:) values(keep=f:);                                                                                                                
     retain rec;                                                                                                                                               
     array TFS[9] $32 T1-T9;                                                                                                                                   
     array  FS[9] $32 F1-F9;                                                                                                                                   
                                                                                                                                                               
     do lyn = 1 to 5;                                                                                                                                          
                                                                                                                                                               
        do ltr=1 to dim(tfs);                                                                                                                                  
           if  RAND("Bernoulli", .25)=1 then tfs[ltr]=cats("~S={foreground=red}T");  /* 1s with .25 probability */                                             
           else tfs[ltr]="F";                                                                                                                                  
        end;                                                                                                                                                   
        rec=1;                                                                                                                                                 
        output template;                                                                                                                                       
                                                                                                                                                               
        do ltr=1 to dim(tfs);                                                                                                                                  
            fs[ltr] = put(int(9*uniform(1357)) + 1,1.);                                                                                                        
        end;                                                                                                                                                   
        rec=lyn;                                                                                                                                               
        output values;                                                                                                                                         
     end;                                                                                                                                                      
     drop ltr lyn;                                                                                                                                             
     stop;                                                                                                                                                     
                                                                                                                                                               
    run;quit;                                                                                                                                                  
                                                                                                                                                               
    %utlfkil(e:/xls/have.xlsx); * delete if exist;                                                                                                             
                                                                                                                                                               
    ods listing;                                                                                                                                               
    ods escapechar="~";                                                                                                                                        
    ods excel file="e:/xls/have.xlsx" style=pearl;                                                                                                             
                                                                                                                                                               
    ods excel options(sheet_name="template");                                                                                                                  
                                                                                                                                                               
    proc report data=template missing                                                                                                                          
     style(header)=[fontsize=12pt]                                                                                                                             
     style(column)=[fontsize=11pt protectspecialchars=off just=c] split="_";                                                                                   
    run;quit;                                                                                                                                                  
                                                                                                                                                               
    ods excel options(sheet_name="VALUES");                                                                                                                    
    proc report data=values missing                                                                                                                            
     style(header)=[fontsize=12pt]                                                                                                                             
     style(column)=[fontsize=11pt protectspecialchars=off just=c] split="_";                                                                                   
    run;quit;                                                                                                                                                  
                                                                                                                                                               
    ods excel close;                                                                                                                                           
                                                                                                                                                               
    Excel e:/xls/map.xlsx                                                                                                                                      
                                                                                                                                                               
        +---------------------------+                                                                                                                          
        |   |A |B |C | ... |X |Y |Z |                                                                                                                          
        |---+--+--+--+-----+--+--+--|                                                                                                                          
      1 |REC|A |B |C | ... |X |Y |Z |                                                                                                                          
        |---+--+--+--+-----+--+--+--|                                                                                                                          
      2 | 1 | F| F| F| ... | F| F| F|   * T has background color red;                                                                                          
        |---+--+--+--+--------+--+--|                                                                                                                          
      3 | 1 | T| F| T| ... | F| F| F|                                                                                                                          
        |---+--+--+--+--------+--+--|                                                                                                                          
      4 | 1 | F| F| T| ... | T| F| T|                                                                                                                          
        |---+--+--+--+--------+--+--|                                                                                                                          
      5 | 1 | F| F| F| ... | F| F| F|                                                                                                                          
        |---+--+--+--+--------+--+--|                                                                                                                          
      5 | 1 | F| T| T| ... | F| F| F|                                                                                                                          
        +---------------------------+                                                                                                                          
                                                                                                                                                               
     [SHEET1]                                                                                                                                                  
                                                                                                                                                               
        +---------------------------+                                                                                                                          
        |   |A |B |C | ... |X |Y |Z |                                                                                                                          
        |---+--+--+--+-----+--+--+--|                                                                                                                          
      1 |REC|A |B |C | ... |X |Y |Z |                                                                                                                          
        |---+--+--+--+-----+--+--+--|                                                                                                                          
      2 | 1 | 1| 4| 7| ... | 7| 2| 1|                                                                                                                          
        |---+--+--+--+--------+--+--|                                                                                                                          
      3 | 2 | 9| 3| 6| ... | 5| 9| 6|                                                                                                                          
        |---+--+--+--+--------+--+--|                                                                                                                          
      4 | 3 | 3| 8| 5| ... | 7| 7| 3|                                                                                                                          
        |---+--+--+--+--------+--+--|                                                                                                                          
      5 | 4 | 3| 9| 7| ... | 4| 5| 8|                                                                                                                          
        |---+--+--+--+--------+--+--|                                                                                                                          
      5 | 5 | 8| 6| 8| ... | 3| 4| 5|                                                                                                                          
        +---------------------------+                                                                                                                          
                                                                                                                                                               
      [SHEET2]                                                                                                                                                 
                                                                                                                                                               
    *            _               _                                                                                                                             
      ___  _   _| |_ _ __  _   _| |_                                                                                                                           
     / _ \| | | | __| '_ \| | | | __|                                                                                                                          
    | (_) | |_| | |_| |_) | |_| | |_                                                                                                                           
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                                                                          
                    |_|                                                                                                                                        
    ;                                                                                                                                                          
        +---------------------------+                                                                                                                          
        |   |A |B |C | ... |X |Y |Z |                                                                                                                          
        |---+--+--+--+-----+--+--+--|                                                                                                                          
      1 |REC|A |B |C | ... |X |Y |Z |                                                                                                                          
        |---+--+--+--+-----+--+--+--|  Asterisk                                                                                                                
      2 | 1 | 1| 4| 7| ... | 7| 2| 1|  * foreground=red                                                                                                        
        |---+--+--+--+--------+--+--|    numbers are in red                                                                                                    
      3 | 2 |*9| 3|*6| ... | 5| 9| 6|                                                                                                                          
        |---+--+--+--+--------+--+--|                                                                                                                          
      4 | 3 | 3| 8|*5| ... |*7| 7|*3|                                                                                                                          
        |---+--+--+--+--------+--+--|                                                                                                                          
      5 | 4 | 3| 6| 7| ... | 4| 5| 8|                                                                                                                          
        |---+--+--+--+--------+--+--|                                                                                                                          
      5 | 5 | 8| 6| 8| ... | 3| 4| 5|                                                                                                                          
        +---------------------------+                                                                                                                          
                                                                                                                                                               
      [SHEET1]                                                                                                                                                 
                                                                                                                                                               
       * foreground=red                                                                                                                                        
         numbers are in red                                                                                                                                    
                                                                                                                                                               
    *          _       _   _                                                                                                                                   
     ___  ___ | |_   _| |_(_) ___  _ __                                                                                                                        
    / __|/ _ \| | | | | __| |/ _ \| '_ \                                                                                                                       
    \__ \ (_) | | |_| | |_| | (_) | | | |                                                                                                                      
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|                                                                                                                      
                                                                                                                                                               
    ;                                                                                                                                                          
                                                                                                                                                               
    * you will need to rename f1-f26 to a--z;                                                                                                                  
    * this is easily done with three of my macros;                                                                                                             
    * interleave spreadsheets;                                                                                                                                 
                                                                                                                                                               
    libname xel "e:/xls/have.xlsx";                                                                                                                            
                                                                                                                                                               
    data want(drop=idx f1-f9 rename=(fx1-fx9=tf1-tf9));                                                                                                        
                                                                                                                                                               
      * create the rename statement;                                                                                                                           
      if mod(_N_,2)=1 then set xel.'template$'n;   * names rec t1-t9;                                                                                          
      else set xel.'values$'n;                     * names rec f1-f9;                                                                                          
                                                                                                                                                               
      array vars[9] $32 t1-t9;                                                                                                                                 
      array fs[9]   f1-f9;                                                                                                                                     
      array hs[9]   $32 fx1-fx9;                                                                                                                               
                                                                                                                                                               
      do idx=1 to dim(vars);                                                                                                                                   
                                                                                                                                                               
        if lag(vars[idx])="T" then hs[idx]=cats("~S={foreground=red}",vars[idx]);                                                                              
        else hs[idx]=put(fs[idx],best.);                                                                                                                       
                                                                                                                                                               
      end;                                                                                                                                                     
                                                                                                                                                               
      if mod(_N_,2)=0;                                                                                                                                         
                                                                                                                                                               
    run;quit;                                                                                                                                                  
    libname xel clear;                                                                                                                                         
                                                                                                                                                               
    %utlfkil(e:/xls/high.xlsx);                                                                                                                                
                                                                                                                                                               
    ods excel file="e:/xls/output.xlsx" style=pearl;                                                                                                           
                                                                                                                                                               
    ods excel options(sheet_name="template");                                                                                                                  
    proc report data=template missing                                                                                                                          
     style(header)=[fontsize=12pt]                                                                                                                             
     style(column)=[fontsize=11pt protectspecialchars=off just=c] split="_";                                                                                   
    run;quit;                                                                                                                                                  
                                                                                                                                                               
    ods listing;                                                                                                                                               
    ods escapechar="~";                                                                                                                                        
    ods excel options(sheet_name="red" sheet_interval="none");                                                                                                 
    proc report data=want(keep=rec tf1-tf9) missing                                                                                                            
     style(header)=[fontsize=12pt]                                                                                                                             
     style(column)=[fontsize=11pt protectspecialchars=off] split="_";                                                                                          
    run;quit;                                                                                                                                                  
                                                                                                                                                               
    ods excel close;                                                                                                                                           
                                                                                                                                                               
