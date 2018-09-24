public class ExcelSheet{
    
    public string viewType{get; set;}
    public string selQtr{get; set;}
    public Decimal selNumOfQtr{get; set;}
    public Decimal startYear{get; set;}
    
    
    public string PORId{get; set;}
    
    public Boolean sheetPanel{get; set;}
    public Boolean createPORTemplateBtn{get; set;}
    public Boolean savePORTemplateBtn{get; set;}
    public Boolean editPORTemplateBtn{get; set;}
    public Boolean deletePORTemplateBtn{get; set;}
    public Boolean newVersionPORTempBtn{get; set;}
    
    public Boolean inputView{get; set;}
    public Boolean outputView{get; set;}
    
    public list<Template_Item__c> tempItemList{get; set;}
    public list<POR_Detail__c> PORDtlsList{get; set;}
    public POR__c PORRec{get; set;}


    public list<excelSheetWrap> xlSttWrapList{get; set;}
    public Map<string, string> temItmFormlaMap{get; set;}

    //public list<string> fiscalQtrList{get; set;}
    public string finalJSONOutput{get; set;}
    public finalExcelSheetWrap finalExcelSheetWrp{get; set;}
    
    
    public class finalExcelSheetWrap{
        public string PORName{get; set;}
        public string PORPartNumber{get; set;}
        public string PORversion{get; set;}
        public string PORversionType{get; set;}
        public string PORDescription{get; set;}
        public string PORTemplateName{get; set;}
        public string PORTemplateId{get; set;}

        public list<excelSheetWrap> xlSttWrapList{get; set;}
        public list<string> xlSttHrWrap{get; set;}
        public list<string> xlStFormulaFld{get; set;}
        public list<excelcolWrap> xlSttcolDataType{get; set;}

        public finalExcelSheetWrap(){
            xlSttWrapList = new list<excelSheetWrap>();
            xlSttHrWrap = new list<string>();
            xlStFormulaFld = new list<string>();
            xlSttcolDataType = new list<excelcolWrap>();
        }
    }

    public class excelcolWrap{
        public string data{get; set;}
        public string type{get; set;}
        public Boolean allowEmpty{get; set;}
        public Boolean readOnly{get; set;}
        public numericFormat numericFormat{get; set;}
    }

    public class numericFormat{
        public string pattern{get; set;}
        public string culture{get; set;}
    }

    public class excelSheetWrap{
        public Decimal sno{get; set;}
        public string costtype{get; set;}
        public string entrytype{get; set;}
        public Map<string, string> quarters{get; set;}
        public string formula{get; set;}
        public string tempItemId{get; set;}
        public string PORDtlId{get; set;}
    }

    public List<SelectOption> getnoofQuarter() {
        List<SelectOption> options = new List<SelectOption>();
        for(Integer i=1; i<13; i++){
            options.add(new SelectOption(string.valueOf(i),string.valueOf(i)));
        }
        return options;
    }


    public List<SelectOption> getstartQuarter() {
        List<SelectOption> options = new List<SelectOption>();
        options.add(new SelectOption('Q1','Q1'));
        options.add(new SelectOption('Q2','Q2'));
        options.add(new SelectOption('Q3','Q3'));
        options.add(new SelectOption('Q4','Q4'));
        return options;
    }

    public List<SelectOption> getPartNumbers() {
        List<SelectOption> options = new List<SelectOption>();
        options.add(new SelectOption('Part1','Part1'));
        options.add(new SelectOption('Part2','Part2'));
        return options;
    }

     public List<SelectOption> getVersionType()
    {
        List<SelectOption> options = new List<SelectOption>();
        Schema.DescribeFieldResult fieldResult = POR__c.Version_Type__c.getDescribe();
        List<Schema.PicklistEntry> ple = fieldResult.getPicklistValues();
        for(Schema.PicklistEntry f : ple){
            options.add(new SelectOption(f.getLabel(), f.getValue()));
        }     
        return options;
    }

    public ExcelSheet(){
        PORId = apexpages.currentpage().getparameters().get('id');
        date currentQtr = system.Today();
        
        startYear = currentQtr.year(); 
        selNumOfQtr = 6;  
        selQtr = 'Q1';   

        list<string> fiscalQtrList = ExcelSheet.findQuarter(currentQtr, selNumOfQtr);
           

        if(PORId != null && PORId != ''){
            sheetPanel = true;  
            inputView = false;
            outputView = true;
            createPORTemplateBtn = false;
            savePORTemplateBtn = false;
            editPORTemplateBtn = true;
            deletePORTemplateBtn = true;
            newVersionPORTempBtn = true; 
            searchExcelSheet();
            viewType = 'ReadOnly'; 
        }else{ 
            sheetPanel = false;  
            inputView = true;
            outputView = false;
            createPORTemplateBtn = true;
            savePORTemplateBtn = false;
            editPORTemplateBtn = false;
            deletePORTemplateBtn = false;
            newVersionPORTempBtn = false; 
            viewType = 'Edit';
        }
    } 

    public void excelSheetValues(list<string> fiscalQtrList){

        finalExcelSheetWrp = new finalExcelSheetWrap();
        temItmFormlaMap = new Map<string, string>(); 
        list<excelSheetWrap> xlSttWrapLt = new list<excelSheetWrap>();
        list<string> xlStFormulaFldLt = new list<string>();
    
        system.debug('PORId++'+PORId);

        PORRec = [ SELECT Id, Name, Description__c,Number_of_Qtr__c,Part_Number__c,Start_Qtr__c,Template__c,Template__r.Name, Version__c,Version_Type__c FROM POR__c Where Id =:PORId ];
          
        finalExcelSheetWrp.PORName = PORRec.Name;
        finalExcelSheetWrp.PORTemplateName = PORRec.Template__r.Name;
        finalExcelSheetWrp.PORTemplateId = PORRec.Template__c;
        finalExcelSheetWrp.PORversionType = PORRec.Version_Type__c;
        finalExcelSheetWrp.PORversion = PORRec.Version__c;
        finalExcelSheetWrp.PORDescription = PORRec.Description__c;
        finalExcelSheetWrp.PORPartNumber = PORRec.Part_Number__c;
    
        tempItemList = [SELECT Id, Name, Cost_Type__c,Entry_Type__c,Formula__c,squence_number__c,Templates__c FROM Template_Item__c Where Templates__c =:PORRec.Template__c Order by squence_number__c  ASC];
        
        PORDtlsList = [SELECT Id,POR__c,POR_Cost__c,POR_Date__c,Which_Qtr__c,Template_Item__c,Template_Item__r.Cost_Type__c,Template_Item__r.Entry_Type__c,Template_Item__r.Formula__c,Template_Item__r.Templates__c,Template_Item__r.squence_number__c FROM POR_Detail__c Where Template_Item__r.Templates__c =:PORRec.Template__c  AND POR__c = :PORId Order by POR_Date__c ASC ];


        if(tempItemList.size()>0){
            for(Template_Item__c ti : tempItemList){
                temItmFormlaMap.put(ti.Id, ti.Formula__c);
            }


            Map<string, POR_Detail__c> tempItmCostMap = new Map<string, POR_Detail__c>();
            Map<string, POR_Detail__c> tempItmCostMap1 = new Map<string, POR_Detail__c>();
            for(string fsQtr : fiscalQtrList){
                for(POR_Detail__c podDtl : PORDtlsList){
                    if(podDtl.Which_Qtr__c == fsQtr){
                        tempItmCostMap.put(podDtl.Template_Item__c+'##'+fsQtr,  podDtl);
                        tempItmCostMap1.put(podDtl.Template_Item__c+'##'+fsQtr,  podDtl);

                    }
                } 
            }
            system.debug('tempItmCostMap+'+tempItmCostMap);
            system.debug('tempItemList+'+tempItemList);
            


            for(Template_Item__c tempIt : tempItemList){

                system.debug('tempItIdVal+'+tempIt);
                excelSheetWrap xlSttWrap = new excelSheetWrap();
                xlSttWrap.quarters = ExcelSheet.costCalculate(tempItemList, tempItmCostMap, tempIt, fiscalQtrList, selNumOfQtr);
                xlSttWrap.entrytype = tempIt.Entry_Type__c;
                xlSttWrap.costtype = tempIt.Cost_Type__c;
                xlSttWrap.sno = tempIt.squence_number__c;
                xlSttWrap.formula = tempIt.Formula__c;
                xlSttWrap.tempItemId = tempIt.Id;

                if(tempItmCostMap1.keyset().contains(tempIt.Id))
                xlSttWrap.PORDtlId = tempItmCostMap1.get(tempIt.Id).Id;

                xlSttWrapLt.add(xlSttWrap);  

                if(tempIt.Entry_Type__c == 'Formula')
                    xlStFormulaFldLt.add(tempIt.Cost_Type__c);
            } 
        }  

        finalExcelSheetWrp.xlStFormulaFld = xlStFormulaFldLt; 
        finalExcelSheetWrp.xlSttWrapList = xlSttWrapLt;
        finalExcelSheetWrp.xlSttHrWrap = ExcelSheet.getxlStHeaderList(fiscalQtrList);
        finalExcelSheetWrp.xlSttcolDataType = ExcelSheet.xlcolmnHeader(selNumOfQtr, true);
  
        finalJSONOutput = JSON.serialize(finalExcelSheetWrp);
        System.debug('====> JSON data: ' + JSON.serialize(finalExcelSheetWrp));
    }

    public static list<excelcolWrap> xlcolmnHeader(Decimal selNoOfQtr, Boolean access){

        list<excelcolWrap> xlColList = new list<excelcolWrap>(); 
        excelcolWrap xlColn = new excelcolWrap(); 
        xlColn.data = 'costtype';
        xlColn.type = 'text';
        xlColn.readOnly = true;

        xlColList.add(xlColn);  

        for (Integer i=1; i<=selNoOfQtr;i++) {
            excelcolWrap xlCol = new excelcolWrap();
            
            numericFormat nfrmt = new numericFormat();
            if(i == 3){
                nfrmt.pattern = '$0,0.00';
                nfrmt.culture = 'en-US'; // this is the default culture, set up for USD
            }else{
                nfrmt.pattern = '0,0.00 $';
                nfrmt.culture = 'de-DE'; // use this for EUR (German),
            }

            xlCol.numericFormat = nfrmt;

            xlCol.allowEmpty = false;
            xlCol.data = 'quarter'+i;
            xlCol.type = 'numeric';
            xlCol.readOnly = access;
            xlColList.add(xlCol);   
        }
        return xlColList;
    }
    
    public static Map<string, string> costCalculate(list<Template_Item__c> tempItList, Map<string, POR_Detail__c> tempItmCostMap,Template_Item__c tempItm, list<string> tempItKey, Decimal selNoOfQtr){
        
        Map<string, string> returnMap = new Map<string, string>();
        Map<Integer, string> columnNames = ExcelSheet.replaceFormulaString();

        system.debug('selNoOfQtr++'+selNoOfQtr);
        
 
        string zeroVal = '0'; 
        Integer i=0;
        for(string qtrNo : tempItKey){
            if(tempItm.Entry_Type__c != 'Formula'){

                if(tempItmCostMap.keyset().contains(tempItm.Id+'##'+qtrNo)){
                    returnMap.put('quarter'+i, string.valueOf(tempItmCostMap.get(tempItm.Id+'##'+qtrNo).POR_Cost__c));
                }else{
                    system.debug('qtrNo++'+'quarter'+i);
                    returnMap.put('quarter'+i, zeroVal);
                }
                
            }else{
                string repString = columnNames.get(i);
                returnMap.put('quarter'+i, tempItm.Formula__c.replace('r', repString));
            }  
            i++;
        }
        return returnMap;
    }


    public static Map<Integer, string> replaceFormulaString(){

        Map<Integer, string> columnNames = new Map<Integer, string>();
        columnNames.put(0, 'A');
        columnNames.put(1, 'B');
        columnNames.put(2, 'C');
        columnNames.put(3, 'D');
        columnNames.put(4, 'E');
        columnNames.put(5, 'F');
        columnNames.put(6, 'G');
        columnNames.put(7, 'H');
        columnNames.put(8, 'I');
        columnNames.put(9, 'J');
        columnNames.put(10, 'K');
        columnNames.put(11, 'L');
        columnNames.put(12, 'M');
        columnNames.put(13, 'N');
        columnNames.put(14, 'O');
        columnNames.put(15, 'P');
        columnNames.put(16, 'Q');
        columnNames.put(17, 'R');
        columnNames.put(18, 'S');
        columnNames.put(19, 'T');
        columnNames.put(20, 'U');
        columnNames.put(21, 'V');
        columnNames.put(22, 'W');
        columnNames.put(23, 'X');
        columnNames.put(24, 'Y');
        columnNames.put(25, 'Z');
        return columnNames;
    }


    public void searchExcelSheet(){

        createPORTemplateBtn = false;
        savePORTemplateBtn = false;
        editPORTemplateBtn = true;
        deletePORTemplateBtn = true;
        newVersionPORTempBtn = true;         

        date findQuarter = system.today();
        if(selQtr == 'Q1'){
            findQuarter = date.parse('1/1/'+startYear);
        }else if(selQtr == 'Q2'){
            findQuarter = date.parse('4/1/'+startYear);
        }else if(selQtr == 'Q3'){
            findQuarter = date.parse('7/1/'+startYear);
        }else if(selQtr == 'Q4'){
            findQuarter = date.parse('10/1/'+startYear);
        }
        
        viewType = 'ReadOnly';
        
        list<string> fiscalQtrList = ExcelSheet.findQuarter(findQuarter, selNumOfQtr);
        excelSheetValues(fiscalQtrList);
    }  


    public static list<string> getxlStHeaderList(list<string> fiscalQtrList){
        list<string> xlHrWrap = new list<string>();
        xlHrWrap.add('Cost Type');
        for(Integer i=0; i<fiscalQtrList.size(); i++){
            xlHrWrap.add(fiscalQtrList[i]);
        }
        return xlHrWrap;
    }

    public void upsertTempItems(){
       /* POR__c porObj = new POR__c();
        porObj.Id = PORId;
        porObj.Template__c = finalExcelSheetWrp.PORTemplateId;
        porObj.Part_Number__c = finalExcelSheetWrp.PORPartNumber;
        porObj.Start_Qtr__c = system.Today();
        porObj.Version__c = finalExcelSheetWrp.PORversion;
        porObj.Description__c = finalExcelSheetWrp.PORDescription;
        porObj.Version_Type__c =finalExcelSheetWrp.PORversionType;
        
        if(porObj.Template__c != null ){
            sheetPanel = true;
            upsert porObj;
            list<string>  fiscalQtrList = ExcelSheet.findQuarter(system.Today(), selNumOfQtr);
            excelSheetValues(fiscalQtrList);
        }*/

        ExcelSheet.updateExcel(finalExcelSheetWrp);
    }

    public static void updateExcel(finalExcelSheetWrap fnExcelSt){

        list<excelSheetWrap> xlSttWrapLt = new list<excelSheetWrap>();

        system.debug('fnExcelSt++'+fnExcelSt.xlSttWrapList);

       /* for(excelSheetWrap xlSttWrapLt :fnExcelSt.xlSttWrapList){
            POR_Detail__c PORDtl = new POR_Detail__c();
            PORDtl.Id = xlSttWrapLt.PORDtlId;
            PORDtl.POR_Cost__c = Decimal.valueOf(PORDtl.)

        }*/

    }

    public void deletePORTemplate(){
        POR__c porObj = new POR__c();
        porObj.Id = PORId;
        delete porObj;
    }

    public void editPORTTemplate(){
        
        createPORTemplateBtn = false;
        savePORTemplateBtn = true;
        editPORTemplateBtn = false;
        deletePORTemplateBtn = true;
        newVersionPORTempBtn = false;
        
        viewType = 'Edit';
        outputView = false;
        inputView = true;

        finalExcelSheetWrp.xlSttcolDataType = ExcelSheet.xlcolmnHeader(selNumOfQtr, false);
        finalJSONOutput = JSON.serialize(finalExcelSheetWrp);

    }

    public PageReference createPORTemplate(){
        
        createPORTemplateBtn = false;
        savePORTemplateBtn = true;
        editPORTemplateBtn = false;
        deletePORTemplateBtn = true;
        newVersionPORTempBtn = false;

        POR__c porObj = new POR__c();
      //  porObj.Template__c = PORTemplateId;
        //porObj.Part_Number__c = PORPartNumber;
       // porObj.Start_Qtr__c = system.Today();
       // porObj.Version__c = PORversion;
       // porObj.Description__c = PORDescription;
        //porObj.Version_Type__c = PORversionType;
        //porObj.Name = PORName;


        if(porObj.Template__c != null ){
            sheetPanel = true;
            upsert porObj;
            list<string> fiscalQtrList = ExcelSheet.findQuarter(system.Today(), selNumOfQtr);
            //PORTemplateId = porObj.Template__c;
            PORId = porObj.Id;
            
            excelSheetValues(fiscalQtrList);
        }
        
        PageReference myVFPage = new PageReference('/apex/Handsontable');
        myVFPage.setRedirect(true);
        myVFPage.getParameters().put('Id', porObj.Id);
        return myVFPage;
    }


     public static list<string> findQuarter(date podDate, Decimal selNumOfQtr){
        list<string> quarterList = new list<string>();
        Set<Integer> Q1 = new Set<Integer>{1,2,3};
        Set<Integer> Q2 = new Set<Integer>{4,5,6};
        Set<Integer> Q3 = new Set<Integer>{7,8,9};
        Set<Integer> Q4 = new Set<Integer>{10,11,12};

        Integer cYear = podDate.year();
        Integer cMonth = podDate.month();
        if(Q1.contains(cMonth)){
            quarterList.add(podDate.year()+'-Q1');
            quarterList.add(podDate.year()+'-Q2');
            quarterList.add(podDate.year()+'-Q3');
            quarterList.add(podDate.year()+'-Q4');
            quarterList.add(podDate.addYears(1).year()+'-Q1');
            quarterList.add(podDate.addYears(1).year()+'-Q2');
            quarterList.add(podDate.addYears(1).year()+'-Q3');
            quarterList.add(podDate.addYears(1).year()+'-Q4');
            quarterList.add(podDate.addYears(2).year()+'-Q1');
            quarterList.add(podDate.addYears(2).year()+'-Q2');
            quarterList.add(podDate.addYears(2).year()+'-Q3');
            quarterList.add(podDate.addYears(2).year()+'-Q4');
          
        }
        if(Q2.contains(cMonth)){
            quarterList.add(podDate.year()+'-Q2');
            quarterList.add(podDate.year()+'-Q3');
            quarterList.add(podDate.year()+'-Q4');
            quarterList.add(podDate.addYears(1).year()+'-Q1');
            quarterList.add(podDate.addYears(1).year()+'-Q2');
            quarterList.add(podDate.addYears(1).year()+'-Q3');
            quarterList.add(podDate.addYears(1).year()+'-Q4');
            quarterList.add(podDate.addYears(2).year()+'-Q1');
            quarterList.add(podDate.addYears(2).year()+'-Q2');
            quarterList.add(podDate.addYears(2).year()+'-Q3');
            quarterList.add(podDate.addYears(2).year()+'-Q4');
            quarterList.add(podDate.addYears(3).year()+'-Q1');
           
        }
        if(Q3.contains(cMonth)){
            quarterList.add(podDate.year()+'-Q3');
            quarterList.add(podDate.year()+'-Q4');
            quarterList.add(podDate.addYears(1).year()+'-Q1');
            quarterList.add(podDate.addYears(1).year()+'-Q2');
            quarterList.add(podDate.addYears(1).year()+'-Q3');
            quarterList.add(podDate.addYears(1).year()+'-Q4');
            quarterList.add(podDate.addYears(2).year()+'-Q1');
            quarterList.add(podDate.addYears(2).year()+'-Q2');
            quarterList.add(podDate.addYears(2).year()+'-Q3');
            quarterList.add(podDate.addYears(2).year()+'-Q4');
            quarterList.add(podDate.addYears(3).year()+'-Q1');
            quarterList.add(podDate.addYears(3).year()+'-Q2');
            
        }  
        if(Q4.contains(cMonth)){
            quarterList.add(podDate.year()+'-Q4');
            quarterList.add(podDate.addYears(1).year()+'-Q1');
            quarterList.add(podDate.addYears(1).year()+'-Q2');
            quarterList.add(podDate.addYears(1).year()+'-Q3');
            quarterList.add(podDate.addYears(1).year()+'-Q4');
            quarterList.add(podDate.addYears(2).year()+'-Q1');
            quarterList.add(podDate.addYears(2).year()+'-Q2');
            quarterList.add(podDate.addYears(2).year()+'-Q3');
            quarterList.add(podDate.addYears(2).year()+'-Q4');
            quarterList.add(podDate.addYears(3).year()+'-Q1');
            quarterList.add(podDate.addYears(3).year()+'-Q2');
            quarterList.add(podDate.addYears(3).year()+'-Q3');
           
        }
        system.debug(podDate.year()+'-QuarterList++'+quarterList);
        return quarterList;
    }

    /*Auto Complete */
    @RemoteAction  
    public static List<templateWrapper> getSearchSuggestions(String searchString){  
       List<templateWrapper> tempWrappers = new List<templateWrapper>();  
       List<List<sObject>> searchObjects = [FIND :searchString + '*' IN ALL FIELDS RETURNING Template__c (Id, Name)];  
       if(!searchObjects.isEmpty()){  
            for(List<SObject> objects : searchObjects){  
                for(SObject obj : objects){
                    if(obj.getSObjectType().getDescribe().getName().equals('Template__c')){  
                       Template__c acct = (Template__c)obj;  
                       tempWrappers.add(new templateWrapper(acct.name, acct.Id));  
                    } 
                }  
            }  
       }  
       return tempWrappers;  
    }
    
    public class templateWrapper {  
        public String label { get; set; }  
        public String value { get; set; }  
        public templateWrapper (String label, String value){  
           this.label = label;  
           this.value = value;  
        }  
    }
    
}