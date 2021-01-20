package com.ds.skeleton_backend.utils;

import com.ds.skeleton_backend.model.common.ExcelModel;
import com.ds.skeleton_backend.model.entity.DocumentEntity;
import com.ds.skeleton_backend.repository.DocumentMongoRepository;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.stereotype.Component;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

@Slf4j
@Component
public class FileHelper {

        private final DocumentMongoRepository repository;

    public FileHelper(DocumentMongoRepository repository) {
        this.repository = repository;
    }

    public ByteArrayInputStream wordWriter(String title){

        List<DocumentEntity> entities = repository.findByTitle(title);

        if(entities.isEmpty()){
            throw new RuntimeException("Title Not Found");
        }
        try(        XWPFDocument document = new XWPFDocument();
                    ByteArrayOutputStream out = new ByteArrayOutputStream();
        ) {

            entities.forEach(entity->{
                entity.getLines().forEach(line->{
                    XWPFParagraph para = document.createParagraph();
                    para.setAlignment(ParagraphAlignment.RIGHT);
                    XWPFRun paragraphRun = para.createRun();
                    paragraphRun.setText(line);
                    paragraphRun.setColor("000000");
                    paragraphRun.setFontSize(14);
                });

                XWPFParagraph footer = document.createParagraph();
                footer.setAlignment(ParagraphAlignment.RIGHT);
                XWPFRun footerRun = footer.createRun();
                footerRun.setText("Page Number:"+entity.getPageNumber());
                footerRun.setColor("000000");
                footerRun.setFontSize(12);
            } );



            document.write(out);
            return  new ByteArrayInputStream(out.toByteArray());
        }
        catch (IOException e) {
            throw new RuntimeException(e.getMessage());
        }


    }

    public ByteArrayInputStream excelWriter(List<ExcelModel> excelModelList, List<String> paths){
        try(Workbook workbook = new XSSFWorkbook();
            ByteArrayOutputStream out = new ByteArrayOutputStream();
        ) {

            //create Sheet
            log.info("Creating Sheet");
            createSheet(excelModelList,workbook,paths);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());

        } catch (IOException e) {
            throw new RuntimeException("writing to excel failed");
        }
    }

    private void createSheet(List<ExcelModel> excelModelList, Workbook workbook, List<String> paths) {

        //Sheet Name
        Sheet sheet = workbook.createSheet("Applications Report");

        //PrepareHeaders
        List<String> headers = getConstantReportsHeaders();
        Map<Integer, String> dynamicHeadersNumbering =new HashMap<>();
        log.info("DynamicHeaders Map Created");
        log.info(("Checking if Paths content"));
        if(paths != null) {
            dynamicHeadersNumbering= addDynamicReportHeadersNumbering(paths, headers);
        }
        log.info("After Checking The dynamicHeadersNumbering={}",dynamicHeadersNumbering);
        log.info("*******************************");
        log.info("Writing Headers");
        writeSheetHeaders(sheet,headers);
        log.info("Headers Written successfully");
        log.info("*******************************");
        log.info("Writing data To Excel Sheet");
       writeDataToSheet(sheet,excelModelList,headers,dynamicHeadersNumbering);

        fixSizing(sheet,headers);
    }

    private Map<Integer,String> addDynamicReportHeadersNumbering(List<String> paths, List<String> headers) {
        log.info("HeadersNumbering Starts with value paths={},headers={}",paths,headers);
        Map<Integer,String> headersNumbering = new HashMap<>();
        int currentNumbering = headers.size();

        log.info("Dynamic Numbering Starts from Column No = {}",currentNumbering);

        for (String path : paths) {
            headersNumbering.put(currentNumbering++,path);
            log.info("currentNumbering={} , path ={} added To HeadersNumbering Map",currentNumbering,path);
            headers.add(path);
            log.info("path={} - added to headers List",path);
        }

        log.info("addDynamicReportHeadersNumbering is returning headersNumbering={}",headersNumbering);
        return headersNumbering;
    }

    private void writeDataToSheet(Sheet sheet, List<ExcelModel> excelModelList, List<String> headers, Map<Integer, String> dynamicHeadersNumbering) {
        log.info("excelModelList={}",excelModelList);
        log.info("headers={}",headers);
        log.info("headers.size={}",headers.size());
        log.info("dynamicHeadersNumbering={}",dynamicHeadersNumbering);
     for (int rowNumber=0; rowNumber<excelModelList.size(); rowNumber++){
         Row currentRow = sheet.createRow(rowNumber+1);
         log.info("currentRow in Excel={}",rowNumber+1);
         //fix THis to do loops for Constants headers
         for(int columnNumber=0 ; columnNumber<headers.size();columnNumber++ ){
             log.info("Current Column in Excel ={}",columnNumber);
             Cell cell = currentRow.createCell(columnNumber);
             switch (columnNumber){
                 case 0 : cell.setCellValue(new Date(System.currentTimeMillis()));break;
                 default:break;
             }
                     }
            log.info("Checking dynamicHeaders if null");
         if(dynamicHeadersNumbering !=null){
             log.info("****Not Null******");
             Set<Integer> columnNumbersKey = dynamicHeadersNumbering.keySet();
             log.info("Numbers Of Columns for Dynamic Data ={}",columnNumbersKey);
             for (Integer key : columnNumbersKey) {
                 String fieldValue="";
                 ExcelModel excelModel = excelModelList.get(rowNumber);
                 log.info("ExcelModel={}",excelModel);
                 Map<String, Object> fields = excelModel.getFields();
                 log.info("Fields={}",fields);
                 log.info("**** Checking If Fields is Null");
                 if(fields!=null){
                     log.info("Fields Not null");
                     Object objValue = fields.getOrDefault(dynamicHeadersNumbering.get(key), " ");
                     log.info("Field Value (Object) ={}",objValue);
                     fieldValue=checkObject(objValue);
                        log.info("Setting Cell Value with fieldValue={}",fieldValue);
                     currentRow.createCell(key).setCellValue(fieldValue);
                 }
             }
         }

         }
     log.info("Excel Is ready to Download");
     }

    private String checkObject(Object objValue) {
        log.info("checking Object");
        String objValueAsString="";
        if(objValue != null) {
            if (objValue instanceof Collection) {
                log.info("Object Is a Collection and combining it to String");
                objValueAsString = ((Collection<?>) objValue).stream().map(o -> o.toString())
                        .collect(Collectors.joining("/"));
            } else {
                log.info("Object Is a primitive and converting to String");
                objValueAsString = objValue.toString();
            }
        }
        log.info("After Checking Object we return objValueAsString={}",objValueAsString);
        return objValueAsString;
    }


    private void fixSizing(Sheet sheet, List<String> headers) {
        for (int columnNumber=0; columnNumber<headers.size(); columnNumber++){
        sheet.autoSizeColumn(columnNumber);
        }
    }

    private void writeSheetHeaders(Sheet sheet, List<String> headers) {
        Row headerRow = sheet.createRow(0);
        log.info("Headers={}",headers);
        for(int columnNumber=0; columnNumber< headers.size();columnNumber++){
            Cell cell= headerRow.createCell(columnNumber);
            cell.setCellValue(headers.get(columnNumber));
        }
    }

    private List<String> getConstantReportsHeaders() {
        //Prepare Headers in a List
        List<String> headers= new ArrayList<>();
        headers.add("Created Date");
        log.info("Headers={} / Constants Headers Created",headers);
        return headers;
    }
}




/*************************************************************************/

    @GetMapping("/getCustomer")
    public ResponseEntity<InputStreamResource>getCustomer(
            @RequestParam(value = "jsonFieldsPaths",required =false) List<String> paths,
            @RequestParam(value = "jsonField",required = false)  String jsonField
            ,@RequestParam(value = "searchValue",required = false) String searchValue){

        System.out.println("paths = " + paths);

        List<ExcelModel> excelModels = new ArrayList<>();
        Query query = new Query();
        query.addCriteria(Criteria.where(jsonField).is(searchValue));
        List<CustomerEntity> customerEntities = mongoTemplate.find(query, CustomerEntity.class);

        mapCustomerEntityToExcelModel(customerEntities,excelModels,paths);
        System.out.println("excelModels = " + excelModels);
        ByteArrayInputStream byteArrayInputStream = fileHelper.excelWriter(excelModels,paths);

        InputStreamResource resource = new InputStreamResource(byteArrayInputStream);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename=test.xlsx")
                .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
                .body(resource);
    }

    private void mapCustomerEntityToExcelModel(List<CustomerEntity> customerEntities, List<ExcelModel> excelModels, List<String> fields) {
        customerEntities.forEach(customerEntity -> {
            Map<String,Object> fieldsMap = new HashMap<>();

            if(fields != null){
                fieldsMap=getSelectedFieldsAsMap(fields,customerEntity);//map<fieldName,Value>
                System.out.println("fieldsMap = " + fieldsMap);
            }


            excelModels.add( ExcelModel.builder()
//                    .applicationId(customerEntity.getApplicationId())
//                    .customerId(customerEntity.getCustomerInfo().getCustomerId())
                    .fields(fieldsMap.isEmpty()? null : fieldsMap)
                    .build());
        });


    }

    private Map<String, Object> getSelectedFieldsAsMap(List<String> fields, Object object) {
        Map<String, Object> target = new HashMap<>(); // <FieldName,Value>

        for (String field : fields) {
                if(!field.contains(".")){
                    Object fieldValue = getFieldValue(object, field);
                    if(fieldValue != null){
                        target.put(field, fieldValue);
                    }
                    else{
                        target.put(field," ");
                    }
                }
                else {
                    List<String> nextNestedObjectName = new ArrayList<>();
                    String currentObjectName = field.substring(0, field.indexOf("."));
                    String nestedObjectNames= field.substring(field.indexOf(".")+1);
                    nextNestedObjectName.add(nestedObjectNames);
                    //check nulls
                    Object fieldValueObject = getFieldValue(object, currentObjectName);
                    if (fieldValueObject != null){
                        if(fieldValueObject instanceof Collection){
                            Collection<Object> objectCollection =(Collection<Object>)fieldValueObject;
                            List<Map<String, Object>> selectedFieldFromCollection = getSelectedFieldFromCollection(objectCollection, nextNestedObjectName);
                           // target.put(currentObjectName,selectedFieldFromCollection);
                            target.put(field,selectedFieldFromCollection);
                        }
                        else {
                            Map<String, Object> selectedFieldsAsMap = getSelectedFieldsAsMap(nextNestedObjectName, fieldValueObject);
                            //target.putAll(selectedFieldsAsMap);
                            Set<String> keySet = selectedFieldsAsMap.keySet();
                            Object[] keys = keySet.toArray();
                            if(keys.length > 0){
                                target.put(field,selectedFieldsAsMap.getOrDefault(keys[0],null));
                            }

                        }

                    }
                }
        }
        return target;
    }

    private List<Map<String,Object>> getSelectedFieldFromCollection(Collection<Object> objectCollection, List<String> nextNestedObjectName) {

        List<Map<String,Object>> selectedFieldFromCollection = new ArrayList<>();
        for (Object obj: objectCollection) {
                selectedFieldFromCollection.add(getSelectedFieldsAsMap(nextNestedObjectName,obj));
        }
        return selectedFieldFromCollection;
    }

    //return Value of Given Field
    private Object getFieldValue(Object source,String fieldName){
        try{
            Field sourceField = source.getClass().getDeclaredField(fieldName);
            sourceField.setAccessible(true);
            return sourceField.get(source);
        } catch (NoSuchFieldException | IllegalAccessException e) {
            e.printStackTrace();
            throw new RuntimeException("Incorrect Field Name Provided");
        }
    }
