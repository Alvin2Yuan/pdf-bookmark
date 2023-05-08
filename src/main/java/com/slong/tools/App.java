package com.slong.tools;

import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PageMode;
import org.apache.pdfbox.pdmodel.common.PDStream;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.destination.PDPageDestination;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.destination.PDPageFitDestination;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDDocumentOutline;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDOutlineItem;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws IOException {
        String excel="D:\\zbj\\12.10卷内目录 -卷内目录----汇总.xls";
        String pdfFolder="D:\\zbj\\pdf";
        String outFolder="D:\\zbj\\out";
        System.out.println("=========开始检查pdf文件夹是否合法===========");
        checkPdf(pdfFolder);
        System.out.println("=========检查pdf文件夹完毕===========");
        System.out.println("=========开始解析Excel===========");
        ExcelReader reader= ExcelUtil.getReader(excel);
        List<Map<String,Object>> dataList= reader.readAll();
        //将记录转换为map
        Map<String,List<Map<String,Object>>> dataMap=new LinkedHashMap<String, List<Map<String, Object>>>();
        for(Map<String,Object> data:dataList){
            String ajh= data.get("案卷号").toString();
            List<Map<String,Object>> list=new ArrayList<Map<String, Object>>();
            if(dataMap.containsKey(ajh)){
                list=dataMap.get(ajh);
            }
            list.add(data);
            dataMap.put(ajh,list);
        }
        //开始处理
        for(Map.Entry<String,List<Map<String,Object>>> map: dataMap.entrySet()){
            String ajh=map.getKey();
            List<Map<String,Object>> data=map.getValue();
            File pdf=getPdf(pdfFolder,ajh);
            if(null==pdf){
                continue;
            }
            processPdf(ajh,pdf,data,outFolder);
        }



    }

    private static void processPdf(String ajh,File pdf,List<Map<String,Object>> data,String outFolder) throws IOException {

        PDDocument document=  PDDocument.load(pdf);
        PDDocumentOutline documentOutline=new PDDocumentOutline();
        document.getDocumentCatalog().setDocumentOutline(documentOutline);
        //PDOutlineItem pagesOutline=new PDOutlineItem();
        //pagesOutline.setTitle("ALL Pages");
        //documentOutline.addLast(pagesOutline);

        for(Map<String,Object> map:data){
            PDPageDestination pageDestination=new PDPageFitDestination();
            String pageStr=map.get("页码").toString();
            if(pageStr.contains("-")){
                pageStr=pageStr.substring(0,pageStr.indexOf("-"));
            }

            pageDestination.setPage(document.getPage(Integer.parseInt(pageStr)-1));

            PDOutlineItem bookmark=new PDOutlineItem();
            bookmark.setDestination(pageDestination);
            bookmark.setTitle(map.get("题名（转移登记）").toString());
            documentOutline.addLast(bookmark);
            //pagesOutline.addLast(bookmark);
            //pagesOutline.openNode();

            documentOutline.openNode();

        }
        document.getDocumentCatalog().setPageMode(PageMode.USE_OUTLINES);
        String out=outFolder+"/"+ajh;
        File outFile=new File(out);
        if(!outFile.exists()){
            outFile.mkdirs();
        }
        String fileName=out+"/"+pdf.getName();
        document.save(fileName);
    }

    private static boolean checkPdf(String pdfFolder){
        File folder=new File(pdfFolder);
        if(!folder.exists()){
            throw new RuntimeException("需要处理的文件夹不存在:"+pdfFolder);
        }
       for(File subFile: folder.listFiles()){
           //如果是案卷文件夹才检查
           if(subFile.isDirectory()){
              //检查案卷号文件夹是否只有一个pdf
               int total=0;
               for(File pdf: subFile.listFiles()){
                   if(pdf.getName().endsWith(".pdf")){
                       total++;
                   }
               }
               if(total==0){
                   throw new RuntimeException("文件夹中没有pdf:"+subFile.getName());
               }
               if(total>1){
                 throw new RuntimeException("文件夹中存在多个pdf:"+subFile.getName());
               }
           }
       }
       return true;

    }

    private static File getPdf(String pdfFolder,String ajh) {
        File folder=new File(pdfFolder);
        for(File subFile: folder.listFiles()){
            if(subFile.getName().equals(ajh.trim())){
                for(File pdf: subFile.listFiles()){
                    if(pdf.getName().endsWith(".pdf")){
                       return pdf;
                    }
                }
            }
        }
        return null;
    }
}
