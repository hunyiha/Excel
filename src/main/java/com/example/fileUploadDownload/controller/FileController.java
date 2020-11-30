package com.example.fileUploadDownload.controller;


import cn.hutool.core.util.ZipUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.example.fileUploadDownload.util.TestFileUtil;
import com.example.fileUploadDownload.vo.PoliceInfo;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;

@RestController
public class FileController {

    @Value("${file.police.original}")
    private String filePoliceOriginal;

    @Value("${file.police.excel-path}")
    private String filePoliceExcelPath;

    @Value("${file.police.zip}")
    private String filePoliceZip;

    @Value("${file.crime.original}")
    private String fileCrimeOriginal;

    @Value("${file.crime.excel-path}")
    private String fileCrimeExcelPath;

    @Value("${file.crime.zip}")
    private String fileCrimeZip;



    @PostMapping("/changePoliceFromOld2New")
    public String changePoliceFromOld2New(@RequestParam("file") MultipartFile file, HttpServletResponse response) throws IOException {

        // 解压
        ZipUtil.unzip(file.getInputStream(), new File(filePoliceOriginal), Charset.defaultCharset());

        File originalFile = new File(filePoliceOriginal);

        File[] files = originalFile.listFiles();

        List<String> fileNameList = new ArrayList<>();

        for (File imageFile : files) {
            fileNameList.add(imageFile.getName());
        }

        List<PoliceInfo> policeInfos = new ArrayList<>();

        for (String fileName : fileNameList) {
            String[] nodes = fileName.substring(0,fileName.indexOf(".")).split("___");

            if(nodes.length == 3){
                PoliceInfo policeInfo = new PoliceInfo();
                policeInfo.setName(nodes[0]);
                policeInfo.setAge(nodes[1]);
                policeInfo.setLocation(nodes[2]);
                policeInfos.add(policeInfo);
            }
        }

        listFill(policeInfos);

        ZipUtil.zip("D:\\001file\\police",filePoliceZip+File.separator+ "police.zip", false);

        File downLoadFile = new File(filePoliceZip+File.separator+ "police.zip");
        if (!downLoadFile.exists()) {
            return "-1";
        }

        response.reset();
        response.setHeader("Content-Disposition", "attachment;fileName=" + "police.zip");

        try {
            InputStream inStream = new FileInputStream(filePoliceZip+File.separator+ "police.zip");
            OutputStream os = response.getOutputStream();

            byte[] buff = new byte[1024];
            int len = -1;
            while ((len = inStream.read(buff)) > 0) {
                os.write(buff, 0, len);
            }
            os.flush();
            os.close();

            inStream.close();
        } catch (Exception e) {
            e.printStackTrace();
            return "-2";
        }

        return "0";
    }

    @PostMapping("/changeCrimeFromOld2New")
    public String changeCrimeFromOld2New(@RequestParam("file") MultipartFile file, HttpServletRequest request) {

        return  "完成罪犯数据转换";
    }

    public void listFill(List<PoliceInfo> policeInfoList) {
        // 模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
        // 填充list 的时候还要注意 模板中{.} 多了个点 表示list
        String templateFileName = "D:\\000_config" + File.separator + "police.xlsx";

        // 方案1 一下子全部放到内存里面 并填充
//        String fileName = TestFileUtil.getPath() + "listFill" + System.currentTimeMillis() + ".xlsx";


        // 这里 会填充到第一个sheet， 然后文件流会自动关闭
        EasyExcel.write(filePoliceExcelPath + File.separator + "police.xlsx").withTemplate(templateFileName).sheet().doFill(policeInfoList);

        // 方案2 分多次 填充 会使用文件缓存（省内存）
       /* fileName = TestFileUtil.getPath() + "listFill" + System.currentTimeMillis() + ".xlsx";
        ExcelWriter excelWriter = EasyExcel.write(fileName).withTemplate(templateFileName).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        excelWriter.fill(policeInfoList, writeSheet);
        excelWriter.fill(policeInfoList, writeSheet);
        // 千万别忘记关闭流
        excelWriter.finish();*/
    }

}
