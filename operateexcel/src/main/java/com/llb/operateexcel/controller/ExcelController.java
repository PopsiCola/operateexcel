package com.llb.operateexcel.controller;

import ch.qos.logback.core.util.FileUtil;
import com.llb.operateexcel.utils.ExcelUtil;
import com.llb.operateexcel.utils.MonthUtil;
import org.springframework.stereotype.Controller;
import org.springframework.util.MultiValueMap;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @Author llb
 * Date on 2020/2/9
 */
@Controller
public class ExcelController {

    @RequestMapping("/index")
    public String showIndex() {
        return "index";
    }


    @RequestMapping("/merge")
    public String mergeExcel(MultipartHttpServletRequest request, HttpServletResponse response) {
        //获取科室名称
        String ksName = request.getParameter("ksName");

        int month = 7;

        //获取文件
        MultiValueMap<String, MultipartFile> maps = request.getMultiFileMap();
        List<MultipartFile> files = maps.get("excels");

        boolean isFirst = true;
        //上传的文件
        String destFile = "F:\\";
        //生成的文件
        if(!files.isEmpty()) {
            //上传文件的目录
            List<String> fileNames = new ArrayList<String>();
            //保存图片
            for (MultipartFile file : files) {

                //判断是否是excel
                String fileType = file.getOriginalFilename().substring(file.getOriginalFilename().lastIndexOf(".")+1);
                String sourceFile = "F:\\excel\\"+file.getOriginalFilename();
                if(!new File("F:\\excel").exists()) {
                    new File("F:\\excel").mkdirs();
                }
                System.out.println("上传文件的路径：" + sourceFile);

                if("xlsx".equals(fileType) || "xls".equals(fileType) ) {
                    //根据文件名来获取改文件的月份
                    String qmDate = new MonthUtil().getMonth(file.getOriginalFilename());
                    //签名日期时间
                    //                String qnDate = "2019/"+ month++ +"/1";

                    try {
                        //判断是否是第一个excel，如果是第一个，则直接保存，从第二个开始依次读取插入第一个excel中
                        if(isFirst) {
                            isFirst = false;
                            destFile = destFile+file.getOriginalFilename();
                            file.transferTo(new File(destFile));
                            new ExcelUtil().changeExcel(destFile, ksName, qmDate);
                        } else {
                            file.transferTo(new File(sourceFile));
                            new ExcelUtil().importFromExcel(sourceFile, destFile, ksName, qmDate);
                        }
                        request.setAttribute("msg", "合并成功，保存文件在:"+destFile);
                    } catch (Exception e) {
                        e.printStackTrace();
                        request.setAttribute("msg", "合并失败");
                        System.out.println("合并失败");
                    }
                } else {
                    request.setAttribute("msg", "请上传.xlsx或.xls文件");
                    return "index";
                }
            }
        } else{
            request.setAttribute("msg", "请上传.xlsx或.xls文件");
            return "index";
        }
        return "index";
    }

}
