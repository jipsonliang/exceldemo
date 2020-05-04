package top.jipson.easyexcel.web.controller;

import com.alibaba.excel.EasyExcel;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import top.jipson.easyexcel.web.entity.StuEntity;
import top.jipson.easyexcel.web.listener.StuListener;
import top.jipson.easyexcel.web.service.StuService;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.List;

@RestController
@RequestMapping("/stu")
public class StuController {
    @Autowired
    private StuService stuService;

    @GetMapping("list")
    public List<StuEntity> list(){
        List<StuEntity> list = stuService.list();
        return list;
    }

    @GetMapping("download")
    public void download(HttpServletResponse response) throws IOException {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        String fileName = URLEncoder.encode("测试下载", "UTF-8");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
        EasyExcel.write(response.getOutputStream(), StuEntity.class).sheet("模板").doWrite(list());
    }

    @PostMapping("upload")
    public String upload(MultipartFile file) throws IOException {
        EasyExcel.read(file.getInputStream(), StuEntity.class, new StuListener(stuService)).sheet().doRead();
        return "upload success";
    }

}
