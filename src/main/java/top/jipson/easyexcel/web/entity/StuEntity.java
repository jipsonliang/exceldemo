package top.jipson.easyexcel.web.entity;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import lombok.Data;

@Data
@TableName("stu")
public class StuEntity {
    @TableId(type = IdType.AUTO)
    private Long id;
    private String name;
    private Integer sex;
    private Integer age;
    private String classid;
    private Double score;
}
