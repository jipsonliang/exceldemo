package top.jipson.easyexcel.web.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import top.jipson.easyexcel.web.entity.StuEntity;
import top.jipson.easyexcel.web.service.StuService;

import java.util.ArrayList;
import java.util.List;

/**
 * 模板的读取类
 */
// 有个很重要的点 DemoDataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
public class StuListener extends AnalysisEventListener<StuEntity> {

    private StuService stuService;//不能被spring管理，所以无法自动注入

    private static final Logger LOGGER =
        LoggerFactory.getLogger(StuListener.class);
    /**
     * 每隔5条存储数据库，实际使用中可以3000条，然后清理list ，方便内存回收
     */
    private static final int BATCH_COUNT = 5;
    List<StuEntity> list = new ArrayList<StuEntity>();

    /**
     * 如果使用了spring,请使用这个构造方法。每次创建Listener的时候需要把spring管理的类传进来
     * 多例，每次都新建一个实例
     *
     * @param stuService
     */
    public StuListener(StuService stuService) {
        this.stuService = stuService;
    }

    /**
     * 这个每一条数据解析都会来调用
     *
     */
    @Override
    public void invoke(StuEntity stuEntity, AnalysisContext analysisContext) {
        LOGGER.info("解析到一条数据:{}", JSON.toJSONString(stuEntity));
        list.add(stuEntity);
        // 达到BATCH_COUNT了，需要去存储一次数据库，防止数据几万条数据在内存，容易OOM
        if (list.size() >= BATCH_COUNT) {
            saveData();
            // 存储完成清理 list
            list.clear();
        }
    }

    /**
     * 所有数据解析完成了 都会来调用
     *
     * @param context
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        // 这里也要保存数据，确保最后遗留的数据也存储到数据库
        saveData();
        LOGGER.info("所有数据解析完成！");
    }

    /**
     * 加上存储数据库
     */
    private void saveData() {
        LOGGER.info("{}条数据，开始存储数据库！", list.size());
        stuService.saveBatch(list);//自动去除了自动增长的id，SQL语句 ：INSERT INTO stu ( name, sex, age, classid, score ) VALUES ( ?, ?, ?, ?, ? )
        LOGGER.info("存储数据库成功！");
    }
}