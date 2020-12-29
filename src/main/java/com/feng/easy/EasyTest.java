package com.feng.easy;

import com.alibaba.excel.EasyExcel;
import org.junit.Test;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class EasyTest {
    static  String Path = "D:\\javaT\\feng-POI\\src\\main\\resources\\";
    private List<DemoData> data(){
        List<DemoData> list  = new ArrayList<DemoData>();
        for (int i = 0; i <10 ; i++) {
            DemoData demoData = new DemoData();
            demoData.setString("字符串"+i);
            demoData.setDoubleDate(0.35);
            demoData.setDate(new Date());
            list.add(demoData);
        }
        return list;
    }
    //根据List写入表格
    @Test
    public void simpleWrite(){
        //写法1
        String fileName = Path+"EasyTest.xlsx";
        /**
         * 这里需要指定用那个class去写，然后写到第一个sheet中，名字为模板，文件流会自动关闭。
         * write(fileName,格式类)
         * sheet(表名)
         * doWrite(数据)
         */
        EasyExcel.write(fileName,DemoData.class).sheet("模板").doWrite(data());
    }

    @Test
    public void simpleRead() {
        // 有个很重要的点 DemoDataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
        // 写法1：
        String fileName =Path+"EasyTest.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
        //重点注意

    }
}
