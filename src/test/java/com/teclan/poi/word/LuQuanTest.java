package com.teclan.poi.word;

import org.javalite.activejdbc.DB;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class LuQuanTest {
    private static final Logger LOGGER = LoggerFactory.getLogger(LuQuanTest.class);

    //String filePath="C:\\Users\\teclan\\Desktop\\luquan\\务工人员花名册（2018）.doc";
    //String filePath="C:\\Users\\teclan\\Desktop\\luquan\\务工人员花名册2019年.doc";
    //String filePath="C:\\Users\\teclan\\Desktop\\luquan\\务工人员花名册2019年第二批.doc";
//    String filePath="C:\\Users\\teclan\\Desktop\\luquan\\务工人员花名册2019年第三批.doc";

    String filePath="C:\\Users\\teclan\\Desktop\\luquan\\务工人员花名册（新）.doc";


    @Test
    public void read(){

        new DB("default").open("com.mysql.cj.jdbc.Driver","jdbc:mysql://rm-bp1k0ef85lmpf2821zo.mysql.rds.aliyuncs.com/luquan?useUnicode=true&characterEncoding=UTF-8&useSSL=false&serverTimezone=UTC","root","Ali123!@#");

        LuQuanWordTableHandler tableHandler = new LuQuanWordTableHandler();
        WorlUtils.readTable(filePath,tableHandler);


        for(Person person:tableHandler.getPeople()) {

            try {

                new DB("default").exec("insert person_all (`name`,`sex`,`type`,`id_card`,`location`,`enter_time`,`out_time`,`cont_no`,`wage`,`card_no`,`bank`,`phone`,`memo`) values (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        person.getName(), person.getSex(), person.getType(), person.getIdCard(), person.getLocation(), person.getEnterTime(), person.getOutTime(), person.getContNo(),
                        person.getWage(), person.getCardNo(), person.getBank(), person.getPhone(), person.getMemo());
            }catch (Exception e){
                LOGGER.error(e.getMessage(),e);
            }
        }

        new DB("default").close();

    }






}
