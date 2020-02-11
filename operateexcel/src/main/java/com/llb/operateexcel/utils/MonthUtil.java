package com.llb.operateexcel.utils;

/**
 * 根据文件名来区分月份
 * @Author llb
 * Date on 2020/2/11
 */
public class MonthUtil {

    /**
     * 根据文件名来获取月份
     * @param filename
     * @return
     */
    public String getMonth(String filename) {
        int month = 0;
        String qnDate = "";
        for (int i = 1; i <= 12; i++) {
            if(filename.contains(i+"")) {
                month = i;
            }
        }
        //拼装签名日期时间
        if(month<7) {
            qnDate = "2020/" + month + "/1";
        } else {
            qnDate = "2019/" + month + "/1";
        }

        return qnDate;
    }


    public static void main(String[] args) {
        MonthUtil monthUtil = new MonthUtil();
        String month = monthUtil.getMonth("6excel1.xls");
        System.out.println(month);
    }
}
