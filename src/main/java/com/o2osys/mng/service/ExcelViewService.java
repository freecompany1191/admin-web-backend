package com.o2osys.mng.service;

import java.sql.SQLException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.o2osys.mng.common.service.CommonService;
import com.o2osys.mng.common.util.CommonUtils;
import com.o2osys.mng.packet.manager.ReqManager;
import com.o2osys.mng.packet.manager.ResManager;

import oracle.sql.TIMESTAMP;

@Service
public class ExcelViewService {

    // 로그
    private final Logger log = LoggerFactory.getLogger(ExcelViewService.class);
    private final String TAG = ExcelViewService.class.getSimpleName();

    @Autowired
    CommonService commonService;

    @Autowired
    ManagerService managerService;


    /**
     * 엑셀 다운로드 서비스 요청
     * @Method Name : getExcel
     * @param reqManager
     * @return
     * @throws Exception
     */
    public Map<String, Object> getExcel(ReqManager reqManager) throws Exception {

        Map<String, Object> resultMap = new HashMap<String, Object>();
        ResManager resManager = new ResManager();

        resultMap = managerService.callExcelApis(reqManager);

        //데이터 조회
        ArrayList<LinkedHashMap<String,Object>> dataList = (ArrayList<LinkedHashMap<String, Object>>) resultMap.get("out_ROW1");

        //헤더 맵 생성
        //Map<String, Object> header = getHeaderMap(dataList);

        //바디 맵 생성
        //List<LinkedHashMap<String, Object>> body = convertMap(dataList);
        //log.info("#DB MnGrSt : "+CommonUtils.jsonStringFromObject(result));

        List<String> header = new ArrayList<String>();

        if(reqManager.getHeader() != null && reqManager.getHeader().length > 0){
            for(String str : reqManager.getHeader()){
                log.info("header : "+str);
                header.add(str);
            }

        }else {

            header = getDefaultHeader(dataList);

        }

        List<List<String>> body = getBody(dataList);

        resultMap.put("header", header);
        resultMap.put("body", body);
        resultMap.put("fileName", "export");

        return resultMap;

    }

    /**
     * 프로시저 결과값 컨버팅 ArrayList
     * @Method Name : convertMap
     * @param object
     * @return
     * @throws Exception
     */
    private List<String> getDefaultHeader(ArrayList<LinkedHashMap<String, Object>> rows) throws Exception {

        List<String> resultList = new ArrayList<String>();
        int[] i = {0};

        if (rows == null || rows.size() <= 0) {
            return null;
        }

        log.debug("ROW SIZE = "+rows.size());

        if (null == rows.get(0) || "NULL".equals((rows.get(0).keySet().iterator().next()))) {
            return null;
        }

        log.debug("ROW.get(0) SIZE = "+rows.get(0).size());
        rows.get(0).forEach((k,v)->{

            if (k != null) {
                //log.debug("## KEYNAME : "+keyName+" | VALUE : "+row.get(keyName)+" | class : "+row.get(keyName).getClass().getName());
                //log.info("NAME : "+row.get(k).getClass().getName());

                //log.info("LIST ["+i[0]+"] KEY : "+ k+", VALUE : "+k);
                resultList.add(k);

            }
            i[0]++;

        });

        return resultList;
    }


    /**
     * 프로시저 결과값 컨버팅 ArrayList
     * @Method Name : convertMap
     * @param object
     * @return
     * @throws Exception
     */
    private List<List<String>> getBody(ArrayList<LinkedHashMap<String, Object>> rows) throws Exception {
        List<List<String>> resultList = new ArrayList<List<String>>();

        if (rows == null || rows.size() <= 0) {
            return null;
        }

        log.info("ROW SIZE = "+rows.size());

        if (null == rows.get(0) || "NULL".equals((rows.get(0).keySet().iterator().next()))) {
            return null;
        }

        int[] i = {0};
        for (LinkedHashMap<String, Object> row : rows) {
            List<String> dataList = new ArrayList<String>();

            row.forEach((k,v)->{

                try {
                    if (k != null) {
                        //log.debug("## KEYNAME : "+keyName+" | VALUE : "+row.get(keyName)+" | class : "+row.get(keyName).getClass().getName());
                        //log.info("NAME : "+row.get(k).getClass().getName());

                        if(v != null){
                            if ("java.sql.Timestamp".equals(row.get(k).getClass().getName())) {
                                /*
                            SimpleDateFormat dateFormat3 = new SimpleDateFormat("yyyy-MM-dd a hh:mm:ss");
                            String returntime = dateFormat3.format(row.get(keyName));
                                 */
                                Timestamp time = (Timestamp) row.get(k);
                                String datetime = CommonUtils.localToDate(time.toLocalDateTime(), null);
                                //log.info("LIST java.sql.Timestamp ["+i[0]+"] KEY : "+ k+", VALUE : "+datetime);
                                dataList.add(datetime);

                            } else if("oracle.sql.TIMESTAMP".equals(row.get(k).getClass().getName())) {

                                TIMESTAMP ts = (TIMESTAMP) row.get(k);
                                Timestamp time = ts.timestampValue();

                                String datetime = CommonUtils.localToDate(time.toLocalDateTime(), null);
                                //log.info("LIST oracle.sql.TIMESTAMP ["+i[0]+"] KEY : "+ k+", VALUE : "+datetime);
                                dataList.add(datetime);

                            } else {
                                //log.debug("["+REQ_TYPE+"]["+i[0]+"] KEYNAME : "+ k);

                                //log.info("LIST ["+i[0]+"] KEY : "+ k+", VALUE : "+v);
                                dataList.add(String.valueOf(v));

                            }

                        } else {
                            //log.info("LIST NULL ["+i[0]+"] KEY : "+ k+", VALUE : "+v);
                            dataList.add("");
                        }


                    }
                } catch (SQLException e) {
                    e.printStackTrace();
                } catch (Exception e) {
                    e.printStackTrace();
                }

            });

            resultList.add(dataList);
            i[0]++;
        }

        return resultList;
    }



    /**
     * 프로시저 결과값 컨버팅 ArrayList
     * @Method Name : convertMap
     * @param object
     * @return
     * @throws Exception
     */
    private Map<String, Object> getHeaderMap(ArrayList<LinkedHashMap<String, Object>> rows) throws Exception {

        Map<String, Object> resultMap = new LinkedHashMap<String, Object>();
        int[] i = {0};

        if (rows == null) {
            return new LinkedHashMap<String, Object>();
        }

        if (rows == null || rows.size() <= 0) {
            return new LinkedHashMap<String, Object>();
        }

        log.info("ROW SIZE = "+rows.size());

        if (null == rows.get(0) || "NULL".equals((rows.get(0).keySet().iterator().next()))) {
            return new LinkedHashMap<String, Object>();
        }

        rows.get(0).forEach((k,v)->{

            if (k != null) {
                //log.debug("## KEYNAME : "+keyName+" | VALUE : "+row.get(keyName)+" | class : "+row.get(keyName).getClass().getName());
                //log.info("NAME : "+row.get(k).getClass().getName());

                //log.info("LIST ["+i[0]+"] KEY : "+ k+", VALUE : "+k);
                resultMap.put(k, k);

            }
            i[0]++;

        });

        return resultMap;
    }

    /**
     * 프로시저 결과값 컨버팅 ArrayList
     * @Method Name : convertMap
     * @param object
     * @return
     * @throws Exception
     */
    private ArrayList<LinkedHashMap<String, Object>> convertMap(ArrayList<LinkedHashMap<String, Object>> rows) throws Exception {
        if (rows == null) {
            return new ArrayList<LinkedHashMap<String, Object>>();
        }

        if (rows == null || rows.size() <= 0) {
            return new ArrayList<LinkedHashMap<String, Object>>();
        }

        log.info("ROW SIZE = "+rows.size());

        if (null == rows.get(0) || "NULL".equals((rows.get(0).keySet().iterator().next()))) {
            return new ArrayList<LinkedHashMap<String, Object>>();
        }

        ArrayList<LinkedHashMap<String, Object>> arrayList = new ArrayList<LinkedHashMap<String, Object>>();

        int[] i = {0};
        for (LinkedHashMap<String, Object> row : rows) {
            LinkedHashMap<String, Object> resultMap = new LinkedHashMap<String, Object>();

            row.forEach((k,v)->{

                try {
                    if (k != null) {
                        //log.debug("## KEYNAME : "+keyName+" | VALUE : "+row.get(keyName)+" | class : "+row.get(keyName).getClass().getName());
                        //log.info("NAME : "+row.get(k).getClass().getName());



                        if(v != null){
                            if ("java.sql.Timestamp".equals(row.get(k).getClass().getName())) {
                                /*
                            SimpleDateFormat dateFormat3 = new SimpleDateFormat("yyyy-MM-dd a hh:mm:ss");
                            String returntime = dateFormat3.format(row.get(keyName));
                                 */
                                Timestamp time = (Timestamp) row.get(k);
                                String datetime = CommonUtils.localToDate(time.toLocalDateTime(), null);
                                //log.info("LIST java.sql.Timestamp ["+i[0]+"] KEY : "+ k+", VALUE : "+datetime);
                                resultMap.put(k, datetime);

                            } else if("oracle.sql.TIMESTAMP".equals(row.get(k).getClass().getName())) {

                                TIMESTAMP ts = (TIMESTAMP) row.get(k);
                                Timestamp time = ts.timestampValue();

                                String datetime = CommonUtils.localToDate(time.toLocalDateTime(), null);
                                //log.info("LIST oracle.sql.TIMESTAMP ["+i[0]+"] KEY : "+ k+", VALUE : "+datetime);
                                resultMap.put(k, datetime);

                            } else {
                                //log.debug("["+REQ_TYPE+"]["+i[0]+"] KEYNAME : "+ k);

                                //log.info("LIST ["+i[0]+"] KEY : "+ k+", VALUE : "+v);
                                resultMap.put(k, String.valueOf(v));

                            }

                        } else {
                            //log.info("LIST ["+i[0]+"] KEY : "+ k+", VALUE : "+v);
                            resultMap.put(k, "");
                        }


                    }
                } catch (SQLException e) {
                    e.printStackTrace();
                } catch (Exception e) {
                    e.printStackTrace();
                }

            });

            /*
            Iterator<String> iterator = row.keySet().iterator();

            String keyName;

            while (iterator.hasNext()) {
                keyName = iterator.next();

                log.debug("["+REQ_TYPE+"]["+i+"] KEYNAME : "+ keyName);

                if (row.get(keyName) != null){

                    log.debug("["+REQ_TYPE+"]["+i+"] KEY : "+ keyName+", VALUE : "+row.get(keyName));
                    resultMap.put(keyName, String.valueOf(row.get(keyName)));

                }else {
                    log.debug("["+REQ_TYPE+"]["+i+"] KEY : "+ keyName+", VALUE : ");
                    resultMap.put(keyName, "");
                }

            }
             */

            arrayList.add(resultMap);
            i[0]++;
        }

        return arrayList;
    }
}
