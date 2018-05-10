package com.o2osys.mng.service;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.o2osys.mng.common.util.CommonUtils;
import com.o2osys.mng.entity.MnGrSt;
import com.o2osys.mng.mapper.MngMapper;

/**
   @FileName  : ConMapperServiceImpl.java
   @Description : 배달대행 연동 서비스 구현체
   @author      : KMS
   @since       : 2017. 7. 18.
   @version     : 1.0

   @개정이력

   수정일           수정자         수정내용
   -----------      ---------      -------------------------------
   2017. 7. 18.     KMS            최초생성

 */
@Service("MngMapperService")
public class MngMapperService {
    // 로그
    private final Logger log = LoggerFactory.getLogger(MngMapperService.class);
    private final String TAG = MngMapperService.class.getSimpleName();

    @Autowired
    MngMapper mngMapper;


    /**
     * 연동대상 목록 출력 (배달대행요청)
     * [PKG_SY_CON_DVRY.BD_GET_CON_LIST]
     * @Method Name : getConList
     * @param version
     * @param language
     * @param device
     * @param serviceYn
     * @return
     * @throws Exception
     */
    public ArrayList<MnGrSt> mnGrStList(int pageno, int pagesize) throws Exception {
        Map<String, Object> map = new HashMap<String, Object>();

        int countPage = pagesize;
        int page = pageno;
        int startCount = (page - 1) * countPage + 1;  // 21 이 되겠죠
        int endCount = page * countPage;  // 30 이 될 겁니다.

        map.put("startCount", startCount);
        map.put("endCount", endCount);

        //연동대상 목록 출력 (배달대행요청)
        ArrayList<MnGrSt> result = mngMapper.mnGrStList(map);
        log.info("#DB MnGrSt : "+CommonUtils.jsonStringFromObject(result));

        return result;
    }


    public int mnGrStListSize() throws Exception {
        return mngMapper.mnGrStListSize();
    }

}
