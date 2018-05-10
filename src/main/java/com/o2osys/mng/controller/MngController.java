package com.o2osys.mng.controller;

import java.util.ArrayList;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.o2osys.mng.common.service.CommonService;
import com.o2osys.mng.common.util.CommonUtils;
import com.o2osys.mng.entity.MnGrSt;
import com.o2osys.mng.packet.mng.ResMnGrStList;
import com.o2osys.mng.service.MngMapperService;

import io.swagger.annotations.ApiOperation;

/**
   @FileName  : BarogoController.java
   @Description : BAROGO 연동 서비스 컨트롤러
   @author      : KMS
   @since       : 2017. 9. 7.
   @version     : 1.0

   @개정이력

   수정일          수정자         수정내용
   -----------     ---------      -------------------------------
   2017. 9. 7.     KMS            최초생성

 */
@RestController
@RequestMapping(value = "/v1")
public class MngController {
    // 로그
    private final Logger log = LoggerFactory.getLogger(MngController.class);
    private final String TAG = MngController.class.getSimpleName();

    private ObjectMapper mObjectMapper = new ObjectMapper();

    /** 공통서비스 */
    @Autowired
    private CommonService commonService;

    @Autowired
    MngMapperService mngMapperSrvice;

    @ApiOperation(value = "가맹점 리스트", notes = "가맹점 리스트", response = ResMnGrStList.class)
    //    @RequestMapping(value = "/mngrsts", method = RequestMethod.GET)
    @GetMapping(path = "/mngrsts", produces = MediaType.APPLICATION_JSON_UTF8_VALUE)
    public @ResponseBody ResMnGrStList mnGrStList(
            @RequestParam(value = "pageno", required=false) int pageno,
            @RequestParam(value = "pagesize", required=false) int pagesize
            ) throws Exception {

        log.info("#pageno : "+pageno+", pagesize : "+pagesize);
        ResMnGrStList result = new ResMnGrStList();

        int totalPageSize = mngMapperSrvice.mnGrStListSize();
        ArrayList<MnGrSt> resultList = mngMapperSrvice.mnGrStList(pageno, pagesize);

        result.setPageno(pageno);
        result.setPagesize(pagesize);
        result.setTotalcount(totalPageSize);
        result.setMnGrStList(resultList);

        log.info("#ResMnGrStList : "+CommonUtils.jsonStringFromObject(result));
        return result;

    }


}