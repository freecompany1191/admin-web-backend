package com.o2osys.mng.excel;

import org.springframework.boot.test.context.SpringBootTest;

//@RunWith(SpringRunner.class)
//@AutoConfigureMockMvc
@SpringBootTest
public class SpringBootExcelApplicationTests {
    /*
    @Autowired
    private MockMvc mvc;

    @Autowired
    ExcelReadComponent excelReadComponent;

    @Autowired
    MngMapperService mngMapperService;
     */
    /*
    @Test
    public void test_readExcel() throws IOException, InvalidFormatException {

        ClassLoader classLoader = this.getClass().getClassLoader();
        //        System.out.println("::RESOURCE PATH : "+resource.getPath());
        // System.out.println("::RESOURCE FILE : "+resource.getFile());
        // System.out.println("::RESOURCE URL : "+resource.getURL());

        File xlsxFile = new File(classLoader.getResource("files/test.xlsx").getFile());
        System.out.println("XLSX : "+xlsxFile.getPath());
        String filename = xlsxFile.getName();
        System.out.println("filename : "+filename);
        System.out.println(filename.endsWith("xlsx"));

        excelReadComponent
        .readExcelToList(new MockMultipartFile("test","test.xlsx","xlsx", new FileInputStream(xlsxFile)),
                Product::rowOf)
        .forEach(System.out::println);

    }


    @Test
    public void excelList() throws Exception {

        this.mvc.perform(get("/v1/exceltest")).andExpect(status().isOk()).andReturn();

    }
     */
}
