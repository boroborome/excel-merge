package com.balance.excel.merge;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

@RunWith(SpringRunner.class)
@SpringBootTest(classes = ExcelMergeApplication.class)
public class ExcelMergeApplicationTests {

    @Autowired
    private MergeCommand mergeCommand;

    @Test
    public void test() throws Exception {
        mergeCommand.run("/Users/gaoyushan/Downloads/2019-04/Balance", "/Users/gaoyushan/Downloads/2019-04/summary.xlsx");
    }
}
