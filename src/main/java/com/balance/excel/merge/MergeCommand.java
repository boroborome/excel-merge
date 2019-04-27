package com.balance.excel.merge;

import org.springframework.boot.CommandLineRunner;
import org.springframework.stereotype.Component;

import java.io.File;

@Component
public class MergeCommand implements CommandLineRunner {
    @Override
    public void run(String... args) {
        if (args.length < 2) {
            System.out.println("stl-editor input.stl output.plc vars.txt");
            return;
        }

        File inputDir = new File(args[0]);
        File summaryFile = new File(args[1]);

        new ExcelMerger(inputDir.listFiles(), summaryFile)
                .merge();
    }
}
