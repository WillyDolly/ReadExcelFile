package com.popland.pop.readexcelfile;

import android.content.res.AssetManager;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.TextView;

import java.io.IOException;
import java.io.InputStream;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.read.biff.PasswordException;

public class MainActivity extends AppCompatActivity{
EditText edt;
TextView tv;
Sheet sheet;
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        tv = (TextView)findViewById(R.id.tv);
        edt = (EditText)findViewById(R.id.edt);

        try {
            AssetManager am = getAssets();
            InputStream is = am.open("wordlist.xls");//Jexcel(lightweigh) only take xls(97-2003) / other library Apache POI
            Workbook workBook = Workbook.getWorkbook(is);
            sheet = workBook.getSheet(0);
            //tv.setText(sheet.getColumns()+"-"+sheet.getRows());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void find_error(View v){
        //test 2 different finding algorithm with 100 words
        //find_alg2();// average 2.40s
        find_alg1();// average 0.57s
    }

    public void find_alg1(){
        String text = edt.getText().toString().toLowerCase();
        String[] wordList = text.split(" ");
        for(int i=0;i<wordList.length;i++){
            int column=0 , lastRow;
            //check first letter with 26 columns'first word -> find column number
            String c1 = wordList[i].substring(0,1);
            for(int j=0;j<sheet.getColumns();j++){
                String c2 = sheet.getCell(j,0).getContents().substring(0,1);
                if(c1.equals(c2)){
                    column = j;
                    break;
                }
            }
            //compare given word with all the column's cell -> compared cells =  26 + k.1000
            Cell[] cells = sheet.getColumn(column);
            lastRow = cells.length;
            Cell cell = sheet.findCell(wordList[i],column,0,column,lastRow,false);//check 50 words within 1s
            if(cell == null)
                tv.append(wordList[i]+"/");
        }
    }

    public void find_alg2(){
        //compare given word with 58 000 cells
        String text = edt.getText().toString().toLowerCase();
        String[] wordList = text.split(" ");
        for(int i=0;i<wordList.length;i++){
            Cell cell = sheet.findCell(wordList[i]);
            if(cell == null)
                tv.append(wordList[i]+"/");
        }
    }
}
