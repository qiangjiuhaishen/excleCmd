#!/usr/bin/env node
"use strict";
const path = require("path");
const program=require("commander")

program
        .option('-r,--replace[type]','new replace old')
        .action(()=>{
          let filePath=path.resolve(__dirname,"../")
          const XLSX = require('xlsx')  ;
          const workbook = XLSX.readFile("./陈集.xlsx");
          // const workbook = XLSX.readFile(filePath+"\\"+`${excelName.replace}`+".xlsx");
          // 获取sheet1
          const sheetFirst= workbook.SheetNames[0]; // 获取工作簿中的工作表名字
          const worksheet1 = workbook.Sheets[sheetFirst]; // 获取对应的工作表对象
          const sheet1 = XLSX.utils.sheet_to_json(worksheet1)
          
          // 获取sheet2
          const sheetSecond=workbook.SheetNames[1]
          const worksheet2=workbook.Sheets[sheetSecond]
          const sheet2=XLSX.utils.sheet_to_json(worksheet2)
          let reg=/[0-9]{12,}[A-Z]{2}[0-9]{5}\s[\u4e00-\u9fa5]{2,}/g
          // 遍历sheet2
          const sheet2New=sheet2.map((item,index)=>{
            
            // 拷贝到新数组
             const itemCopy=Object.assign({},item);
             ["ZDSZB", "ZDSZD", "ZDSZN", "ZDSZX"].forEach((key)=>{
          
               let str=item[key]
               let group=str.match(reg)
               if(group!==null){
                 group.forEach((match)=>{
                  const findItem =sheet1.find((it)=> it["原"]==match)
                  if(findItem!=undefined){
                    
                    itemCopy[key]=itemCopy[key].replace(match,findItem ["现"])
                   
                  }
                 })
               }  
             })
             return itemCopy 
          })
          
          // 保存
          
          workbook.Sheets["Sheet2"] = XLSX.utils.json_to_sheet(sheet2New);
          XLSX.writeFile(workbook, "out.xlsx");
        })
        .parse(process.argv);
     if(process.replace)console.log("ok");
     
