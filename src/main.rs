use umya_spreadsheet::{self, Worksheet};
use std::{fs::{self}, io::{self}, path::Path};
use chrono::{Datelike, Local, NaiveDate};


fn main() {
    //setting
    let new_name:&str=&Local::now().format("%Y%m").to_string();
    let target_column='H';


    //work
    let input_path="excel_input/";
    let output_path="excel_output/";
    let bak_path="excel_bak/";

    //test
    // let input_path="test/";
    // let output_path="test/";

    //get_excel
    if let Ok(excels)=get_excels(input_path,bak_path){
        println!("excel文件输入目录:{}",input_path);
        println!("----------------------------------------------------------");
        println!("excel文件输出目录:{}",output_path);
        println!("----------------------------------------------------------");
        println!("excel文件备份目录:{}",bak_path);
        println!("----------------------------------------------------------");
        for (index,path) in excels.iter().enumerate(){
            println!("文件 {} :  {} ",index+1,path);
            //整合path
            let input_path=Path::new(input_path).join(path);
            let output_path=Path::new(output_path).join(path);
            //read excel
            let mut book = umya_spreadsheet::reader::xlsx::read(&input_path).unwrap();
            //get sheet len
            let len=book.get_sheet_count();
            println!("      共有{} 张sheet",len);

            let last_sheet=book.get_sheet(&(len-1)).unwrap();
            //get last sheet clone
            let mut clone_sheet = last_sheet.clone();
            //set name
            clone_sheet.set_name(new_name);
            // println!("{:?}",clone_sheet.get_cell_mut("H27").get_style().get_number_format());
            // println!("{:?}",clone_sheet.get_cell_mut("H24").get_style().get_number_format());
            // println!("{:?}",clone_sheet.get_cell_mut("H25").get_style().get_number_format());

            //获取H列日期对应行号
            let taget_row=get_row_date_by_column(&clone_sheet, target_column);

            //修改日期为当月最后一天
            clone_sheet.get_cell_mut(format!("{}{}",target_column,taget_row))
            .set_value(get_last_day_of_current_month().to_string());
            
            //拼接公式  '{sheetname}'!{column}{row}
            /*
            账面已计提
            应提
            账面余额
            空行
            日期
             */
            let new_formlua=format!("'{}'!{}{}",last_sheet.get_name(),target_column,taget_row-2);
            //修改账面已计提为last_sheet的账面余额
            clone_sheet.get_cell_mut(format!("{}{}",target_column,taget_row-4)).set_formula(new_formlua);
            
            //add sheet
            let _ = book.add_sheet(clone_sheet);
            //write excel
            let _ = umya_spreadsheet::writer::xlsx::write(&book, &output_path);
            println!("      复制完成");
            //check
            book = umya_spreadsheet::reader::xlsx::read(&output_path).unwrap();
            if book.get_sheet_count()==len+1{
                println!("      写入成功");
                println!("      共有{} 张sheet",book.get_sheet_count());
            }else {
                println!("      写入失败");
                break;
            }
        }
    }
}
//get excels from dir
fn get_excels(path:&str,bak_path:&str)->io::Result<Vec<String>>{
    let mut excels=Vec::new();

    //read dir
    let entries=fs::read_dir(path);
    //遍历文件
    if let Ok(read_dir) =entries{
        for entry in read_dir{
            let entry = entry?;
            let excel_path = entry.path();
            if excel_path.is_file(){
                if let Some(excel_name) = excel_path.file_name() {
                    if let Some(name_str) = excel_name.to_str(){
                        //push into vec
                        excels.push(name_str.to_string());
                        //back up
                        let bak_path = Path::new(bak_path).join(name_str);
                        fs::copy(&excel_path, &bak_path)?;
                        println!("{:?}备份成功",bak_path);
                    }
                }
            }
        }
    }
    Ok(excels)
}

//日期序列化
fn date_to_excel_serial(date: NaiveDate) -> i32 {
    let base_date = NaiveDate::from_ymd_opt(1900, 1, 1).unwrap();
    let duration = date.signed_duration_since(base_date);
    (duration.num_days() + 2) as i32  // Excel从1900年1月1日开始，且它错误地将1900年当作闰年，因此需要+2
}

//获取当月最后一日 日期
fn get_last_day_of_current_month() -> i32 {
    let today = Local::now().naive_local();
    let (year, month) = (today.year(), today.month());

    // 获取当前月的最后一天
    let last_day_of_month = NaiveDate::from_ymd_opt(year, month, 1).unwrap()
        .with_month(month % 12 + 1)
        .map(|d| d.pred_opt())
        .unwrap().unwrap();

    // 将最后一天的日期转换为Excel序列号
    date_to_excel_serial(last_day_of_month)
}

//获取某列的最后有值行号，且属性为日期
fn get_row_date_by_column(sheet:&Worksheet,column:char)->u32{
    let mut target_row=0;
    if let Some(column_num) = letter_to_number(column) {
        for (row_num,cell) in sheet.get_collection_by_column_to_hashmap(&column_num){
            if !cell.get_value().trim().is_empty(){
                target_row = if target_row <= row_num { row_num } else { target_row };
            }
        }

        //check target cell is date?
        let cell_type=sheet.get_cell(format!("{}{}",column,target_row)).unwrap()
        .get_style().get_number_format().unwrap().get_number_format_id();
        if *cell_type==14u32{
            target_row
        }else {
            panic!("wrong column")
        }
    }else {
        panic!("error")
    }
}

//字母转数字
fn letter_to_number(c: char) -> Option<u32> {
    if c.is_ascii_alphabetic() {
        // 处理大小写字母都统一转换为小写来映射
        Some(c.to_ascii_lowercase() as u32 - 'a' as u32 + 1)
    } else {
        None
    }
}