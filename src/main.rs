use umya_spreadsheet;
use std::{fs::{self}, io::{self}, path::Path};


fn main() {
    //work
    let input_path="excel_input/";
    let output_path="excel_output/";
    let bak_path="excel_bak/";
    let new_name="202502";

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
            //get last sheet clone
            let mut clone_sheet = book.get_sheet(&(len-1)).unwrap().clone();
            //set name
            // clone_sheet.set_name(format!("Sheet{}", len+1));
            clone_sheet.set_name(new_name);
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