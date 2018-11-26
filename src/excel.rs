extern crate calamine;

use self::calamine::{Reader, open_workbook, Xlsx};

pub fn openbook() {
    println!("openbook called");
    let mut ws: Xlsx<_> = open_workbook("sampledata/sample.xlsx").expect("Cannot open file");

    if let Some(Ok(r)) = ws.worksheet_range("Tabelle1") {
        println!("rowcount {}",r.rows().count());  
        for row in r.rows() {
            println!("row={:?}, row[0]={:?}", row, row[0]);
            for col in row {
                println!("hmmm{}",col);
            }
        }
    }
}
