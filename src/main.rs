use calamine::{open_workbook, Xlsx, Reader};

use std::sync::mpsc;
use std::fs::File;
use std::thread;

extern crate csv;

use csv::Writer;
struct EmployeeRow {
    pub last_name: String,
    pub first_name: String,
    pub preferred_name: String,
    pub job_family_group: String,
    pub email: String,
}

fn main() {
    let path = format!("{}/example-list.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook: Xlsx<_> = open_workbook(path).unwrap();

    let (tx, rx) = mpsc::channel();

    thread::spawn(move || {
        if let Some(Ok(r)) = workbook.worksheet_range("Sheet1") {
            let mut i = 0;

            for row in r.rows() {
                if i >= 0 && i <= 2 {
                    i += 1;
                    continue;
                }

                if row[12].to_string() == "" { continue }

                tx.send(EmployeeRow{
                    last_name: row[2].to_string(),
                    first_name: row[3].to_string(),
                    preferred_name: row[5].to_string(),
                    job_family_group: row[8].to_string(),
                    email: row[12].to_string()
                }).unwrap();
            }
        }
    });

    let valid_types = vec![
        "Faculty",
        "Administrative & Professional",
        "Executive Service",
        "UCF Athletic Association",
        "USPS"
    ];

    let mut with_writer = Writer::from_path("./with.csv").unwrap();
    let mut without_writer = Writer::from_path("./without.csv").unwrap();

    write_headers(&mut with_writer);
    write_headers(&mut without_writer);

    for received in rx {
        if valid_types.iter().any(|i| *i == received.job_family_group) {
            write_record(&mut without_writer, &received);
        }

        write_record(&mut with_writer, &received);
    }

    with_writer.flush().unwrap();
    without_writer.flush().unwrap();

    println!("All done!");

}

fn write_headers(writer: &mut Writer<File>) {
    let headers = vec![
        "first_name",
        "last_name",
        "email",
        "preferred_name"
    ];

    writer.write_record(&headers).unwrap();
}

fn write_record(writer: &mut Writer<File>, record: &EmployeeRow) {
    writer.write_record(&[
        record.first_name.clone(),
        record.last_name.clone(),
        record.email.clone(),
        record.preferred_name.clone()
    ]).unwrap();
}
