use quick_xml::events::Event;
use quick_xml::Reader;

use zip;

use std::fs;
use std::io;
use std::io::prelude::*;

#[derive(Debug)]
pub enum ConvertError {
    Io(io::Error),
    Zip(zip::result::ZipError),
    Xml(quick_xml::Error),
    Csv(csv::Error),
}

pub fn from_xlsx(file_name: String) -> Result<String, ConvertError> {
    let file = fs::File::open(&file_name).map_err(ConvertError::Io)?;
    let mut archive = zip::ZipArchive::new(file).map_err(ConvertError::Zip)?;
    let mut file = archive
        .by_name("xl/sharedStrings.xml")
        .map_err(ConvertError::Zip)?;
    let mut contents = String::new();
    file.read_to_string(&mut contents)
        .map_err(ConvertError::Io)?;
    let mut reader = Reader::from_str(&contents);
    let mut shared_strings: Vec<String> = Vec::new();
    let mut buf = Vec::new();
    let mut is_t = false;
    loop {
        match reader.read_event(&mut buf) {
            Ok(Event::Start(ref e)) => is_t = b"t" == e.name(),
            Ok(Event::Text(e)) => {
                if is_t {
                    shared_strings.push(e.unescape_and_decode(&reader).unwrap());
                    is_t = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(err) => return Err(err).map_err(ConvertError::Xml),
            _ => (),
        }
        buf.clear();
    }
    std::mem::drop(file);

    let mut file = archive
        .by_name("xl/worksheets/sheet1.xml")
        .map_err(ConvertError::Zip)?;

    let mut contents = String::new();
    file.read_to_string(&mut contents)
        .map_err(ConvertError::Io)?;
    let mut reader = Reader::from_str(&contents);
    let mut wtr = csv::Writer::from_writer(io::stdout());
    let mut rows_count = 0;
    let mut is_v = false;
    let mut buf = Vec::new();
    let mut value = String::new();
    let mut value_t = false;
    let mut row: Vec<String> = Vec::new();

    loop {
        match reader.read_event(&mut buf) {
            Ok(Event::Text(e)) => {
                if is_v {
                    value = e.unescape_and_decode(&reader).unwrap();
                    is_v = false;
                }
            }
            Ok(Event::Start(ref e)) => {
                match e.local_name() {
                    b"v" => is_v = true,
                    b"c" => {
                        let t = e
                            .attributes()
                            .map(|ar| ar.expect("Expecting attribute parsing to succeed."))
                            .filter(|kv| kv.key.starts_with(b"t"))
                            .next();
                        match t {
                            Some(_) => value_t = true,
                            None => value_t = false,
                        };
                    }
                    _ => (),
                };
            }
            Ok(Event::End(ref e)) => match e.name() {
                b"c" => {
                    if !value_t {
                        let v: String = value.parse().unwrap();
                        row.push(v);
                    } else {
                        let index: usize = value.parse().unwrap();
                        let v = shared_strings.get(index).unwrap().clone();
                        row.push(v);
                    }
                    value.clear();
                }
                b"row" => {
                    rows_count = rows_count + 1;
                    wtr.write_record(&row).map_err(ConvertError::Csv)?;
                    wtr.flush().map_err(ConvertError::Io)?;
                    row.clear();
                }
                _ => (),
            },
            Ok(Event::Empty(ref e)) => match e.name() {
                b"c" => row.push(String::new()),
                _ => (),
            },
            Ok(Event::Eof) => break,
            Err(err) => return Err(err).map_err(ConvertError::Xml),
            _ => (),
        };
        buf.clear();
    }
    Ok(String::new())
}
