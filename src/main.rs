mod from_xlsx;

use clap;

use std;

fn main() {
    let app = clap::App::new("xlsx2csv")
    .version("0.1")
    .author("Ruslan Nasonov <rus.nasonov@gmail.com>")
    .about("Convert xlsx to csv ");
    
    let app = app.arg(
        clap::Arg::with_name("source")
        .help("Path to file or '-' for stdin")
        .required(true)
        .index(1)
    );

    let matches = app.get_matches();

    let source = match matches.value_of("source"){
        Some(source) => {source},
        None => {
            println!("Path to xslx does't set");
            std::process::exit(1);
        }
    };
    
    let csv = match from_xlsx::from_xlsx(source.to_string()) {
        Ok(csv) => {csv},
        Err(err) => {
            println!("{:?}", err);
            std::process::exit(1);
        }
    };
    println!("{}", csv);
}
