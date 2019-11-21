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

    let app =  app.arg(
        clap::Arg::with_name("sheet")
        .short("s")
        .long("sheet")
        .takes_value(true)
        .default_value("1")
        .help("Sheet number")
    );

    let matches = app.get_matches();

    let source = matches.value_of("source").expect("Source is not set");
    let sheet = matches.value_of("sheet").expect("Sheet is not set");

    match from_xlsx::from_xlsx(source, sheet) {
        Ok(csv) => {csv},
        Err(err) => {
            println!("{:?}", err);
            std::process::exit(1);
        }
    };
}
