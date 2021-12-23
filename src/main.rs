use std::path::PathBuf;

use anyhow::Result;
use clap::Parser;
use chrono::NaiveDate;

#[derive(Debug, Parser)]
#[clap(version, about)]
struct Options {
    /// Print allocation for the given date.
    #[clap(short, long)]
    date: Option<NaiveDate>,

    /// The spreadsheet to parse.
    /// The format should be *.xlsx.
    file: PathBuf,
}

fn main() -> Result<()> {
    let opts = Options::parse();

    #[cfg(debug_assertions)]
    eprintln!("Options: {:?}", &opts);

    Ok(())
}
