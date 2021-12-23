use std::path::PathBuf;

use anyhow::{anyhow, Result};
use clap::Parser;
use chrono::{Local, NaiveDate};
use calamine::{Xlsx, Reader, open_workbook};

#[derive(Debug, Parser)]
#[clap(version, about)]
struct Options {
    /// Print allocation for the given date.
    /// Defaults to today. Format: YYYY-MM-DD
    #[clap(short, long)]
    date: Option<NaiveDate>,

    /// The spreadsheet to parse.
    /// The format should be *.xlsx.
    file: PathBuf,
}

fn main() -> Result<()> {
    let opts = Options::parse();

    #[cfg(debug_assertions)]
    dbg!(&opts);

    let date =
        opts.date
        .unwrap_or_else(|| Local::today().naive_local());

    #[cfg(debug_assertions)]
    dbg!(date);

    let mut excel: Xlsx<_> = open_workbook(opts.file)?;

    let sheet_name = "zebra";

    let sheet =
        excel.worksheet_range(sheet_name)
        .ok_or_else(|| anyhow!("worksheet '{}' does not exist", sheet_name))??;

    for row in sheet.rows() {
        println!("row={:?}", row);
    }

    Ok(())
}
