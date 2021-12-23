use std::path::PathBuf;

use anyhow::{anyhow, Result};
use clap::Parser;
use chrono::{Local, NaiveDate};
use calamine::{Xlsx, Reader, open_workbook, DataType};

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

    let mut rows = sheet.rows();

    let header = rows.next()
        .ok_or_else(|| anyhow!("no rows"))?;

    let column_day = 7;
    let column_month = 8;
    let column_year = 9;
    expect_header(header, column_day, "Day")?;
    expect_header(header, column_month, "Month")?;
    expect_header(header, column_year, "Year")?;

    for row in rows {
        let year = get_number(row, column_year)?.try_into()?;
        let month = get_number(row, column_month)?.try_into()?;
        let day = get_number(row, column_day)?.try_into()?;
        let row_date = NaiveDate::from_ymd(year, month, day);
        println!("row={:?}, date={:?}", row, row_date);
    }

    Ok(())
}

fn get_string(row: &[DataType], col: usize) -> Result<&str> {
    row
        .get(col)
        .ok_or_else(|| anyhow!("no column of index {}", col))?
        .get_string()
        .ok_or_else(|| anyhow!("expected String of column {}", col))
}

fn expect_header(header: &[DataType], col: usize, text: &str) -> Result<()> {
    if get_string(header, col)? != text {
        Err(anyhow!("Expected header '{}' in column '{}'", text, col))
    } else {
        Ok(())
    }
}

fn get_number(row: &[DataType], col: usize) -> Result<i64> {
    let value =
        row.get(col)
        .ok_or_else(|| anyhow!("no column of index {}", col))?;
    match value {
        &DataType::Int(x) => Ok(x),
        &DataType::Float(x) => Ok(x as i64),
        x => Err(anyhow!("Expected number in column '{}', but got '{}'", col, x)),
    }
}
