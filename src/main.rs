use std::path::PathBuf;

use anyhow::{anyhow, Result};
use clap::Parser;
use chrono::{Local, NaiveDate, Datelike};
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

#[derive(Debug, Clone)]
struct Node {
    name: String,
    col: usize,
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
        .ok_or_else(|| anyhow!("worksheet '{sheet_name}' does not exist"))??;

    let mut rows = sheet.rows();

    let header = rows.next()
        .ok_or_else(|| anyhow!("no rows"))?;

    let col_date = 0;
    expect_header(header, col_date, "Date")?;

    let mut selected_row = None;

    for row in rows {
        let row_date = get_as_date(row, col_date)?;
        if row_date == date {
            if selected_row.is_some() {
                return Err(anyhow!("Multiple rows match the requested date '{date}':\n{0:?}\n{row:?}", selected_row.unwrap()));
            }
            selected_row = Some(row);
        }
    }

    let selected_row = selected_row
        .ok_or_else(|| anyhow!("No row matches the requested date '{date}'"))?;

    #[cfg(debug_assertions)]
    dbg!(selected_row);

    let nodes = [
        Node { name: "zebra-node01".into(), col: 4 },
        Node { name: "zebra-node02".into(), col: 5 },
        Node { name: "zebra-node03".into(), col: 6 },
    ];

    println!("Node allocation for {} ({}):", date, date.weekday());
    for node in &nodes {
        expect_header(header, node.col, &node.name)?;
        let user = get_as_string(selected_row, node.col)?;
        let user_display = if user.is_empty() { "<unassigned>" } else { &user };
        println!("    {}: {}", node.name, user_display);
    }

    Ok(())
}

fn get_string(row: &[DataType], col: usize) -> Result<&str> {
    row
        .get(col)
        .ok_or_else(|| anyhow!("no column of index {col}"))?
        .get_string()
        .ok_or_else(|| anyhow!("expected String in column {col}"))
}

fn expect_header(header: &[DataType], col: usize, text: &str) -> Result<()> {
    if get_string(header, col)? != text {
        Err(anyhow!("Expected header '{text}' in column '{col}'"))
    } else {
        Ok(())
    }
}

#[allow(dead_code)]
fn get_number(row: &[DataType], col: usize) -> Result<i64> {
    let value =
        row.get(col)
        .ok_or_else(|| anyhow!("no column of index {col}"))?;
    match value {
        &DataType::Int(x) => Ok(x),
        &DataType::Float(x) => Ok(x as i64),
        x => Err(anyhow!("expected number in column {col}, but got '{x}'")),
    }
}

fn get_as_date(row: &[DataType], col: usize) -> Result<NaiveDate> {
    row.get(col)
        .ok_or_else(|| anyhow!("no column of index {col}"))?
        .as_date()
        .ok_or_else(|| anyhow!("expected date in column {col}"))
}

fn get_as_string(row: &[DataType], col: usize) -> Result<String> {
    let value =
        row.get(col)
        .ok_or_else(|| anyhow!("no column of index {col}"))?;
    match value {
        &DataType::Bool(x) => Ok(format!("{x}")),
        &DataType::Int(x) => Ok(format!("{x}")),
        &DataType::Float(x) => Ok(format!("{x}")),
        &DataType::DateTime(x) => Ok(format!("{x}")),
        &DataType::String(ref x) => Ok(x.clone()),
        &DataType::Empty => Ok(String::new()),
        &DataType::Error(ref x) => Ok(format!("#ERROR: {x}")),
    }
}
