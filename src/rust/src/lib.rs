use chrono::{NaiveDate, NaiveDateTime, Timelike};
use extendr_api::prelude::*;
use fastexcel_rs::{
    read_excel, ExcelReader, FastExcelSeries, IdxOrName, LoadSheetOrTableOptions, SelectedColumns,
};
use std::str::FromStr;

#[extendr]
fn read_excel_columns(
    source: Robj,
    sheet: Robj,
    range: Robj,
    col_names: Robj,
    n_max: Robj,
) -> Result<List> {
    let mut reader = reader_from_source(source)?;
    let mut opts = LoadSheetOrTableOptions::new_for_sheet();

    if col_names.as_bool() == Some(false) {
        opts = opts.no_header_row();
    } else if let Some(names) = col_names.as_str_vector() {
        opts = opts.no_header_row().column_names(names);
    }

    if let Some(n) = n_max.as_integer() {
        if n >= 0 {
            opts = opts.n_rows(n as usize);
        }
    }

    if let Some(selection) = range.as_str() {
        if !selection.is_empty() && selection != "NA" {
            opts = opts.selected_columns(selected_columns_from_range(selection)?);
        }
    }

    let idx_or_name = sheet_to_idx_or_name(sheet)?;
    let sheet = reader.load_sheet(idx_or_name, opts).map_err(to_r_error)?;
    let columns = sheet.to_columns().map_err(to_r_error)?;

    let mut pairs: Vec<(&str, Robj)> = Vec::with_capacity(columns.len());
    let mut names: Vec<String> = Vec::with_capacity(columns.len());
    let mut values: Vec<Robj> = Vec::with_capacity(columns.len());

    for column in columns {
        names.push(column.name().to_string());
        values.push(series_to_robj(column.data(), column.len())?);
    }

    for (name, value) in names.iter().zip(values.into_iter()) {
        pairs.push((name.as_str(), value));
    }

    Ok(List::from_pairs(pairs))
}

#[extendr]
fn excel_sheets(source: Robj) -> Result<Vec<String>> {
    let reader = reader_from_source(source)?;
    Ok(reader.sheet_names().into_iter().map(String::from).collect())
}

#[extendr]
fn excel_tables(source: Robj, sheet: Robj) -> Result<Vec<String>> {
    let mut reader = reader_from_source(source)?;
    let sheet_name = sheet.as_str().filter(|s| !s.is_empty() && *s != "NA");
    Ok(reader
        .table_names(sheet_name)
        .map_err(to_r_error)?
        .into_iter()
        .map(String::from)
        .collect())
}

#[extendr]
fn excel_defined_names(source: Robj) -> Result<List> {
    let mut reader = reader_from_source(source)?;
    let names = reader.defined_names().map_err(to_r_error)?;

    let name_values: Vec<String> = names.iter().map(|item| item.name.to_string()).collect();
    let formula_values: Vec<String> = names.iter().map(|item| item.formula.to_string()).collect();
    let sheet_values: Vec<String> = std::iter::repeat(String::new()).take(names.len()).collect();

    Ok(list!(
        name = name_values,
        formula = formula_values,
        sheet_name = sheet_values
    ))
}

fn reader_from_source(source: Robj) -> Result<ExcelReader> {
    if let Some(path) = source.as_str() {
        return read_excel(path).map_err(to_r_error);
    }

    if let Some(bytes) = source.as_raw_slice() {
        return ExcelReader::try_from(bytes).map_err(|err| {
            Error::Other(format!("could not load excel file from raw bytes: {err}"))
        });
    }

    Err(Error::Other(
        "`path` must be a single non-empty string or a raw vector.".to_string(),
    ))
}

fn sheet_to_idx_or_name(sheet: Robj) -> Result<IdxOrName> {
    if let Some(name) = sheet.as_str() {
        return Ok(IdxOrName::Name(name.to_string()));
    }

    let idx = sheet
        .as_integer()
        .ok_or_else(|| Error::Other("`sheet` must be a string or integer".to_string()))?;

    if idx < 1 {
        return Err(Error::Other("`sheet` index must be >= 1".to_string()));
    }

    Ok(IdxOrName::Idx((idx - 1) as usize))
}

fn selected_columns_from_range(selection: &str) -> Result<SelectedColumns> {
    if let Some((start, end)) = selection.split_once(':') {
        if let (Some(start_idx), Some(end_idx)) =
            (column_label_to_idx(start), column_label_to_idx(end))
        {
            let (first, last) = if start_idx <= end_idx {
                (start_idx, end_idx)
            } else {
                (end_idx, start_idx)
            };
            return Ok(SelectedColumns::Selection(
                (first..=last).map(IdxOrName::Idx).collect(),
            ));
        }
    }

    SelectedColumns::from_str(selection)
        .map_err(|err| Error::Other(format!("invalid range `{selection}`: {err}")))
}

fn column_label_to_idx(label: &str) -> Option<usize> {
    let label = label.trim();
    if label.is_empty() || !label.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return None;
    }

    let mut idx = 0usize;
    for byte in label.bytes() {
        idx = idx * 26 + usize::from(byte.to_ascii_uppercase() - b'A' + 1);
    }
    Some(idx - 1)
}

fn series_to_robj(series: &FastExcelSeries, len: usize) -> Result<Robj> {
    Ok(match series {
        FastExcelSeries::Null => null_logical_vector(len),
        FastExcelSeries::Bool(values) => logical_vector(values),
        FastExcelSeries::String(values) => string_vector(values),
        FastExcelSeries::Int(values) => integer_or_numeric_vector(values),
        FastExcelSeries::Float(values) => numeric_vector(values),
        FastExcelSeries::Datetime(values) => datetime_vector(values)?,
        FastExcelSeries::Date(values) => date_vector(values)?,
        FastExcelSeries::Duration(values) => duration_vector(values),
    })
}

fn null_logical_vector(len: usize) -> Robj {
    let values: Vec<Rbool> = std::iter::repeat(Rbool::na()).take(len).collect();
    Robj::from(values)
}

fn logical_vector(values: &[Option<bool>]) -> Robj {
    Robj::from(
        values
            .iter()
            .map(|value| value.map(Rbool::from).unwrap_or_else(Rbool::na))
            .collect::<Vec<_>>(),
    )
}

fn string_vector(values: &[Option<String>]) -> Robj {
    Robj::from(
        values
            .iter()
            .map(|value| value.as_deref())
            .collect::<Vec<_>>(),
    )
}

fn integer_or_numeric_vector(values: &[Option<i64>]) -> Robj {
    let fits_i32 = values
        .iter()
        .flatten()
        .all(|value| i32::try_from(*value).is_ok());

    if fits_i32 {
        Robj::from(
            values
                .iter()
                .map(|value| value.map(|value| value as i32).or(NA_INTEGER))
                .collect::<Vec<_>>(),
        )
    } else {
        Robj::from(
            values
                .iter()
                .map(|value| value.map(|value| value as f64).or(NA_REAL))
                .collect::<Vec<_>>(),
        )
    }
}

fn numeric_vector(values: &[Option<f64>]) -> Robj {
    Robj::from(values.iter().copied().collect::<Vec<_>>())
}

fn date_vector(values: &[Option<NaiveDate>]) -> Result<Robj> {
    let epoch = NaiveDate::from_ymd_opt(1970, 1, 1)
        .ok_or_else(|| Error::Other("failed to construct Unix epoch date".to_string()))?;
    let mut out = Robj::from(
        values
            .iter()
            .map(|value| {
                value
                    .map(|value| (value - epoch).num_days() as f64)
                    .or(NA_REAL)
            })
            .collect::<Vec<_>>(),
    );
    out.set_class(&["Date"])?;
    Ok(out)
}

fn datetime_vector(values: &[Option<NaiveDateTime>]) -> Result<Robj> {
    let mut out = Robj::from(
        values
            .iter()
            .map(|value| {
                value
                    .map(|value| {
                        value.and_utc().timestamp() as f64 + f64::from(value.nanosecond()) / 1e9
                    })
                    .or(NA_REAL)
            })
            .collect::<Vec<_>>(),
    );
    out.set_class(&["POSIXct", "POSIXt"])?;
    out.set_attrib("tzone", "UTC")?;
    Ok(out)
}

fn duration_vector(values: &[Option<chrono::Duration>]) -> Robj {
    Robj::from(
        values
            .iter()
            .map(|value| {
                value
                    .map(|value| value.num_milliseconds() as f64 / 1000.0)
                    .or(NA_REAL)
            })
            .collect::<Vec<_>>(),
    )
}

fn to_r_error<E: std::fmt::Display>(err: E) -> Error {
    Error::Other(err.to_string())
}

extendr_module! {
    mod fastexcel;
    fn read_excel_columns;
    fn excel_sheets;
    fn excel_tables;
    fn excel_defined_names;
}
