use arrow_array::ffi::{FFI_ArrowArray, FFI_ArrowSchema};
use arrow_array::{
    ArrayRef, BooleanArray, Date32Array, DurationMillisecondArray, Float64Array, Int64Array,
    NullArray, StringArray, StructArray, TimestampNanosecondArray,
};
use arrow_schema::{DataType, Field, TimeUnit};
use chrono::{NaiveDate, NaiveDateTime, Timelike};
use extendr_api::prelude::*;
use extendr_api::R_ExternalPtrAddr;
use fastexcel_rs::{
    read_excel, ColumnInfo, DType, DTypeCoercion, DTypes, ExcelReader, FastExcelColumn,
    FastExcelSeries, IdxOrName, LoadSheetOrTableOptions, SelectedColumns, SheetVisible, SkipRows,
};
use std::collections::HashMap;
use std::fs::File;
use std::io::{Cursor, Read, Seek};
use std::str::FromStr;
use std::sync::Arc;
use zip::result::ZipError;
use zip::ZipArchive;

const ZIP_MAGIC: &[u8; 4] = b"PK\x03\x04";
const EMPTY_ZIP_MAGIC: &[u8; 4] = b"PK\x05\x06";
const SPANNED_ZIP_MAGIC: &[u8; 4] = b"PK\x07\x08";

struct ZipLimits {
    max_entries: usize,
    max_entry_size: u64,
    max_total_size: u64,
    max_compression_ratio: u64,
}

#[extendr]
#[allow(clippy::too_many_arguments)]
fn read_excel_columns(
    source: Robj,
    zip_limits: Robj,
    sheet: Robj,
    range: Robj,
    columns: Robj,
    col_names: Robj,
    header_row: Robj,
    skip_rows: Robj,
    n_max: Robj,
    schema_sample_rows: Robj,
    dtype_coercion: Robj,
    dtypes: Robj,
    skip_whitespace_tail_rows: bool,
    whitespace_as_null: bool,
) -> Result<List> {
    let columns = load_columns(
        source,
        zip_limits,
        sheet,
        range,
        columns,
        col_names,
        header_row,
        skip_rows,
        n_max,
        schema_sample_rows,
        dtype_coercion,
        dtypes,
        skip_whitespace_tail_rows,
        whitespace_as_null,
    )?;

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
#[allow(clippy::too_many_arguments)]
fn read_excel_arrow(
    source: Robj,
    zip_limits: Robj,
    sheet: Robj,
    range: Robj,
    columns: Robj,
    col_names: Robj,
    header_row: Robj,
    skip_rows: Robj,
    n_max: Robj,
    schema_sample_rows: Robj,
    dtype_coercion: Robj,
    dtypes: Robj,
    skip_whitespace_tail_rows: bool,
    whitespace_as_null: bool,
    array: Robj,
    schema: Robj,
    single_column: bool,
) -> Result<()> {
    let columns = load_columns(
        source,
        zip_limits,
        sheet,
        range,
        columns,
        col_names,
        header_row,
        skip_rows,
        n_max,
        schema_sample_rows,
        dtype_coercion,
        dtypes,
        skip_whitespace_tail_rows,
        whitespace_as_null,
    )?;

    export_columns(columns, array, schema, single_column)
}

#[extendr]
#[allow(clippy::too_many_arguments)]
fn read_excel_table_arrow(
    source: Robj,
    zip_limits: Robj,
    table: &str,
    columns: Robj,
    col_names: Robj,
    header_row: Robj,
    skip_rows: Robj,
    n_max: Robj,
    schema_sample_rows: Robj,
    dtype_coercion: Robj,
    dtypes: Robj,
    skip_whitespace_tail_rows: bool,
    whitespace_as_null: bool,
    array: Robj,
    schema: Robj,
    single_column: bool,
) -> Result<()> {
    let columns = load_table_columns(
        source,
        zip_limits,
        table,
        columns,
        col_names,
        header_row,
        skip_rows,
        n_max,
        schema_sample_rows,
        dtype_coercion,
        dtypes,
        skip_whitespace_tail_rows,
        whitespace_as_null,
    )?;

    export_columns(columns, array, schema, single_column)
}

#[allow(clippy::too_many_arguments)]
fn load_columns(
    source: Robj,
    zip_limits: Robj,
    sheet: Robj,
    range: Robj,
    columns: Robj,
    col_names: Robj,
    header_row: Robj,
    skip_rows: Robj,
    n_max: Robj,
    schema_sample_rows: Robj,
    dtype_coercion: Robj,
    dtypes: Robj,
    skip_whitespace_tail_rows: bool,
    whitespace_as_null: bool,
) -> Result<Vec<FastExcelColumn>> {
    let mut reader = reader_from_source(source, zip_limits)?;
    let opts = load_options(
        LoadSheetOrTableOptions::new_for_sheet(),
        range,
        columns,
        col_names,
        header_row,
        skip_rows,
        n_max,
        schema_sample_rows,
        dtype_coercion,
        dtypes,
        skip_whitespace_tail_rows,
        whitespace_as_null,
    )?;

    let idx_or_name = sheet_to_idx_or_name(sheet)?;
    let sheet = reader.load_sheet(idx_or_name, opts).map_err(to_r_error)?;
    sheet.to_columns().map_err(to_r_error)
}

#[allow(clippy::too_many_arguments)]
fn load_table_columns(
    source: Robj,
    zip_limits: Robj,
    table: &str,
    columns: Robj,
    col_names: Robj,
    header_row: Robj,
    skip_rows: Robj,
    n_max: Robj,
    schema_sample_rows: Robj,
    dtype_coercion: Robj,
    dtypes: Robj,
    skip_whitespace_tail_rows: bool,
    whitespace_as_null: bool,
) -> Result<Vec<FastExcelColumn>> {
    let mut reader = reader_from_source(source, zip_limits)?;
    let opts = load_options(
        LoadSheetOrTableOptions::new_for_table(),
        Robj::from(()),
        columns,
        col_names,
        header_row,
        skip_rows,
        n_max,
        schema_sample_rows,
        dtype_coercion,
        dtypes,
        skip_whitespace_tail_rows,
        whitespace_as_null,
    )?;

    let table = reader.load_table(table, opts).map_err(to_r_error)?;
    table.to_columns().map_err(to_r_error)
}

#[allow(clippy::too_many_arguments)]
fn load_options(
    mut opts: LoadSheetOrTableOptions,
    range: Robj,
    columns: Robj,
    col_names: Robj,
    header_row: Robj,
    skip_rows: Robj,
    n_max: Robj,
    schema_sample_rows: Robj,
    dtype_coercion: Robj,
    dtypes: Robj,
    skip_whitespace_tail_rows: bool,
    whitespace_as_null: bool,
) -> Result<LoadSheetOrTableOptions> {
    if col_names.as_bool() == Some(false) {
        opts = opts.no_header_row();
    } else if let Some(names) = col_names.as_str_vector() {
        opts = opts.no_header_row().column_names(names);
    } else if let Some(row) = optional_positive_integer(header_row)? {
        opts = opts.header_row(row - 1);
    }

    if let Some(n) = optional_non_negative_integer(skip_rows)? {
        opts = opts.skip_rows(SkipRows::Simple(n));
    }

    if let Some(n) = n_max.as_integer() {
        if n >= 0 {
            opts = opts.n_rows(n as usize);
        }
    }

    if let Some(n) = optional_positive_integer(schema_sample_rows)? {
        opts = opts.schema_sample_rows(n);
    }

    if let Some(value) = dtype_coercion.as_str() {
        opts = opts.dtype_coercion(DTypeCoercion::from_str(value).map_err(to_r_error)?);
    }

    if let Some(value) = dtypes_from_robj(&dtypes)? {
        opts = opts.with_dtypes(value);
    }

    opts = opts
        .skip_whitespace_tail_rows(skip_whitespace_tail_rows)
        .whitespace_as_null(whitespace_as_null);

    if let Some(selection) = range.as_str() {
        if !selection.is_empty() && selection != "NA" {
            opts = opts.selected_columns(selected_columns_from_range(selection)?);
        }
    }

    if let Some(selection) = selected_columns_from_robj(&columns)? {
        opts = opts.selected_columns(selection);
    }

    Ok(opts)
}

fn optional_positive_integer(value: Robj) -> Result<Option<usize>> {
    optional_integer(value, false)
}

fn optional_non_negative_integer(value: Robj) -> Result<Option<usize>> {
    optional_integer(value, true)
}

fn optional_integer(value: Robj, zero_allowed: bool) -> Result<Option<usize>> {
    match value.as_integer() {
        Some(value) if value == i32::MIN => Ok(None),
        Some(value) if value > 0 || (zero_allowed && value == 0) => Ok(Some(value as usize)),
        Some(_) => Err(Error::Other("invalid row count option".to_string())),
        None => Ok(None),
    }
}

fn dtypes_from_robj(dtypes: &Robj) -> Result<Option<DTypes>> {
    let Some(values) = dtypes.as_str_vector() else {
        return Ok(None);
    };

    if values.is_empty() || (values.len() == 1 && values[0] == "NA") {
        return Ok(None);
    }

    let names = dtypes.names().map(|names| names.collect::<Vec<_>>());
    if values.len() == 1 && names.is_none() {
        return Ok(Some(DTypes::All(
            DType::from_str(values[0]).map_err(to_r_error)?,
        )));
    }

    let names =
        names.ok_or_else(|| Error::Other("named `dtypes` require column names".to_string()))?;
    let mut map = HashMap::with_capacity(values.len());
    for (name, dtype) in names.into_iter().zip(values.into_iter()) {
        map.insert(
            IdxOrName::Name(name.to_string()),
            DType::from_str(dtype).map_err(to_r_error)?,
        );
    }
    Ok(Some(DTypes::Map(map)))
}

#[extendr]
fn excel_sheets(source: Robj, zip_limits: Robj) -> Result<Vec<String>> {
    let reader = reader_from_source(source, zip_limits)?;
    Ok(reader.sheet_names().into_iter().map(String::from).collect())
}

#[extendr]
fn excel_sheet_info(source: Robj, zip_limits: Robj, sheet: Robj) -> Result<List> {
    let mut reader = reader_from_source(source, zip_limits)?;
    let sheet_refs = sheet_refs(&reader, sheet)?;

    let mut names = Vec::with_capacity(sheet_refs.len());
    let mut widths = Vec::with_capacity(sheet_refs.len());
    let mut heights = Vec::with_capacity(sheet_refs.len());
    let mut total_heights = Vec::with_capacity(sheet_refs.len());
    let mut visibilities = Vec::with_capacity(sheet_refs.len());

    for sheet_ref in sheet_refs {
        let mut sheet = reader
            .load_sheet(sheet_ref, LoadSheetOrTableOptions::new_for_sheet())
            .map_err(to_r_error)?;
        names.push(sheet.name().to_string());
        widths.push(sheet.width() as i32);
        heights.push(sheet.height() as i32);
        total_heights.push(sheet.total_height() as i32);
        visibilities.push(sheet_visibility_to_str(sheet.visible()).to_string());
    }

    Ok(list!(
        name = names,
        width = widths,
        height = heights,
        total_height = total_heights,
        visibility = visibilities
    ))
}

#[extendr]
#[allow(clippy::too_many_arguments)]
fn excel_sheet_columns(
    source: Robj,
    zip_limits: Robj,
    sheet: Robj,
    range: Robj,
    columns: Robj,
    col_names: Robj,
    header_row: Robj,
    skip_rows: Robj,
    n_max: Robj,
    schema_sample_rows: Robj,
    dtype_coercion: Robj,
    dtypes: Robj,
    skip_whitespace_tail_rows: bool,
    whitespace_as_null: bool,
    available: bool,
) -> Result<List> {
    let mut reader = reader_from_source(source, zip_limits)?;
    let opts = load_options(
        LoadSheetOrTableOptions::new_for_sheet(),
        range,
        columns,
        col_names,
        header_row,
        skip_rows,
        n_max,
        schema_sample_rows,
        dtype_coercion,
        dtypes,
        skip_whitespace_tail_rows,
        whitespace_as_null,
    )?;

    let mut sheet = reader
        .load_sheet(sheet_to_idx_or_name(sheet)?, opts)
        .map_err(to_r_error)?;
    let columns = if available {
        sheet.available_columns().map_err(to_r_error)?
    } else {
        sheet.selected_columns().clone()
    };
    column_info_to_list(columns)
}

#[extendr]
fn excel_tables(source: Robj, zip_limits: Robj, sheet: Robj) -> Result<Vec<String>> {
    let mut reader = reader_from_source(source, zip_limits)?;
    let sheet_name = sheet.as_str().filter(|s| !s.is_empty() && *s != "NA");
    Ok(reader
        .table_names(sheet_name)
        .map_err(to_r_error)?
        .into_iter()
        .map(String::from)
        .collect())
}

#[extendr]
fn excel_table_info(source: Robj, zip_limits: Robj, table: Robj) -> Result<List> {
    let mut reader = reader_from_source(source, zip_limits)?;
    let table_names = if let Some(name) = table.as_str().filter(|s| !s.is_empty() && *s != "NA") {
        vec![name.to_string()]
    } else {
        reader
            .table_names(None)
            .map_err(to_r_error)?
            .into_iter()
            .map(String::from)
            .collect()
    };

    let mut names = Vec::with_capacity(table_names.len());
    let mut sheet_names = Vec::with_capacity(table_names.len());
    let mut widths = Vec::with_capacity(table_names.len());
    let mut heights = Vec::with_capacity(table_names.len());
    let mut total_heights = Vec::with_capacity(table_names.len());

    for name in table_names {
        let mut table = reader
            .load_table(&name, LoadSheetOrTableOptions::new_for_table())
            .map_err(to_r_error)?;
        names.push(table.name().to_string());
        sheet_names.push(table.sheet_name().to_string());
        widths.push(table.width() as i32);
        heights.push(table.height() as i32);
        total_heights.push(table.total_height() as i32);
    }

    Ok(list!(
        name = names,
        sheet_name = sheet_names,
        width = widths,
        height = heights,
        total_height = total_heights
    ))
}

#[extendr]
#[allow(clippy::too_many_arguments)]
fn excel_table_columns(
    source: Robj,
    zip_limits: Robj,
    table: &str,
    columns: Robj,
    col_names: Robj,
    header_row: Robj,
    skip_rows: Robj,
    n_max: Robj,
    schema_sample_rows: Robj,
    dtype_coercion: Robj,
    dtypes: Robj,
    skip_whitespace_tail_rows: bool,
    whitespace_as_null: bool,
    available: bool,
) -> Result<List> {
    let mut reader = reader_from_source(source, zip_limits)?;
    let opts = load_options(
        LoadSheetOrTableOptions::new_for_table(),
        Robj::from(()),
        columns,
        col_names,
        header_row,
        skip_rows,
        n_max,
        schema_sample_rows,
        dtype_coercion,
        dtypes,
        skip_whitespace_tail_rows,
        whitespace_as_null,
    )?;

    let mut table = reader.load_table(table, opts).map_err(to_r_error)?;
    let columns = if available {
        table.available_columns().map_err(to_r_error)?
    } else {
        table.selected_columns()
    };
    column_info_to_list(columns)
}

fn sheet_refs(reader: &ExcelReader, sheet: Robj) -> Result<Vec<IdxOrName>> {
    if sheet.is_null() {
        return Ok(reader
            .sheet_names()
            .into_iter()
            .map(|name| IdxOrName::Name(name.to_string()))
            .collect());
    }

    Ok(vec![sheet_to_idx_or_name(sheet)?])
}

fn sheet_visibility_to_str(visibility: SheetVisible) -> &'static str {
    match visibility {
        SheetVisible::Visible => "visible",
        SheetVisible::Hidden => "hidden",
        SheetVisible::VeryHidden => "very_hidden",
    }
}

fn column_info_to_list(columns: Vec<ColumnInfo>) -> Result<List> {
    let mut names = Vec::with_capacity(columns.len());
    let mut indices = Vec::with_capacity(columns.len());
    let mut absolute_indices = Vec::with_capacity(columns.len());
    let mut dtypes = Vec::with_capacity(columns.len());
    let mut column_name_from = Vec::with_capacity(columns.len());
    let mut dtype_from = Vec::with_capacity(columns.len());

    for column in columns {
        names.push(column.name);
        indices.push((column.index + 1) as i32);
        absolute_indices.push((column.absolute_index + 1) as i32);
        dtypes.push(column.dtype.to_string());
        column_name_from.push(column.column_name_from.to_string());
        dtype_from.push(column.dtype_from.to_string());
    }

    Ok(list!(
        name = names,
        index = indices,
        absolute_index = absolute_indices,
        dtype = dtypes,
        column_name_from = column_name_from,
        dtype_from = dtype_from
    ))
}

#[extendr]
fn excel_defined_names(source: Robj, zip_limits: Robj) -> Result<List> {
    let mut reader = reader_from_source(source, zip_limits)?;
    let names = reader.defined_names().map_err(to_r_error)?;

    let name_values: Vec<String> = names.iter().map(|item| item.name.to_string()).collect();
    let formula_values: Vec<String> = names.iter().map(|item| item.formula.to_string()).collect();
    let sheet_values: Vec<String> = std::iter::repeat_n(String::new(), names.len()).collect();

    Ok(list!(
        name = name_values,
        formula = formula_values,
        sheet_name = sheet_values
    ))
}

fn reader_from_source(source: Robj, zip_limits: Robj) -> Result<ExcelReader> {
    let zip_limits = parse_zip_limits(zip_limits)?;

    if let Some(path) = source.as_str() {
        preflight_zip_path(path, &zip_limits)?;
        return read_excel(path).map_err(to_r_error);
    }

    if let Some(bytes) = source.as_raw_slice() {
        preflight_zip_bytes(bytes, &zip_limits)?;
        return ExcelReader::try_from(bytes).map_err(|err| {
            Error::Other(format!("could not load excel file from raw bytes: {err}"))
        });
    }

    Err(Error::Other(
        "`path` must be a single non-empty string or a raw vector.".to_string(),
    ))
}

fn parse_zip_limits(zip_limits: Robj) -> Result<ZipLimits> {
    let values = zip_limits
        .as_real_vector()
        .ok_or_else(|| Error::Other("ZIP preflight limits must be a numeric vector".to_string()))?;
    if values.len() != 4
        || values
            .iter()
            .any(|value| !value.is_finite() || *value <= 0.0)
    {
        return Err(Error::Other(
            "ZIP preflight limits must contain four positive finite numbers".to_string(),
        ));
    }

    Ok(ZipLimits {
        max_entries: values[0] as usize,
        max_entry_size: values[1] as u64,
        max_total_size: values[2] as u64,
        max_compression_ratio: values[3] as u64,
    })
}

fn preflight_zip_path(path: &str, limits: &ZipLimits) -> Result<()> {
    let mut file = match File::open(path) {
        Ok(file) => file,
        Err(_) => return Ok(()),
    };
    let mut magic = [0_u8; 4];
    if file.read_exact(&mut magic).is_err() || !is_zip_magic(&magic) {
        return Ok(());
    }
    preflight_zip(file, limits)
}

fn preflight_zip_bytes(bytes: &[u8], limits: &ZipLimits) -> Result<()> {
    if bytes.len() < 4 || !is_zip_magic(&bytes[..4]) {
        return Ok(());
    }
    preflight_zip(Cursor::new(bytes), limits)
}

fn is_zip_magic(magic: &[u8]) -> bool {
    magic == ZIP_MAGIC || magic == EMPTY_ZIP_MAGIC || magic == SPANNED_ZIP_MAGIC
}

fn preflight_zip<R>(reader: R, limits: &ZipLimits) -> Result<()>
where
    R: Read + Seek,
{
    let mut archive = ZipArchive::new(reader).map_err(zip_preflight_error)?;
    if archive.len() > limits.max_entries {
        return Err(Error::Other(format!(
            "Workbook ZIP contains {} entries, exceeding the limit of {}.",
            archive.len(),
            limits.max_entries
        )));
    }

    let mut total_uncompressed = 0_u64;
    for idx in 0..archive.len() {
        let file = archive.by_index(idx).map_err(zip_preflight_error)?;
        let uncompressed_size = file.size();
        let compressed_size = file.compressed_size();

        if uncompressed_size > limits.max_entry_size {
            return Err(Error::Other(format!(
                "Workbook ZIP entry `{}` expands to {} bytes, exceeding the per-entry limit of {} bytes.",
                file.name(),
                uncompressed_size,
                limits.max_entry_size
            )));
        }

        total_uncompressed = total_uncompressed.saturating_add(uncompressed_size);
        if total_uncompressed > limits.max_total_size {
            return Err(Error::Other(format!(
                "Workbook ZIP expands to more than {} bytes, exceeding the total uncompressed limit.",
                limits.max_total_size
            )));
        }

        if compressed_size > 0 && uncompressed_size / compressed_size > limits.max_compression_ratio
        {
            return Err(Error::Other(format!(
                "Workbook ZIP entry `{}` has a suspicious compression ratio greater than {}:1.",
                file.name(),
                limits.max_compression_ratio
            )));
        }
    }

    Ok(())
}

fn zip_preflight_error(err: ZipError) -> Error {
    Error::Other(format!("could not inspect workbook ZIP metadata: {err}"))
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

fn selected_columns_from_robj(columns: &Robj) -> Result<Option<SelectedColumns>> {
    if let Some(names) = columns.as_str_vector() {
        if names.len() == 1 && names[0] == "NA" {
            return Ok(None);
        }
        return Ok(Some(SelectedColumns::Selection(
            names
                .into_iter()
                .map(|name| IdxOrName::Name(name.to_string()))
                .collect(),
        )));
    }

    if let Some(indices) = columns.as_integer_vector() {
        if indices.len() == 1 && indices[0] == i32::MIN {
            return Ok(None);
        }
        return Ok(Some(SelectedColumns::Selection(
            indices
                .into_iter()
                .map(|idx| IdxOrName::Idx((idx - 1) as usize))
                .collect(),
        )));
    }

    Ok(None)
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

fn column_to_arrow(column: FastExcelColumn) -> Result<(Arc<Field>, ArrayRef)> {
    let name = column.name().to_string();
    let len = column.len();
    let (data_type, values): (DataType, ArrayRef) = match column.data() {
        FastExcelSeries::Null => (DataType::Null, Arc::new(NullArray::new(len))),
        FastExcelSeries::Bool(values) => (
            DataType::Boolean,
            Arc::new(BooleanArray::from(values.to_vec())),
        ),
        FastExcelSeries::String(values) => {
            (DataType::Utf8, Arc::new(StringArray::from(values.to_vec())))
        }
        FastExcelSeries::Int(values) => {
            (DataType::Int64, Arc::new(Int64Array::from(values.to_vec())))
        }
        FastExcelSeries::Float(values) => (
            DataType::Float64,
            Arc::new(Float64Array::from(values.to_vec())),
        ),
        FastExcelSeries::Datetime(values) => {
            let converted = values
                .iter()
                .map(|value| value.and_then(|value| value.and_utc().timestamp_nanos_opt()))
                .collect::<Vec<_>>();
            (
                DataType::Timestamp(TimeUnit::Nanosecond, Some("UTC".into())),
                Arc::new(TimestampNanosecondArray::from(converted)),
            )
        }
        FastExcelSeries::Date(values) => {
            let epoch = NaiveDate::from_ymd_opt(1970, 1, 1)
                .ok_or_else(|| Error::Other("failed to construct Unix epoch date".to_string()))?;
            let converted = values
                .iter()
                .map(|value| value.map(|value| (value - epoch).num_days() as i32))
                .collect::<Vec<_>>();
            (DataType::Date32, Arc::new(Date32Array::from(converted)))
        }
        FastExcelSeries::Duration(values) => {
            let converted = values
                .iter()
                .map(|value| value.map(|value| value.num_milliseconds()))
                .collect::<Vec<_>>();
            (
                DataType::Duration(TimeUnit::Millisecond),
                Arc::new(DurationMillisecondArray::from(converted)),
            )
        }
    };

    Ok((Arc::new(Field::new(name, data_type, true)), values))
}

fn export_columns(
    columns: Vec<FastExcelColumn>,
    array: Robj,
    schema: Robj,
    single_column: bool,
) -> Result<()> {
    if single_column {
        if columns.len() != 1 {
            return Err(Error::Other(
                "`as = \"arrow_array\"` requires exactly one selected column.".to_string(),
            ));
        }
        let (_field, values) = column_to_arrow(columns.into_iter().next().unwrap())?;
        export_array(values, array, schema)?;
    } else {
        let mut arrays = Vec::with_capacity(columns.len());
        for column in columns {
            let (field, values) = column_to_arrow(column)?;
            arrays.push((field, values));
        }
        let struct_array: ArrayRef = Arc::new(StructArray::from(arrays));
        export_array(struct_array, array, schema)?;
    }

    Ok(())
}

fn export_array(values: ArrayRef, array: Robj, schema: Robj) -> Result<()> {
    let array_ptr = external_pointer_addr::<FFI_ArrowArray>(&array)?;
    let schema_ptr = external_pointer_addr::<FFI_ArrowSchema>(&schema)?;
    unsafe {
        let data = values.to_data();
        let ffi_array = FFI_ArrowArray::new(&data);
        let ffi_schema = FFI_ArrowSchema::try_from(data.data_type()).map_err(to_r_error)?;
        std::ptr::write_unaligned(array_ptr, ffi_array);
        std::ptr::write_unaligned(schema_ptr, ffi_schema);
    }
    Ok(())
}

fn external_pointer_addr<T>(ptr: &Robj) -> Result<*mut T> {
    if !ptr.is_external_pointer() {
        return Err(Error::Other("expected an external pointer".to_string()));
    }
    let addr = unsafe { R_ExternalPtrAddr(ptr.get()).cast::<T>() };
    if addr.is_null() {
        return Err(Error::Other("external pointer is NULL".to_string()));
    }
    Ok(addr)
}

fn null_logical_vector(len: usize) -> Robj {
    let values: Vec<Rbool> = std::iter::repeat_n(Rbool::na(), len).collect();
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
    Robj::from(values.to_vec())
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
    fn read_excel_arrow;
    fn read_excel_table_arrow;
    fn excel_sheets;
    fn excel_sheet_info;
    fn excel_sheet_columns;
    fn excel_tables;
    fn excel_table_info;
    fn excel_table_columns;
    fn excel_defined_names;
}
