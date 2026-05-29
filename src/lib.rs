// fastxlsx — fast, forward-only .xlsx export for Python.
//
// Thin PyO3 binding over rust_xlsxwriter's constant_memory mode, which streams
// each completed row to a temp file (flat RSS on large exports). The Python API
// is deliberately append-only: no random cell access, no editing existing files.

use pyo3::exceptions::{PyRuntimeError, PyTypeError};
use pyo3::prelude::*;
use chrono::{DateTime, Datelike, FixedOffset, NaiveDate, NaiveDateTime, Timelike};
use pyo3::types::{PyBool, PyList, PyString};
use rust_xlsxwriter::{ExcelDateTime, Format, Workbook, Worksheet, XlsxError};

fn to_pyerr(err: XlsxError) -> PyErr {
    PyRuntimeError::new_err(err.to_string())
}

#[pyclass(name = "Workbook")]
struct PyWorkbook {
    workbook: Workbook,
    sheets: Vec<Py<PyWorksheet>>,
}

#[pymethods]
impl PyWorkbook {
    #[new]
    fn new() -> Self {
        PyWorkbook {
            workbook: Workbook::new(),
            sheets: Vec::new(),
        }
    }

    fn add_sheet(&mut self, py: Python<'_>, name: &str) -> PyResult<Py<PyWorksheet>> {
        let mut worksheet = self.workbook.new_worksheet_with_constant_memory();
        worksheet.set_name(name).map_err(to_pyerr)?;
        let py_worksheet = Py::new(
            py,
            PyWorksheet {
                inner: Some(worksheet),
                next_row: 0,
                date_format: Format::new().set_num_format("yyyy-mm-dd"),
                datetime_format: Format::new().set_num_format("yyyy-mm-dd hh:mm:ss"),
            },
        )?;
        self.sheets.push(py_worksheet.clone_ref(py));
        Ok(py_worksheet)
    }

    fn save(&mut self, py: Python<'_>, path: &str) -> PyResult<()> {
        for sheet in self.sheets.drain(..) {
            if let Some(worksheet) = sheet.bind(py).borrow_mut().inner.take() {
                self.workbook.push_worksheet(worksheet);
            }
        }
        self.workbook.save(path).map_err(to_pyerr)
    }
}

#[pyclass(name = "Worksheet")]
struct PyWorksheet {
    // None once the workbook has been saved (the worksheet is moved into it).
    inner: Option<Worksheet>,
    next_row: u32,
    date_format: Format,
    datetime_format: Format,
}

#[pymethods]
impl PyWorksheet {
    fn append(&mut self, row: &Bound<'_, PyList>) -> PyResult<()> {
        let row_index = self.next_row;
        let date_format = &self.date_format;
        let datetime_format = &self.datetime_format;
        let worksheet = self
            .inner
            .as_mut()
            .ok_or_else(|| PyRuntimeError::new_err("worksheet already saved"))?;
        for (col, value) in row.iter().enumerate() {
            write_value(worksheet, row_index, col as u16, &value, date_format, datetime_format)?;
        }
        self.next_row += 1;
        Ok(())
    }
}

// Map a Python scalar onto the right rust_xlsxwriter writer. Type checks are
// ordered for Python's subclassing: bool before int (bool subclasses int), and
// datetime before date (datetime subclasses date). None leaves the cell blank.
fn write_value(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: &Bound<'_, PyAny>,
    date_format: &Format,
    datetime_format: &Format,
) -> PyResult<()> {
    if value.is_none() {
        return Ok(());
    }
    if let Ok(boolean) = value.downcast::<PyBool>() {
        worksheet
            .write_boolean(row, col, boolean.is_true())
            .map_err(to_pyerr)?;
    } else if let Ok(datetime) = value.extract::<NaiveDateTime>() {
        worksheet
            .write_datetime_with_format(row, col, &excel_datetime(datetime)?, datetime_format)
            .map_err(to_pyerr)?;
    } else if let Ok(aware) = value.extract::<DateTime<FixedOffset>>() {
        // Timezone-aware: write the naive wall-clock fields, dropping the tz.
        worksheet
            .write_datetime_with_format(row, col, &excel_datetime(aware.naive_local())?, datetime_format)
            .map_err(to_pyerr)?;
    } else if let Ok(date) = value.extract::<NaiveDate>() {
        worksheet
            .write_datetime_with_format(row, col, &excel_date(date)?, date_format)
            .map_err(to_pyerr)?;
    } else if let Ok(integer) = value.extract::<i64>() {
        worksheet
            .write_number(row, col, integer as f64)
            .map_err(to_pyerr)?;
    } else if let Ok(number) = value.extract::<f64>() {
        worksheet.write_number(row, col, number).map_err(to_pyerr)?;
    } else if let Ok(text) = value.downcast::<PyString>() {
        worksheet
            .write_string(row, col, text.to_cow()?)
            .map_err(to_pyerr)?;
    } else {
        return Err(PyTypeError::new_err(format!(
            "unsupported cell type: {}",
            value.get_type().name()?
        )));
    }
    Ok(())
}

fn excel_date(date: NaiveDate) -> PyResult<ExcelDateTime> {
    ExcelDateTime::from_ymd(date.year() as u16, date.month() as u8, date.day() as u8).map_err(to_pyerr)
}

fn excel_datetime(datetime: NaiveDateTime) -> PyResult<ExcelDateTime> {
    let seconds = datetime.second() as f64 + datetime.nanosecond() as f64 / 1_000_000_000.0;
    excel_date(datetime.date())?
        .and_hms(datetime.hour() as u16, datetime.minute() as u8, seconds)
        .map_err(to_pyerr)
}

#[pymodule]
fn fastxlsx(module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<PyWorkbook>()?;
    module.add_class::<PyWorksheet>()?;
    Ok(())
}
