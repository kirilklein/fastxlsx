// Go/no-go benchmark: does rust_xlsxwriter's constant_memory mode actually
// keep peak RSS flat on a large export? Run each mode in its own process and
// measure peak RSS externally (e.g. /usr/bin/time -l on macOS).
//
//   cargo run --release --bin bench -- standard
//   cargo run --release --bin bench -- constant
//
// Writes ROWS x COLS cells of mixed types to /tmp.

use rust_xlsxwriter::Workbook;
use std::time::Instant;

const ROWS: u32 = 200_000;
const COLS: u16 = 5;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mode = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "standard".to_string());
    let start = Instant::now();

    let mut workbook = Workbook::new();
    let worksheet = match mode.as_str() {
        "constant" => workbook.add_worksheet_with_constant_memory(),
        _ => workbook.add_worksheet(),
    };

    for row in 0..ROWS {
        worksheet.write_string(row, 0, format!("patient_{row}"))?;
        worksheet.write_number(row, 1, row as f64)?;
        worksheet.write_number(row, 2, (row as f64) * 1.5)?;
        worksheet.write_boolean(row, 3, row % 2 == 0)?;
        worksheet.write_string(row, 4, "NSQIP")?;
        let _ = COLS;
    }

    let path = format!("/tmp/bench_{mode}.xlsx");
    workbook.save(&path)?;

    let elapsed = start.elapsed();
    println!(
        "mode={mode} rows={ROWS} cols={COLS} cells={} time={:.3}s file={path}",
        ROWS as u64 * COLS as u64,
        elapsed.as_secs_f64()
    );
    Ok(())
}
