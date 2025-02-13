use wasm_bindgen;
use wasm_bindgen::prelude::*;
use std::io::prelude::*;

// serializing helpers
use serde::Deserialize;
use gloo_utils::format::JsValueSerdeExt;

use zip;

use zip::write::FileOptions;

use std::io::Cursor;

use std::collections::HashMap;

type Dict = HashMap<String, String>;

pub mod xml;
use crate::xml::Element;
pub mod style;
use crate::style::StyleTable;
pub mod utils;
pub mod formulas;

const WIDTH_COEF: f32 = 8.5;
const HEIGHT_COEF: f32 = 0.75;

const ROOT_RELS: &'static [u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;


#[derive(Deserialize)]
pub struct ColumnData {
    pub width: f32,
}

#[derive(Deserialize)]
pub struct RowData {
    pub height: f32,
}

#[derive(Deserialize)]
pub struct CellCoords {
    pub column: u32,
    pub row: u32,
}

#[derive(Deserialize)]
pub struct MergedCell {
    pub from: CellCoords,
    pub to: CellCoords,
}

#[derive(Deserialize)]
pub struct Cell {
    pub v: Option<String>,
    pub s: Option<u32>,
}

#[derive(Deserialize)]
pub struct SheetData {
    name: Option<String>,
    cells: Option<Vec<Vec<Option<Cell>>>>,
    plain: Option<Vec<Vec<Option<String>>>>,
    cols: Option<Vec<Option<ColumnData>>>,
    rows: Option<Vec<Option<RowData>>>,
    merged: Option<Vec<MergedCell>>,
}

#[derive(Deserialize)]
pub struct SpreadsheetData {
    data: Vec<SheetData>,
    styles: Option<Vec<Dict>>,
}

struct InnerCell {
    cell: String,
    value: CellValue,
    style: Option<u32>
}

impl InnerCell {
    pub fn new(cell: String, style: &Option<u32>) -> InnerCell {
        InnerCell {
            cell,
            value: CellValue::None,
            style: style.to_owned()
        }
    }
}

enum CellValue {
    None,
    Value(String),
    Formula(String),
    SharedString(u32)
}

#[wasm_bindgen]
pub fn import_to_xlsx(raw_data: &JsValue) -> Vec<u8> {
    utils::set_panic_hook();

    let data: SpreadsheetData = raw_data.into_serde().unwrap();
    let futures = formulas::get_future_functions();

    let mut shared_strings = vec!();
    let mut shared_strings_count = 0;
    let style_table = StyleTable::new(data.styles);

    let mut sheets_info: Vec<(String, String)> = vec!();

    let buf: Vec<u8> = vec!();
    let w = Cursor::new(buf);
    let mut zip = zip::ZipWriter::new(w);
    let options = FileOptions::default().compression_method(zip::CompressionMethod::Stored).unix_permissions(0o755);

    for (sheet_index, sheet) in data.data.iter().enumerate() {
        let mut rows: Vec<Vec<InnerCell>> = vec!();
        match &sheet.cells {
            Some(cells) => {
                for (row_index, row) in cells.iter().enumerate() {
                    let mut inner_row: Vec<InnerCell> = vec!();
                    for (col_index, cell) in row.iter().enumerate() {
                        match cell {
                            Some(cell) => {
                                let cell_name = cell_offsets_to_index(row_index, col_index);
                                let mut inner_cell = InnerCell::new(cell_name, &cell.s);

                                match &cell.v {
                                    Some(value) => {
                                        if !value.is_empty() {
                                            match value.parse::<f64>() {
                                                Ok(_) if !value.starts_with("0") || value.len() == 1 => {
                                                    inner_cell.value = CellValue::Value(value.to_owned());
                                                },
                                                Err(_) | _ => {
                                                    if value.starts_with("=") {
                                                        // [FIXME] formula can be corrupted
                                                        inner_cell.value = CellValue::Formula(formulas::fix_formula(&value[1..], &futures));
                                                    } else {
                                                        shared_strings_count += 1;
                                                        // [FIXME] N^2, slow
                                                        match shared_strings.iter().position(|s| s == value) {
                                                            Some(index) => {
                                                                inner_cell.value = CellValue::SharedString(index as u32);
                                                            },
                                                            None => {
                                                                inner_cell.value = CellValue::SharedString(shared_strings.len() as u32);
                                                                shared_strings.push(value.to_owned());
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    None => ()
                                }

                                inner_row.push(inner_cell);
                            },
                            None => ()
                        }
                    }
                    rows.push(inner_row);
                }
            },
            None => {
                match &sheet.plain {
                    Some(plain) => {
                        for (row_index, row) in plain.iter().enumerate() {
                            let mut inner_row: Vec<InnerCell> = vec!();
                            for (col_index, cell) in row.iter().enumerate() {
                                match cell {
                                    Some(value) => {
                                        let cell_name = cell_offsets_to_index(row_index, col_index);
                                        let mut inner_cell = InnerCell::new(cell_name, &None);
                                        match value.parse::<f64>() {
                                            Ok(_) if !value.starts_with("0") || value.len() == 1 => {
                                                inner_cell.value = CellValue::Value(value.to_owned());
                                            },
                                            Err(_) | _ => {
                                                if value.starts_with("=") {
                                                    // [FIXME] formula can be corrupted
                                                    inner_cell.value = CellValue::Formula(formulas::fix_formula(&value[1..], &futures));
                                                } else {
                                                    inner_cell.value = CellValue::SharedString(shared_strings.len() as u32);
                                                    shared_strings.push(value.to_owned());
                                                }
                                            }
                                        }
                                        inner_row.push(inner_cell);
                                    },
                                    None => ()
                                }
                            }
                            rows.push(inner_row);
                        }
                    },
                    None => ()
                }
            }
        }

        let sheet_info = get_sheet_info(sheet.name.clone(), sheet_index);
        zip.start_file(sheet_info.0.clone(), options).unwrap();
        zip.write_all(get_sheet_data(rows, &sheet.cols, &sheet.rows, &sheet.merged).as_bytes()).unwrap();
        sheets_info.push(sheet_info);
    }

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(ROOT_RELS).unwrap();

    let (content_types, rels, workbook) = get_nav(sheets_info);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(rels.as_bytes()).unwrap();
    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook.as_bytes()).unwrap();
    zip.start_file("xl/sharedStrings.xml", options).unwrap();
    zip.write_all(get_shared_strings_data(shared_strings, shared_strings_count).as_bytes()).unwrap();

    zip.start_file("xl/styles.xml", options).unwrap();
    zip.write_all(get_styles_data(style_table).as_bytes()).unwrap();

    let res = zip.finish().unwrap();
    res.get_ref().to_vec()
}

fn get_styles_data(style_table: StyleTable) -> String {
    let mut style_sheet = Element::new("styleSheet");

    let mut formats_element = Element::new("numFmts");
    let mut formats_children = vec!();
    for (format, index) in &style_table.custom_formats {
        let mut format_element = Element::new("numFmt");
        format_element
            .add_attr("formatCode", format)
            .add_attr("numFmtId", index.to_string());
        formats_children.push(format_element);
    }
    formats_element
        .add_attr("count", style_table.custom_formats.len().to_string())
        .add_children(formats_children);

    let mut fonts_element = Element::new("fonts");
    let fonts_children: Vec<Element> = style_table.fonts.iter().map(|ref font| {
        let mut font_element = Element::new("font");
        let mut font_children = vec!();
        match font.name {
            Some(ref font) => {
                let mut name_el = Element::new("name");
                name_el.add_attr("val", font);
                font_children.push(name_el);
            }
            None => (),
        }
        match font.size {
            Some(ref size) => {
                let mut sz = Element::new("sz");
                sz.add_attr("val", size.to_string());
                font_children.push(sz);
            },
            None => ()
        }
        match font.color {
            Some(ref color) => {
                let mut color_element = Element::new("color");
                color_element.add_attr("rgb", color);
                font_children.push(color_element);
            },
            None => ()
        }
        if font.bold {
            let b = Element::new("b");
            font_children.push(b);
        }
        if font.italic {
            let i = Element::new("i");
            font_children.push(i);
        }
        if font.underline {
            let u = Element::new("u");
            font_children.push(u);
        }
        if font.strike {
            let s = Element::new("strike");
            font_children.push(s);
        }
        font_element.add_children(font_children);
        font_element
    }).collect();

    fonts_element
        .add_attr("count", style_table.fonts.len().to_string())
        .add_children(fonts_children);

    let mut borders_element = Element::new("borders");
    let border_children: Vec<Element> = style_table.borders.iter().map(|border| {
        let mut border_element = Element::new("border");
        let mut children: Vec<Element> = vec![];

        if let Some(b) = &border.left {
            children.push(b.to_xml_el())
        }
        if let Some(b) = &border.right {
            children.push(b.to_xml_el())
        }
        if let Some(b) = &border.top {
            children.push(b.to_xml_el())
        }
        if let Some(b) = &border.bottom {
            children.push(b.to_xml_el())
        }

        border_element.add_children(children);
        border_element
    }).collect();
    borders_element.add_attr("count", border_children.len().to_string());
    borders_element.add_children(vec![Element::new("border")]);
    borders_element.add_children(border_children);


    let mut fills_element = Element::new("fills");
    let fills_children: Vec<Element> = style_table.fills.iter().map(|ref fill| {
        let mut fill_element = Element::new("fill");
        let mut pattern_fill = Element::new("patternFill");
        pattern_fill.add_attr("patternType", &fill.pattern_type);
        fill.color.clone().map(|color| {
            let mut fg_color = Element::new("fgColor");
            let mut bg_color = Element::new("bgColor");
            fg_color.add_attr("rgb", color.clone());
            bg_color.add_attr("rgb", color.clone());
            pattern_fill.add_children(vec![fg_color, bg_color]);
        });
        fill_element.add_children(vec![pattern_fill]);
        fill_element
    }).collect();

    fills_element
        .add_attr("count", style_table.fills.len().to_string())
        .add_children(fills_children);

    let mut cell_xfs = Element::new("cellXfs");
    let xfs_children: Vec<Element> = style_table.xfs.iter().map(|p| {
        let mut xf = Element::new("xf");

        p.font_id.map(|id|{
            xf
                .add_attr("applyFont", "1")
                .add_attr("fontId", id.to_string());
        });
        p.fill_id.map(|id| {
            xf
                .add_attr("applyFill", "1")
                .add_attr("fillId", id.to_string());
        });
        p.format_id.map(|id| {
            xf
                .add_attr("applyNumberFormat", "1")
                .add_attr("numFmtId", id.to_string());
        });            
        p.border_id.map(|id| {
            xf
                .add_attr("applyBorder", "1")
                .add_attr("borderId", id.to_string());
        });
        if p.align_h.is_some() || p.align_v.is_some() {
            let mut alignment = Element::new("alignment");
            p.align_h.as_ref().map(|v| alignment.add_attr("horizontal", v));
            p.align_v.as_ref().map(|v| alignment.add_attr("vertical", v));
            xf.add_attr("applyAlignment", "1").add_children(vec![alignment]);
        }

        xf
    }).collect();

    cell_xfs
        .add_attr("count", style_table.xfs.len().to_string())
        .add_children(xfs_children);

    style_sheet    
        .add_attr("xmlns:xm", "http://schemas.microsoft.com/office/excel/2006/main")
        .add_attr("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
        .add_attr("xmlns:x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
        .add_attr("xmlns:mv", "urn:schemas-microsoft-com:mac:vml")
        .add_attr("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
        .add_attr("xmlns:mx", "http://schemas.microsoft.com/office/mac/excel/2008/main")
        .add_attr("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        .add_attr("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        .add_children(vec![formats_element, fonts_element, fills_element, borders_element, cell_xfs]);
    
    style_sheet.to_xml()
}

fn cell_offsets_to_index(row: usize, col: usize) -> String {
    let mut num = col;
    let mut chars = vec!();

    while num > 25 {
        let part = num % 26;
        num = (num - part) / 26 - 1;
        chars.push(65u8 + part as u8);
    }
    chars.push(65u8 + num as u8);
    chars.reverse();
    format!("{}{}", String::from_utf8(chars).unwrap(), row + 1)
}

fn get_sheet_data(cells: Vec<Vec<InnerCell>>, columns: &Option<Vec<Option<ColumnData>>>, rows: &Option<Vec<Option<RowData>>>, merged: &Option<Vec<MergedCell>>) -> String {
    let mut worksheet = Element::new("worksheet");
    let mut sheet_view = Element::new("sheetView");
    sheet_view.add_attr("workbookViewId", "0");
    let mut sheet_views = Element::new("sheetViews");
    sheet_views.add_children(vec![sheet_view]);
    let mut sheet_format_pr = Element::new("sheetFormatPr");
    sheet_format_pr
        .add_attr("customHeight", "1")
        .add_attr("defaultRowHeight", "15.75")
        .add_attr("defaultColWidth", "14.43");

    let mut cols = Element::new("cols");
    let mut cols_children = vec!();

    match columns {
        Some(columns) => {
            for (index, column) in columns.iter().enumerate() {
                match column {
                    Some(col) => {
                        let mut column_element = Element::new("col");
                        column_element
                            .add_attr("min", (index + 1).to_string())
                            .add_attr("max", (index + 1).to_string())
                            .add_attr("customWidth", "1")
                            .add_attr("width", (col.width / WIDTH_COEF).to_string());
                        cols_children.push(column_element)
                    },
                    None => ()
                }
            }
        },
        None => ()
    }
    let mut rows_info: HashMap<usize, &RowData> = HashMap::new();
    match rows {
        Some(rows) => {
            for (index, column) in rows.iter().enumerate() {
                match column {
                    Some(row) => {
                        rows_info.insert(index, row);
                    },
                    None => ()
                }
            }
        },
        None => ()
    }

    let mut sheet_data = Element::new("sheetData");
    let mut sheet_data_rows = vec!();
    for (index, row) in cells.iter().enumerate() {
        let mut row_el = Element::new("row");
        row_el.add_attr("r", (index + 1).to_string());
        match rows_info.get(&index) {
            Some(row_data) => {
                row_el
                    .add_attr("ht", (row_data.height * HEIGHT_COEF).to_string())
                    .add_attr("customHeight", "1");
            },
            None => ()
        }
        let mut row_cells = vec!();
        for cell in row {
            let mut cell_el = Element::new("c");
            cell_el.add_attr("r", &cell.cell);
            match &cell.value {
                CellValue::Value(ref v) => {
                    let mut value_cell = Element::new("v");
                    value_cell.add_value(v);
                    cell_el.add_children(vec![value_cell]);
                    utils::log!("value {}", v)
                },
                CellValue::Formula(ref v) => {
                    let mut value_cell = Element::new("f");
                    value_cell.add_value(v);
                    cell_el.add_children(vec![value_cell]);
                    utils::log!("formula {}", v)
                },
                CellValue::SharedString(ref s) => {
                    cell_el.add_attr("t", "s");
                    let mut value_cell = Element::new("v");
                    value_cell.add_value(s.to_string());
                    cell_el.add_children(vec![value_cell]);
                },
                CellValue::None => ()
            }
            match &cell.style {
                Some(ref v) => {
                    // style index should be incremented as zero style was prepended
                    cell_el.add_attr("s", (v+1).to_string());
                },
                None => ()
            }
            row_cells.push(cell_el);
        }

        row_el.add_children(row_cells);
        sheet_data_rows.push(row_el);
    }
    sheet_data.add_children(sheet_data_rows);

    let mut worksheet_children = vec![sheet_views, sheet_format_pr];
    if cols_children.len() > 0{ 
        cols.add_children(cols_children);
        worksheet_children.push(cols);
    }
    worksheet_children.push(sheet_data);

    match merged {
        Some(merged) => {
            if merged.len() > 0 {
                let mut merged_cells_element = Element::new("mergeCells");
                merged_cells_element
                    .add_attr("count", merged.len().to_string())
                    .add_children(merged.iter().map(|MergedCell {from, to}| {
                        let p1 = cell_offsets_to_index(from.row as usize, from.column as usize);
                        let p2 = cell_offsets_to_index(to.row as usize, to.column as usize);
                        let cell_ref = format!("{}:{}", p1, p2);
                        let mut merged_cell = Element::new("mergeCell");
                        merged_cell.add_attr("ref", cell_ref);
                        merged_cell
                    }).collect());
                worksheet_children.push(merged_cells_element);
            }
        },
        None => ()
    }

    worksheet
        .add_attr("xmlns:xm", "http://schemas.microsoft.com/office/excel/2006/main")
        .add_attr("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
        .add_attr("xmlns:x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
        .add_attr("xmlns:mv", "urn:schemas-microsoft-com:mac:vml")
        .add_attr("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
        .add_attr("xmlns:mx", "http://schemas.microsoft.com/office/mac/excel/2008/main")
        .add_attr("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        .add_attr("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        .add_children(worksheet_children);

    worksheet.to_xml()
}

fn get_shared_strings_data(shared_strings: Vec<String>, shared_strings_count: i32) -> String {
    let mut sst = Element::new("sst");
    let sst_children: Vec<Element> = shared_strings.iter().map(|s| {
        let mut t = Element::new("t");
        t.add_value(s);
        let mut si = Element::new("si");
        si.add_children(vec![t]);
        si
    }).collect();

    sst
        .add_attr("uniqueCount", shared_strings.len().to_string())
        .add_attr("count", shared_strings_count.to_string())
        .add_attr("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        .add_children(sst_children);
    
    sst.to_xml()
}

fn get_sheet_info(name: Option<String>, index: usize) -> (String, String) {
    let sheet_name = name.unwrap_or(format!("sheet{}", index + 1));
    (format!("xl/worksheets/sheet{}.xml", index + 1), sheet_name)
}

fn get_nav(sheets: Vec<(String, String)>) -> (String, String, String) {
    let mut content_types = Element::new("Types");
    content_types.add_attr("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");

    let mut overrides = vec!();

    let mut default_xml = Element::new("Default");
    default_xml
        .add_attr("ContentType", "application/xml")
        .add_attr("Extension", "xml");
    let mut default_rels = Element::new("Default");
    default_rels
        .add_attr("ContentType", "application/vnd.openxmlformats-package.relationships+xml")
        .add_attr("Extension", "rels");

    let mut root_rels = Element::new("Override");
    root_rels
        .add_attr("ContentType", "application/vnd.openxmlformats-package.relationships+xml")
        .add_attr("PartName", "/_rels/.rels");
    let mut root_workbook_rels = Element::new("Override");
    root_workbook_rels
        .add_attr("ContentType", "application/vnd.openxmlformats-package.relationships+xml")
        .add_attr("PartName", "/xl/_rels/workbook.xml.rels");
    let mut shared_strings_rels = Element::new("Override");
    shared_strings_rels
        .add_attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
        .add_attr("PartName", "/xl/sharedStrings.xml");
    let mut style_rels = Element::new("Override");
    style_rels
        .add_attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
        .add_attr("PartName", "/xl/styles.xml");
    let mut workbook_rels = Element::new("Override");
    workbook_rels
        .add_attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
        .add_attr("PartName", "/xl/workbook.xml");

    overrides.push(default_xml);
    overrides.push(default_rels);
    overrides.push(root_rels);
    overrides.push(root_workbook_rels);
    overrides.push(shared_strings_rels);
    for (path, _) in sheets.iter() {
        let mut sheet_rel = Element::new("Override");
        sheet_rel
            .add_attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
            .add_attr("PartName", String::from("/") + path);
        overrides.push(sheet_rel);
    }
    overrides.push(style_rels);
    overrides.push(workbook_rels);

    content_types.add_children(overrides);

    let mut relationships = Element::new("Relationships");
    relationships.add_attr("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
    let mut relationships_children = vec!();

    let mut style_relationship = Element::new("Relationship");
    style_relationship
        .add_attr("Id", "rId1")
        .add_attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
        .add_attr("Target", "styles.xml");
    let mut shared_string_relationship = Element::new("Relationship");
    shared_string_relationship
        .add_attr("Id", "rId2")
        .add_attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
        .add_attr("Target", "sharedStrings.xml");
    
    relationships_children.push(style_relationship);
    relationships_children.push(shared_string_relationship);

    let mut workbook = Element::new("workbook");
    let mut workbook_children = vec!();
    workbook_children.push(Element::new("workbookPr"));
    let mut sheets_element = Element::new("sheets");
    let mut sheet_children = vec!();

    let mut last_id = 3;
    for (index, (_, name)) in sheets.iter().enumerate() {
        let mut sheet_relationship = Element::new("Relationship");
        sheet_relationship
            .add_attr("Id", format!("rId{}", last_id))
            .add_attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
            .add_attr("Target", format!("worksheets/sheet{}.xml", (index+1)));
        relationships_children.push(sheet_relationship);   
        let mut sheet_element = Element::new("sheet");
        sheet_element
            .add_attr("r:id", format!("rId{}", last_id))
            .add_attr("name", name)
            .add_attr("sheetId", (index+1).to_string());
        
        if index == 0 {
            sheet_element.add_attr("state", "visible");
        }
        sheet_children.push(sheet_element);

        last_id += 1;
    }
    relationships.add_children(relationships_children);
    sheets_element.add_children(sheet_children); 
    workbook_children.push(sheets_element);
    workbook_children.push(Element::new("definedNames"));
    workbook_children.push(Element::new("calcPr"));

    workbook
        .add_attr("xmlns:xm", "http://schemas.microsoft.com/office/excel/2006/main")
        .add_attr("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
        .add_attr("xmlns:x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
        .add_attr("xmlns:mv", "urn:schemas-microsoft-com:mac:vml")
        .add_attr("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
        .add_attr("xmlns:mx", "http://schemas.microsoft.com/office/mac/excel/2008/main")
        .add_attr("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        .add_attr("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        .add_children(workbook_children);
    
    (content_types.to_xml(), relationships.to_xml(), workbook.to_xml())
}