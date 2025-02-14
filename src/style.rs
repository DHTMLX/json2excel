use std::collections::HashMap;

use crate::xml::Element;

#[derive(PartialEq)]
pub struct Font {
    pub name: Option<String>,
    pub size: Option<String>,
    pub color: Option<String>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
}

#[derive(PartialEq)]
pub struct Fill {
    pub pattern_type: String,
    pub color: Option<String>,
}

#[derive(PartialEq)]
pub struct Border {
    pub top: Option<BorderProps>,
    pub right: Option<BorderProps>,
    pub bottom: Option<BorderProps>,
    pub left: Option<BorderProps>,
}

pub struct StyleProps {
    pub font: Option<Font>,
    pub fill: Option<Fill>,
    pub border: Option<Border>,
    pub align_h: Option<String>,
    pub align_v: Option<String>,
}

#[derive(PartialEq, Debug)]
pub enum BorderStyle {
    Medium,
    Dotted,
    Thick,
    Thin,
    Dashed,
    Double,
}

#[derive(PartialEq, Debug)]
pub enum BorderPosition {
    Top,
    Right,
    Bottom,
    Left,
}

#[derive(PartialEq)]
pub struct BorderProps {
    pub position: BorderPosition,
    pub size: f32,
    pub style: BorderStyle,
    pub color: String,
}

pub struct StyleTable {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub borders: Vec<Border>,
    pub xfs: Vec<XFSProps>,
    pub custom_formats: HashMap<String, u32>,
    next_custom_format: u32,
}

pub struct XFSProps {
    pub font_id: Option<usize>,
    pub border_id: Option<usize>,
    pub fill_id: Option<usize>,
    pub format_id: Option<usize>,
    pub align_h: Option<String>,
    pub align_v: Option<String>,
}

impl StyleTable {
    pub fn new(css: Option<Vec<HashMap<String, String>>>) -> StyleTable {
        let mut table = StyleTable {
            fonts: vec![Font::new()],
            fills: vec![Fill::new(None, "none"), Fill::new(None, "gray125")],
            borders: vec![Border::new()],
            xfs: vec![XFSProps::new()],
            custom_formats: HashMap::new(),
            next_custom_format: 165,
        };
        table.custom_formats.insert(String::from("General"), 164);

        css.map(|map| {
            for style in map {
                table.add(style);
            }
        });

        table
    }
    pub fn add(&mut self, style: HashMap<String, String>) {
        let mut xsf_props: XFSProps = XFSProps::new();
        let st = style_to_props(&style);

        xsf_props.font_id = st
            .font
            .map(|font| match self.fonts.iter().position(|f| f == &font) {
                Some(v) => v,
                None => {
                    self.fonts.push(font);
                    self.fonts.len() - 1
                }
            });
        xsf_props.fill_id = st
            .fill
            .map(|fill| match self.fills.iter().position(|f| f == &fill) {
                Some(v) => v,
                None => {
                    self.fills.push(fill);
                    self.fills.len() - 1
                }
            });
        xsf_props.border_id =
            st.border.map(
                |border| match self.borders.iter().position(|b| b == &border) {
                    Some(v) => v,
                    None => {
                        self.borders.push(border);
                        self.borders.len() - 1
                    }
                },
            );
        xsf_props.align_h = st.align_h;
        xsf_props.align_v = st.align_v;
        xsf_props.format_id = style.get("format").map(|format_name| {
            let i = match get_format_code(format_name) {
                Some(format) => format,
                None => {
                    if self.custom_formats.contains_key(format_name) {
                        self.custom_formats.get(format_name).unwrap().to_owned()
                    } else {
                        let index = self.next_custom_format;
                        self.custom_formats.insert(format_name.to_owned(), index);
                        self.next_custom_format += 1;
                        index
                    }
                }
            };
            i as usize
        });

        self.xfs.push(xsf_props);
    }
}

impl Fill {
    pub fn new(color: Option<String>, pattern_type: &str) -> Fill {
        Fill {
            pattern_type: pattern_type.to_owned(),
            color: color,
        }
    }
}

impl Font {
    pub fn new() -> Font {
        Font {
            name: None,
            size: None,
            color: None,
            bold: false,
            italic: false,
            underline: false,
            strike: false,
        }
    }
}

impl Border {
    pub fn new() -> Border {
        Border {
            top: None,
            right: None,
            bottom: None,
            left: None,
        }
    }
}

impl BorderProps {
    pub fn to_xml_el(&self) -> Element {
        let el_name = match self.position {
            BorderPosition::Top => "top",
            BorderPosition::Right => "right",
            BorderPosition::Bottom => "bottom",
            BorderPosition::Left => "left",
        };

        let mut el = Element::new(el_name);

        let attr_style = match self.style {
            BorderStyle::Dotted => "dotted",
            BorderStyle::Thick => "thick",
            BorderStyle::Double => "double",
            BorderStyle::Dashed => "hair",
            BorderStyle::Medium => "medium",
            BorderStyle::Thin => "thin",
        };

        el.add_attr("style", attr_style);

        if !self.color.is_empty() {
            let mut el_color = Element::new("color");
            el_color.add_attr("rgb", self.color.to_owned());
            el.add_children(vec![el_color]);
        }

        el
    }
}

impl StyleProps {
    pub fn new() -> StyleProps {
        StyleProps {
            align_h: None,
            fill: None,
            font: None,
            border: None,
            align_v: None,
        }
    }
}

impl XFSProps {
    pub fn new() -> XFSProps {
        XFSProps {
            font_id: None,
            fill_id: None,
            border_id: None,
            format_id: None,
            align_h: None,
            align_v: None,
        }
    }
}

fn style_to_props(styles: &HashMap<String, String>) -> StyleProps {
    let mut font: Font = Font::new();
    let mut border: Border = Border::new();
    let mut st = StyleProps::new();

    for (key, value) in styles {
        match key.as_ref() {
            "background" => match color_to_argb(value) {
                Some(v) => st.fill = Some(Fill::new(Some(v), "solid")),
                None => (),
            },
            "color" => font.color = color_to_argb(value),
            "fontWeight" => font.bold = value == "bold",
            "fontStyle" => font.italic = value == "italic",
            "fontFamily" => font.name = Some(value.to_string()),
            "textDecoration" => {
                font.underline = value.contains("underline");
                font.strike = value.contains("line-through");
            }
            "fontSize" => font.size = px_to_pt(&value),
            "align" => st.align_h = Some(value.to_owned()),
            "verticalAlign" => st.align_v = Some(value.to_owned()),
            "borderTop" => border.top = str_to_border(&value, BorderPosition::Top),
            "borderRight" => border.right = str_to_border(&value, BorderPosition::Right),
            "borderBottom" => border.bottom = str_to_border(&value, BorderPosition::Bottom),
            "borderLeft" => border.left = str_to_border(&value, BorderPosition::Left),
            _ => (),
        }
    }

    st.font = Some(font);
    st.border = Some(border);

    st
}

fn color_to_argb(color: &str) -> Option<String> {
    let len = color.len();
    let mut argb_color = String::new();
    if len == 4 && &color[0..1] == "#" {
        let hex3 = &color[1..4];
        let r = &hex3[0..1];
        let g = &hex3[1..2];
        let b = &hex3[2..3];
        let argb = format!("FF{}{}{}{}{}{}", r, r, g, g, b, b);
        Some(argb)
    } else if len == 7 && &color[0..1] == "#" {
        argb_color.push_str("FF");
        argb_color.push_str(&color[1..]);
        Some(argb_color)
    } else if len > 11 && &color[0..5] == "rgba(" && &color[len - 1..] == ")" {
        let colors_part = &color[5..len - 1];
        let colors = colors_part
            .split(",")
            .map(|s| s.trim())
            .collect::<Vec<&str>>();
        if colors.len() < 4 {
            return None;
        }
        let r = str_to_hex(colors[0]);
        let g = str_to_hex(colors[1]);
        let b = str_to_hex(colors[2]);
        let a = str_alpha_to_hex(colors[3]);
        if r.is_none() || g.is_none() || b.is_none() || a.is_none() {
            return None;
        }
        argb_color.push_str(&a.unwrap());
        argb_color.push_str(&r.unwrap());
        argb_color.push_str(&g.unwrap());
        argb_color.push_str(&b.unwrap());
        Some(argb_color)
    } else if len > 10 && &color[0..4] == "rgb(" && &color[len - 1..] == ")" {
        let colors_part = &color[4..len - 1];
        let colors = colors_part
            .split(",")
            .map(|s| s.trim())
            .collect::<Vec<&str>>();
        if colors.len() < 3 {
            return None;
        }
        let r = str_to_hex(colors[0]);
        let g = str_to_hex(colors[1]);
        let b = str_to_hex(colors[2]);
        if r.is_none() || g.is_none() || b.is_none() {
            return None;
        }
        argb_color.push_str("FF");
        argb_color.push_str(&r.unwrap());
        argb_color.push_str(&g.unwrap());
        argb_color.push_str(&b.unwrap());
        Some(argb_color)
    } else {
        None
    }
}

fn str_to_hex(s: &str) -> Option<String> {
    match s.parse::<u32>() {
        Ok(v) => {
            let res = format!("{:X}", v);
            match res.len() {
                1 => Some(String::from("0") + &res),
                2 => Some(res),
                _ => None,
            }
        }
        Err(_) => None,
    }
}

fn str_alpha_to_hex(s: &str) -> Option<String> {
    match s.parse::<f32>() {
        Ok(v) => {
            let res = format!("{:X}", (v * 255f32) as u32);
            match res.len() {
                1 => Some(String::from("0") + &res),
                2 => Some(res),
                _ => None,
            }
        }
        Err(_) => None,
    }
}

fn px_to_pt(size: &str) -> Option<String> {
    let len = size.len();
    if &size[len - 2..].to_owned() != "px" {
        None
    } else {
        match size[0..len - 2].to_owned().parse::<f32>() {
            Ok(v) => Some((v * 0.75).to_string()),
            Err(_) => None,
        }
    }
}

fn get_format_code(format: &str) -> Option<u32> {
    match format {
        "" | "General" => Some(0),
        "0" => Some(1),
        "0.00" => Some(2),
        "#,##0" => Some(3),
        "#,##0.00" => Some(4),
        "0%" => Some(9),
        "0.00%" => Some(10),
        "0.00E+00" => Some(11),
        "# ?/?" => Some(12),
        "# ??/??" => Some(13),
        "mm-dd-yy" => Some(14),
        "d-mmm-yy" => Some(15),
        "d-mmm" => Some(16),
        "mmm-yy" => Some(17),
        "h:mm AM/PM" => Some(18),
        "h:mm:ss AM/PM" => Some(19),
        "h:mm" => Some(20),
        "h:mm:ss" => Some(21),
        "m/d/yy h:mm" => Some(22),
        "#,##0 ;(#,##0)" => Some(37),
        "#,##0 ;[Red](#,##0)" => Some(38),
        "#,##0.00;[Red](#,##0.00)" => Some(40),
        "mm:ss" => Some(45),
        "[h]:mm:ss" => Some(46),
        "mmss.0" => Some(47),
        "##0.0E+0" => Some(48),
        "@" => Some(49),
        _ => None,
    }
}

fn str_to_border(v: &str, pos: BorderPosition) -> Option<BorderProps> {
    let parts = v.split(" ");
    let vals: Vec<&str> = parts.collect();

    if vals.len() != 3 {
        return None;
    }

    let size = match vals[0].to_string().trim_end_matches("px").parse::<f32>() {
        Ok(s) => s,
        Err(_) => return None,
    };

    let style = match vals[1] {
        "thin" => BorderStyle::Thin,
        "dotted" => BorderStyle::Dotted,
        "double" => BorderStyle::Double,
        "dashed" => BorderStyle::Dashed,
        "solid" => {
            let mut st = BorderStyle::Thin;

            if size == 0.5 {
                st = BorderStyle::Thin
            } else if size == 1.0 {
                st = BorderStyle::Medium
            } else if size == 2.0 {
                st = BorderStyle::Thick
            }

            st
        }
        _ => return None,
    };

    let color: String = match color_to_argb(vals[2]) {
        Some(s) => s,
        None => return None,
    };

    Some(BorderProps {
        position: pos,
        size: size,
        style: style,
        color: color,
    })
}

#[test]
fn style_to_props_test() {
    let mut styles: HashMap<String, String> = HashMap::new();
    styles.insert(String::from("background"), String::from("#FF0000"));
    styles.insert(String::from("color"), String::from("#FFFF00"));
    styles.insert(String::from("fontWeight"), String::from("bold"));
    styles.insert(String::from("fontStyle"), String::from("italic"));
    styles.insert(String::from("fontSize"), String::from("24px"));
    styles.insert(String::from("fontFamily"), String::from("Calibri"));
    styles.insert(String::from("textDecoration"), String::from("underline"));
    styles.insert(String::from("align"), String::from("left"));
    styles.insert(String::from("verticalAlign"), String::from("bottom"));
    styles.insert(String::from("borderTop"), String::from("1px solid #9AFF02"));
    styles.insert(
        String::from("borderRight"),
        String::from("1px solid #000000"),
    );

    let st = style_to_props(&styles);

    let font = st.font.unwrap();
    assert_eq!(font.size, Some(String::from("18")));
    assert_eq!(font.color, Some(String::from("FFFFFF00")));
    assert_eq!(font.bold, true);
    assert_eq!(font.name, Some(String::from("Calibri")));
    assert_eq!(font.italic, true);
    assert_eq!(font.underline, true);
    assert_eq!(st.fill.unwrap().color, Some(String::from("FFFF0000")));
    assert_eq!(st.align_h, Some(String::from("left")));
    assert_eq!(st.align_v, Some(String::from("bottom")));

    let border = st.border.unwrap();
    assert_eq!(border.top.unwrap().color, String::from("FF9AFF02"));
    assert_eq!(border.right.unwrap().color, String::from("FF000000"));
}

#[test]
fn str_to_hex_test() {
    assert_eq!(str_to_hex("255"), Some(String::from("FF")));
    assert_eq!(str_alpha_to_hex("0.5"), Some(String::from("7F")));
}

#[test]
fn color_to_argb_test() {
    assert_eq!(
        color_to_argb("rgba(255, 255, 255, 1)"),
        Some(String::from("FFFFFFFF"))
    );
    assert_eq!(
        color_to_argb("rgb(254,254,254)"),
        Some(String::from("FFFEFEFE"))
    );
    assert_eq!(color_to_argb("#FF6347"), Some(String::from("FFFF6347")));
    assert_eq!(color_to_argb("#A4B"), Some(String::from("FFAA44BB")));
}

#[test]
fn str_to_border_test() {
    let mut maybe_border = str_to_border("1px solid #9900CC", BorderPosition::Top);
    let mut b = maybe_border.unwrap();
    assert_eq!(b.color, "FF9900CC");
    assert_eq!(b.position, BorderPosition::Top);
    assert_eq!(b.size, 1.0);
    assert_eq!(b.style, BorderStyle::Medium);

    maybe_border = str_to_border("0.5px solid #B3FFB3", BorderPosition::Right);
    b = maybe_border.unwrap();
    assert_eq!(b.color, "FFB3FFB3");
    assert_eq!(b.position, BorderPosition::Right);
    assert_eq!(b.size, 0.5);
    assert_eq!(b.style, BorderStyle::Thin);

    maybe_border = str_to_border("2px solid #BCD", BorderPosition::Bottom);
    b = maybe_border.unwrap();
    assert_eq!(b.color, "FFBBCCDD");
    assert_eq!(b.position, BorderPosition::Bottom);
    assert_eq!(b.size, 2.0);
    assert_eq!(b.style, BorderStyle::Thick);

    maybe_border = str_to_border("2px dashed #BCD", BorderPosition::Left);
    b = maybe_border.unwrap();
    assert_eq!(b.style, BorderStyle::Dashed);

    maybe_border = str_to_border("2px dotted #BCD", BorderPosition::Left);
    b = maybe_border.unwrap();
    assert_eq!(b.style, BorderStyle::Dotted);

    maybe_border = str_to_border("2px double #BCD", BorderPosition::Left);
    b = maybe_border.unwrap();
    assert_eq!(b.style, BorderStyle::Double);
}
