use std::borrow::Cow;

struct Attr<'a>(Cow<'a, str>, Cow<'a, str>);

pub struct Element<'a> {
    tag: Cow<'a, str>,
    attributes: Vec<Attr<'a>>,
    content: Content<'a>
}

enum Content<'a> {
    Empty,
    Value(Cow<'a, str>),
    Children(Vec<Element<'a>>)
}

impl<'a> Element<'a> {
    pub fn new<S>(tag: S) -> Element<'a> where S: Into<Cow<'a, str>> {
        Element {
            tag: tag.into(),
            attributes: vec!(),
            content: Content::Empty
        }
    }
    pub fn add_attr<S, T>(&mut self, name: S, value: T) -> &mut Self where S: Into<Cow<'a, str>>, T: Into<Cow<'a, str>> {
        self.attributes.push(Attr(name.into(), to_safe_attr_value(value.into())));
        self
    }
    pub fn add_value<S>(&mut self, value: S) where S: Into<Cow<'a, str>> {
        self.content = Content::Value(to_safe_string(value.into()));
    }
    pub fn add_children(&mut self, children: Vec<Element<'a>>) {
        if children.len() != 0 {
            self.content = Content::Children(children);
        }
    }
    pub fn to_xml(&mut self) -> String {
        let mut result = String::new();
        result.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        result.push_str(&self.to_string());
        result
    }
}

impl<'a> ToString for Element<'a> {
    fn to_string(&self) -> String {
        let mut attrs = String::new();
        for Attr(name, value) in &self.attributes {
            let attr = format!(" {}=\"{}\"", name, value);
            attrs.push_str(&attr);
        } 
        let mut result = String::new();
        match self.content {
            Content::Empty => {
                let tag = format!("<{}{}/>", self.tag, attrs);
                result.push_str(&tag);
            },
            Content::Value(ref v) => {
                let tag = format!("<{t}{a}>{c}</{t}>", t = self.tag, a = attrs, c = v);
                result.push_str(&tag);
            },
            Content::Children(ref c) => {
                let mut children = String::new();
                for element in c {
                    children.push_str(&element.to_string());
                }
                let tag = format!("<{t}{a}>{c}</{t}>", t = self.tag, a = attrs, c = children);
                result.push_str(&tag);
            }
        }
        result
    }
}

fn to_safe_attr_value<'a>(input: Cow<'a, str>) -> Cow<'a, str> {
    let mut result = String::new();
    for c in input.chars() {
        match c {
            '"' => result.push_str("&quot;"),
            _ => result.push(c)
        }
    }
    Cow::from(result)
}

fn to_safe_string<'a>(input: Cow<'a, str>) -> Cow<'a, str> {
    let mut result = String::new();
    for c in input.chars() {
        match c {
            '<' => result.push_str("&lt;"),
            '>' => result.push_str("&gt;"),
            '&' => result.push_str("&amp;"),
            _ => result.push(c)
        }
    }
    Cow::from(result)
}

 #[test]
fn element_test() {
    let mut el = Element::new("test");
    el
        .add_attr("val", "42")
        .add_attr("val2", "42")
        .add_value("inner 42");

    assert_eq!(el.to_string(), r#"<test val="42" val2="42">inner 42</test>"#);
}
#[test]
fn element_inner_test() {
    let mut root = Element::new("root");
    root.add_attr("isroot", "true");

    let child1 = Element::new("child1");
    let mut child2 = Element::new("child2");

    let child2_1 = Element::new("child2_1");
    let child2_2 = Element::new("child2_2");

    child2.add_children(vec![child2_1, child2_2]);
    root.add_children(vec![child1, child2]);

    assert_eq!(root.to_string(), r#"<root isroot="true"><child1/><child2><child2_1/><child2_2/></child2></root>"#);
}