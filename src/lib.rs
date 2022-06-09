//! # Excel_XMLWriter
//!
//! `excel_xmlwriter` is library for writing XML in the same format and with
//!  the same escaping as used by Excel in xlsx xml files.
//!
//! This is a test crate for a future application and isn't currently
//! very useful on its own.
//!
//!
//! ```
//! use std::fs::File;
//! use excel_xmlwriter::XMLWriter;
//!
//! fn main() -> Result<(), std::io::Error> {
//!     let xmlfile = File::create("test.xml")?;
//!     let mut writer = XMLWriter::new(&xmlfile);
//!
//!     writer.xml_declaration();
//!
//!     let attributes = vec![("bar", "1")];
//!     writer.xml_data_element("foo", "some text", &attributes);
//!
//!     Ok(())
//! }
//!```
//! Output in `test.xml`:
//!
//! ```xml
//! <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
//! <foo bar="1">some text</foo>
//! ```
// SPDX-License-Identifier: MIT
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use std::fs::File;
use std::io::Write;

pub struct XMLWriter<'a> {
    xmlfile: &'a File,
}

impl<'a> XMLWriter<'a> {
    /// Create a new XMLWriter struct to write XML to a given filehandle.
    /// ```
    /// # use std::fs::File;
    /// # use excel_xmlwriter::XMLWriter;
    /// #
    /// # fn main() -> Result<(), std::io::Error> {
    /// let xmlfile = File::create("test.xml")?;
    /// let mut writer = XMLWriter::new(&xmlfile);
    /// #
    /// # Ok(())
    /// # }
    /// ```
    pub fn new(xmlfile: &File) -> XMLWriter {
        XMLWriter { xmlfile }
    }

    /// Write an XML file declaration.
    /// ```
    /// # use std::fs::File;
    /// # use excel_xmlwriter::XMLWriter;
    /// #
    /// # fn main() -> Result<(), std::io::Error> {
    /// # let xmlfile = File::create("test.xml")?;
    /// # let mut writer = XMLWriter::new(&xmlfile);
    /// #
    /// writer.xml_declaration();
    /// // Output: <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    /// #
    /// # Ok(())
    /// # }
    ///
    pub fn xml_declaration(&mut self) {
        writeln!(
            &mut self.xmlfile,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#
        )
        .expect("Couldn't write to file");
    }

    /// Write an XML start tag with attributes.
    /// ```
    /// # use std::fs::File;
    /// # use excel_xmlwriter::XMLWriter;
    /// #
    /// # fn main() -> Result<(), std::io::Error> {
    /// # let xmlfile = File::create("test.xml")?;
    /// # let mut writer = XMLWriter::new(&xmlfile);
    /// #
    /// let attributes = vec![("bar", "1")];
    /// writer.xml_data_element("foo", "some text", &attributes);
    /// // Output: <foo bar="1">some text</foo>
    /// #
    /// # Ok(())
    /// # }
    /// ```
    pub fn xml_start_tag(&mut self, tag: &str, attributes: &Vec<(&str, &str)>) {
        let mut attribute_str = String::from("");

        for attribute in attributes {
            let pair = format!(r#" {}="{}""#, attribute.0, escape_attributes(attribute.1));
            attribute_str.push_str(&pair);
        }

        write!(&mut self.xmlfile, r"<{}{}>", tag, attribute_str).expect("Couldn't write to file");
    }

    /// Write an XML end tag.
    /// ```
    /// # use std::fs::File;
    /// # use excel_xmlwriter::XMLWriter;
    /// #
    /// # fn main() -> Result<(), std::io::Error> {
    /// # let xmlfile = File::create("test.xml")?;
    /// # let mut writer = XMLWriter::new(&xmlfile);
    /// #
    /// writer.xml_end_tag("foo");
    /// // Output: </foo>
    /// // Output: <foo bar="1">some text</foo>
    /// #
    /// # Ok(())
    /// # }
    /// ```
    pub fn xml_end_tag(&mut self, tag: &str) {
        write!(&mut self.xmlfile, r"</{}>", tag).expect("Couldn't write to file");
    }

    /// Write an empty XML tag with attributes.
    /// ```
    /// # use std::fs::File;
    /// # use excel_xmlwriter::XMLWriter;
    /// #
    /// # fn main() -> Result<(), std::io::Error> {
    /// # let xmlfile = File::create("test.xml")?;
    /// # let mut writer = XMLWriter::new(&xmlfile);
    /// #
    /// let attributes = vec![("bar", "1"), ("car", "y")];
    /// writer.xml_empty_tag("foo", &attributes);
    /// // Output: <foo bar="1" car="y"/>
    /// #
    /// # Ok(())
    /// # }
    /// ```
    pub fn xml_empty_tag(&mut self, tag: &str, attributes: &Vec<(&str, &str)>) {
        let mut attribute_str = String::from("");

        for attribute in attributes {
            let pair = format!(r#" {}="{}""#, attribute.0, escape_attributes(attribute.1));
            attribute_str.push_str(&pair);
        }

        write!(&mut self.xmlfile, r"<{}{}/>", tag, attribute_str).expect("Couldn't write to file");
    }

    /// Write an XML element containing data with optional attributes.
    /// ```
    /// # use std::fs::File;
    /// # use excel_xmlwriter::XMLWriter;
    /// #
    /// # fn main() -> Result<(), std::io::Error> {
    /// # let xmlfile = File::create("test.xml")?;
    /// # let mut writer = XMLWriter::new(&xmlfile);
    /// #
    /// let attributes = vec![("bar", "1")];
    /// writer.xml_data_element("foo", "some text", &attributes);
    /// // Output: <foo bar="1">some text</foo>
    /// #
    /// # Ok(())
    /// # }
    /// ```
    pub fn xml_data_element(&mut self, tag: &str, data: &str, attributes: &Vec<(&str, &str)>) {
        let mut attribute_str = String::from("");

        for attribute in attributes {
            let pair = format!(r#" {}="{}""#, attribute.0, escape_attributes(attribute.1));
            attribute_str.push_str(&pair);
        }

        write!(
            &mut self.xmlfile,
            r"<{}{}>{}</{}>",
            tag,
            attribute_str,
            escape_data(data),
            tag
        )
        .expect("Couldn't write to file");
    }

    /// Optimized tag writer for `<c>` cell string elements in the inner loop.
    pub fn xml_string_element(&mut self, index: u32, attributes: &Vec<(&str, &str)>) {
        let mut attribute_str = String::from("");

        for attribute in attributes {
            let pair = format!(r#" {}="{}""#, attribute.0, escape_attributes(attribute.1));
            attribute_str.push_str(&pair);
        }

        write!(
            &mut self.xmlfile,
            r#"<c{} t="s"><v>{}</v></c>"#,
            attribute_str, index
        )
        .expect("Couldn't write to file");
    }

    /// Optimized tag writer for `<c>` cell number elements in the inner loop.
    pub fn xml_number_element(&mut self, number: f64, attributes: &Vec<(&str, &str)>) {
        // TODO: make this generic with the previous function.
        let mut attribute_str = String::from("");

        for attribute in attributes {
            let pair = format!(r#" {}="{}""#, attribute.0, escape_attributes(attribute.1));
            attribute_str.push_str(&pair);
        }

        write!(
            &mut self.xmlfile,
            r#"<c{} t="s"><v>{}</v></c>"#,
            attribute_str, number
        )
        .expect("Couldn't write to file");
    }

    /// Optimized tag writer for `<c>` cell formula elements in the inner loop.
    pub fn xml_formula_element(
        &mut self,
        formula: &str,
        result: f64,
        attributes: &Vec<(&str, &str)>,
    ) {
        let mut attribute_str = String::from("");

        for attribute in attributes {
            let pair = format!(r#" {}="{}""#, attribute.0, escape_attributes(attribute.1));
            attribute_str.push_str(&pair);
        }

        write!(
            &mut self.xmlfile,
            r#"<c{}><f>{}</f><v>{}</v></c>"#,
            attribute_str,
            escape_data(formula),
            result
        )
        .expect("Couldn't write to file");
    }

    /// Optimized tag writer for shared strings `<si>` elements.
    pub fn xml_si_element(&mut self, string: &str, attributes: &Vec<(&str, &str)>) {
        let mut attribute_str = String::from("");

        for attribute in attributes {
            let pair = format!(r#" {}="{}""#, attribute.0, escape_attributes(attribute.1));
            attribute_str.push_str(&pair);
        }

        write!(
            &mut self.xmlfile,
            r#"<si><t{}>{}</t></si>"#,
            attribute_str,
            escape_data(string)
        )
        .expect("Couldn't write to file");
    }

    /// Optimized tag writer for shared strings <si> rich string elements.
    pub fn xml_rich_si_element(&mut self, string: &str) {
        write!(&mut self.xmlfile, r#"<si>{}</si>"#, string).expect("Couldn't write to file");
    }
}

// Escape XML characters in attributes.
fn escape_attributes(attribute: &str) -> String {
    attribute
        .replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('\n', "&#xA;")
}

// Escape XML characters in data sections of tags.  Note, this
// is different from escape_attributes() because double quotes
// and newline are not escaped by Excel.
fn escape_data(attribute: &str) -> String {
    attribute
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

#[cfg(test)]
mod tests {

    use super::XMLWriter;
    use std::fs::File;
    use std::io::{Read, Seek, SeekFrom};
    use tempfile::tempfile;

    use pretty_assertions::assert_eq;

    fn read_xmlfile_data(tempfile: &mut File) -> String {
        let mut got = String::new();
        tempfile.seek(SeekFrom::Start(0)).unwrap();
        tempfile.read_to_string(&mut got).unwrap();
        got
    }

    #[test]
    fn test_xml_declaration() {
        let expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_declaration();

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_start_tag() {
        let expected = "<foo>";
        let attributes = vec![];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_start_tag("foo", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_start_tag_with_attributes() {
        let expected = r#"<foo span="8" baz="7">"#;
        let attributes = vec![("span", "8"), ("baz", "7")];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_start_tag("foo", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_end_tag() {
        let expected = "</foo>";

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_end_tag("foo");

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_empty_tag() {
        let expected = "<foo/>";
        let attributes = vec![];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_empty_tag("foo", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_empty_tag_with_attributes() {
        let expected = r#"<foo span="8"/>"#;
        let attributes = vec![("span", "8")];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_empty_tag("foo", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_data_element() {
        let expected = r#"<foo>bar</foo>"#;
        let attributes = vec![];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_data_element("foo", "bar", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_data_element_with_attributes() {
        let expected = r#"<foo span="8">bar</foo>"#;
        let attributes = vec![("span", "8")];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_data_element("foo", "bar", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_data_element_with_escapes() {
        let expected = r#"<foo span="8">&amp;&lt;&gt;"</foo>"#;
        let attributes = vec![("span", "8")];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_data_element("foo", "&<>\"", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_string_element() {
        let expected = r#"<c span="8" t="s"><v>99</v></c>"#;
        let attributes = vec![("span", "8")];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_string_element(99, &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_number_element() {
        let expected = r#"<c span="8" t="s"><v>99</v></c>"#;
        let attributes = vec![("span", "8")];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_number_element(99.0, &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_formula_element() {
        let expected = r#"<c span="8"><f>1+2</f><v>3</v></c>"#;
        let attributes = vec![("span", "8")];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_formula_element("1+2", 3.0, &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_si_element() {
        let expected = r#"<si><t span="8">foo</t></si>"#;
        let attributes = vec![("span", "8")];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_si_element("foo", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_rich_si_element() {
        let expected = r#"<si>foo</si>"#;

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_rich_si_element("foo");

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }
}
