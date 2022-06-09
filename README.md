`excel_xmlwriter` is library for writing XML in the same format and with
 the same escaping as used by Excel in xlsx xml files.

This is a test crate for a future application and isn't currently
very useful on its own.

```
use std::fs::File;
use excel_xmlwriter::XMLWriter;

fn main() -> Result<(), std::io::Error> {
    let xmlfile = File::create("test.xml")?;
    let mut writer = XMLWriter::new(&xmlfile);

    writer.xml_declaration();

    let attributes = vec![("bar", "1")];
    writer.xml_data_element("foo", "some text", &attributes);

    Ok(())
}
```
Output in `test.xml`:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<foo bar="1">some text</foo>
```