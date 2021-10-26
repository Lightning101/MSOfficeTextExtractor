
'use strict';
const fs = require('fs').promises;
const StreamZip = require('node-stream-zip');
const striptags = require('striptags');
const path = require('path');

class DocxConversion {

    static async readDoc(filePath) {
        filePath = path.normalize(filePath);
        try {
            const data = await fs.readFile(filePath, 'utf8')
            const strData = data.toString();
            const lines = strData.split(".");
            let outtext = "";
            lines.forEach(line => {
                line = line.replace(/\0/g, '')
                if (line.length == 0) {
                    return;
                } else {
                    outtext += line + "\n";
                }
            })
            return outtext.replace("/[^a-zA-Z0-9\s\,\.\-\t@\/\_\(\)]/g", "");
        } catch (err) {
            console.error(err)
        }
        return "";
    }

    static async readDocx(filePath) {
        filePath = path.normalize(filePath);
        const zip = new StreamZip.async({ file: filePath, storeEntries: true });
        let data = null;
        try {
            const entries = await zip.entries();
            for (const entry of Object.values(entries)) {
                if (entry.name != "word/document.xml") continue;
                const desc = entry.isDirectory ? 'directory' : `${entry.size} bytes`;
                console.log(`Entry ${entry.name}: ${desc}`);
                data = await zip.entryData(entry.name);
            }

            // Do not forget to close the file once you're done

        } catch (error) {
            console.error(error);
        } finally {
            await zip.close();
        }
        let content = data.toString();
        content = content.replace(new RegExp('</w:r></w:p></w:tc><w:tc>', 'g'), " ");
        content = content.replace(new RegExp('</w:r></w:p>', 'g'), '\r\n');
        content = striptags(content);
        return content;

    }

    /************************excel sheet************************************/

    // xlsx_to_text(input_file) {
    //     xml_filename = "xl/sharedStrings.xml"; //content file name
    //     zip_handle = new ZipArchive;
    //     output_text = "";
    //     if (true === zip_handle.open(input_file)) {
    //         if ((xml_index = zip_handle.locateName(xml_filename)) !== false) {
    //             xml_datas = zip_handle.getFromIndex(xml_index);
    //             xml_handle = DOMDocument:: loadXML(xml_datas, LIBXML_NOENT | LIBXML_XINCLUDE | LIBXML_NOERROR | LIBXML_NOWARNING);
    //             output_text = strip_tags(xml_handle.saveXML());
    //         } else {
    //             output_text.= "";
    //         }
    //         zip_handle.close();
    //     } else {
    //         output_text.= "";
    //     }
    //     return output_text;
    // }

    /*************************power point files*****************************/
    // pptx_to_text(input_file) {
    //     zip_handle = new ZipArchive;
    //     output_text = "";
    //     if (true === zip_handle.open(input_file)) {
    //         slide_number = 1; //loop through slide files
    //         while ((xml_index = zip_handle.locateName("ppt/slides/slide".slide_number.".xml")) !== false) {
    //             xml_datas = zip_handle.getFromIndex(xml_index);
    //             xml_handle = DOMDocument:: loadXML(xml_datas, LIBXML_NOENT | LIBXML_XINCLUDE | LIBXML_NOERROR | LIBXML_NOWARNING);
    //             output_text.= strip_tags(xml_handle.saveXML());
    //             slide_number++;
    //         }
    //         if (slide_number == 1) {
    //             output_text.= "";
    //         }
    //         zip_handle.close();
    //     } else {
    //         output_text.= "";
    //     }
    //     return output_text;
    // }
}

module.exports = DocxConversion;