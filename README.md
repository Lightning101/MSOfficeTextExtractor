# MSOffice Text Extractor
Small nodejs application to extract text from .doc, .docx, .xls and .ppt
Usage is 

```
const doc = DocxConversion('PATH_TO_FILE');
// If .doc file
doc.read_doc().then(res=>console.log(res));

// If .docx file
doc.read_docx().then(res=>console.log(res));
```

This code is a ported version of the PHP solution given in this stackoverflow question https://stackoverflow.com/questions/19503653/how-to-extract-text-from-word-file-doc-docx-xlsx-pptx-php
