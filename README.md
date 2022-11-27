# docx2pdf
convert docx to pdf on Windows by MS Office/WPS Office/LibreOffice

## OS
Only Windows is supported
## Install
You should first install MS Office/WPS Office/LibreOffice
```shell
pip install -r requirements.txt
```
## Usage
```python
from convert import convert_to_pdf
convert_to_pdf("C:/test/out.docx", "C:/test/out.pdf")
```