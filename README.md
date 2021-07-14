# docx_microreader  _(is under development)_

Module for processing docx documents: 
* convert to text, HTML or other formats;
* extract images;
* changes docx files;
* create documents;

_etc._

All these and some other features will be available after the release of the first version of the module (presumably in early 2022)
However, at the moment this module can already perform some tasks.

## At the moment this module can be used for
* convert docx files or parts of it to HTML 
> behavior of convertation can be managed and changed by 
> inheritance custom classes from special Translator classes and overload their methods

* extract images
> except of background images and some others

* editing and formatting documents
> Although the editing interfaces are not very user-friendly right now. This will be fixed soon.

### Simple example of code for convert docx to HTML
```python
from docx_microreader.models import Document

doc = Document("example.docx", "folder for extracted images") # by default second argument is forlder i of example.docx file
with open("result.html", "w", encoding="utf-8") as file:
    file.write(doc1.translate("html"))
```
Example of documents: __left__ - docx, __right__ - HTML.
![Example](https://lh3.googleusercontent.com/S7kcORioPyKoiaYZw7CEhYVH2ANDjqga1HVWnX5xHu-R4CzfoYIicgWGs8aOd0V3mnNEZ_vlZiPBcacmHv-AxTsjqjdltTWyu6IQ3_dqDvNMGp_P_NCypa0kor3agATsYDKYpJqQc2tDyRzFLnPANJpLg-z5VxuJphhGdt8Jb0KWNxIur_MMLDY_R6G6kVc-RyNGfMO-9QsIgDRpK_MdLfX-O1nBzhHExPNH7SU6aV_LhJY9rxDeLpmGEnrDglS5iaHDZkBZVOM7E_62ualP9rgu8NIHHqPklsifsUs2NgRQG8nAnhsHAbx8eiOxAXs4elXLN1D3inKQP1nn6ZkxS13lqdhJs4eM8K6tU_a1OdlcUOT01G-WXp2v_LWjF3knIQx0tgC9nJuEISmB92rWPlE45heSMpXDzL-BgdiDKwtTF6IF1gDiWNne6sHvc3AC05lbzlhlpOacFDCpM5jOmtN9XPW0UGsK0pHB1MOBtF4v0fJHkIqORbzWH6Ud2DYEiyPijR3Gecs_Hj2y_1j4YHLtRin_8Nto9AGIZv0aK2L5Wbu2VpJIHmQ5PM0XMFkpIUzguayDZxTjR4OHtIQQiZWdigGT0sV84ooDCOpdFB1TskKJy39VZVmD78Qe4vaIDtfvnLW-MmiQmGJiktgGDZPri0AnueuAtNw76ei_dWR6qpHv7J-ecC-Qba3es9KpyLsRBeP6PtGVP8COow=w1366-h434-no?authuser=0)

### Simple example of code for editing docx file

Insert word "Text" to paragraphs, that is children of Body of document and haven't runs.

```python
from docx_microreader import models as docx

doc1 = docx.Document("test/file.docx")

body: docx.Body = doc1.get_inner_element(0)
for element in body.iterate_by_inner_elements():
    if isinstance(element, docx.Paragraph) and element.count_inner_elements() == 0:
        new_run = docx.Run.create()
        element.append_inner_element(new_run)
        new_text = docx.Text.create()
        new_text.content = "Text"
        new_run.append_inner_element(new_text)

doc1.save_as_docx("test/result.docx")
```

At the moment __save_as_docx__ method don't save styles of document.
This will be fixed soon.
