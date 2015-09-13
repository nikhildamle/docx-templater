# docx-templater

# Features

1) Simple
```py
    from docxtemplater.docxtemplater import write
    replacements = {
        "NAME": "docx-templater"
    }
    write('./docx_samples/sample1.docx', replacements, './output.docx')
```
![Sample Input 1]('screenshots/sample1_docx.png')
![Sample Output 1]('screenshots/output1_docx.png')

2) Loops
```py
    from docxtemplater.docxtemplater import write
    replacements = {
        "PRODUCTS": [
            {"NAME": "Product 1", "ID": "AC456"},
            {"NAME": "Product 2", "ID": "BV564"}
        ]
    }
    write('./docx_samples/sample2.docx', replacements, './output.docx')
```
![Sample Input 2]('screenshots/sample2_docx.png')
![Sample Output 2]('screenshots/output2_docx.png')

3) It retains style in the template
```py
    from docxtemplater.docxtemplater import write
    replacements = {
        "NAME": "docx-templater"
    }
    write('./docx_samples/sample3.docx', replacements, './output.docx')
```
![Sample Input 3]('screenshots/sample3_docx.png')
![Sample Output 3]('screenshots/output3_docx.png')

4) It can replace placeholders inside tables
```py
    from docxtemplater.docxtemplater import write
    replacements = {
        "NAME": "docx-templater", "ID": "19"
    }
    write('./docx_samples/sample4.docx', replacements, './output.docx')
```
![Sample Input 4]('screenshots/sample4_docx.png')
![Sample Output 4]('screenshots/output4_docx.png')

5) It can loop over tables
```py
    from docxtemplater.docxtemplater import write
    replacements = {
        "PRODUCTS": [
            {"NAME": "Product 1", "ID": "AC456"},
            {"NAME": "Product 2", "ID": "BV564"}
        ]
    }
    write('./docx_samples/sample5.docx', replacements, './output.docx')
```
![Sample Input 5]('screenshots/sample5_docx.png')
![Sample Output 5]('screenshots/output5_docx.png')

6) It can loop over table rows
```py
    from docxtemplater.docxtemplater import write
    replacements = {
        "PRODUCTS": [
            {"NAME": "Product 1", "ID": "AC456"},
            {"NAME": "Product 2", "ID": "BV564"}
        ]
    }
    write('./docx_samples/sample6.docx', replacements, './output.docx')
```
![Sample Input 6]('screenshots/sample6_docx.png')
![Sample Output 6]('screenshots/output6_docx.png')
