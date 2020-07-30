#NOTE: It doesn't work anymore. The API has changed.It must be reviewed.

# vuln_find-report
This python script will find vulnerabilities in Vulners and NVD databases of the SW provided in a XLS file. Then, it creates a XLS and DOCX report.

Pre-requisites
- Sign up in Vulners service https://vulners.com/ . Discover your public IP address and gerenate a new API KEY.
- Python 3.7. (Probably it works with python 2.7 but I didn't test it)
- Install all dependencies from requirements.txt:
    pip install -r requirements.txt
    
Usage
./find_vulns.py <input_file.xls> <output_file_name> <VULNERS_API_KEY>

Demo
NOTE:
Notice that it uses CPE format for searching in NVD database so you need to maintain the CPE names.
See https://nvd.nist.gov/products/cpe for more information. 

Input
- XLS
![alt text](https://github.com/jorgeluengar/vuln_find-report/blob/master/demo_images/demo_input.JPG)

Output
- XLS
![alt text](https://github.com/jorgeluengar/vuln_find-report/blob/master/demo_images/demo_output_XLS.JPG)

- DOCX
![alt text](https://github.com/jorgeluengar/vuln_find-report/blob/master/demo_images/demo_output_docx.JPG)
