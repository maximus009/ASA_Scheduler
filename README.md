# ASA_Scheduler
Sheet Parsing and Sorting for Newcomb Centers and Services @ UVA.
Made with the intent of organizing the layout of the sheets to make it convenient for the employees.
And yes, to save paper!

### Specific to Newcomb Hall. Not for Public use.

SimpleHTTP based Web app hosted on the cs.virginia servers.
### How to use the app

Upload the conventional event schedules sheet.
Download the well-organized sheet.

You can view your previous uploaded and processed files on the same page.

#### Requirements
1. xlwt
2. xlrd

Both are available on PyPI.

`$ sudo pip install xlrd xlwt xlutils`

### How to run the server

Create the folder you wish to expose.

`$ mkdir web_folder`

`$ cd web_folder`

`$ python ../uploader.py <port_num>`

Access it using <public_ip>:<port_num>

To check the <public_ip>

`$ ifconfig`

### Contact
<a href="http://www.cs.virginia.edu/~ks6cq/" >Shiva</a> for further details.
