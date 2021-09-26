# exsql-converter
This is my first serious project with more than 1000 line of codes.\
I wrote this project for my university purposes in my second year of study.\
At the beginning of the project, I didn't know anything about Java, PostgreSQL and SQL.

## About the program
The essence of the program is very simple:
1) The program takes some data from .xls files
2) It saves them in PostgreSQL database

## Usage
1) Clear all main tables in PostgreSQL database<br>
<code>java -jar сonverter.jar clear</code>
2) Add data to PostgreSQL database from all .xls files<br>
<code>java -jar сonverter.jar update_all</code>
3) Add data to PostgreSQL database from one specific .xls file<br>
<code>java -jar сonverter.jar update fileName</code>
<br>
<code>fileName</code> - name of an .xls file<br>
<br>
<br>

- Before using the program, you need to make sure that the data located in the <em>database.properties</em> file is valid.
- In order to parse several files at the same time, you MUST move the files to a folder called <em>"files".</em>

## License
Copyright (c) 2021 Shiroganari. [MIT License](https://github.com/Shiroganari/exsql-converter/blob/main/LICENSE)