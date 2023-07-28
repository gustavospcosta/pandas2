<h1 align="center"> Python Pandas </h1>
<div dir="auto" align="center">
  <br>
  <a target="_blank" rel="noopener noreferrer nofollow" href="https://raw.githubusercontent.com/devicons/devicon/master/icons/vscode/vscode-original.svg"><img align="center" alt="Gustavo-VSCode" height="30" width="40" src="https://raw.githubusercontent.com/devicons/devicon/master/icons/vscode/vscode-original.svg" style="max-width: 100%;"></a>
  <a target="_blank" rel="noopener noreferrer nofollow" href="https://raw.githubusercontent.com/devicons/devicon/master/icons/python/python-original.svg"><img align="center" alt="Gustavo-Python" height="30" width="40" src="https://raw.githubusercontent.com/devicons/devicon/master/icons/python/python-original.svg" style="max-width: 100%;"></a>
  <a target="_blank" rel="noopener noreferrer nofollow" href="https://raw.githubusercontent.com/devicons/devicon/master/icons/pandas/pandas-original.svg"><img align="center" alt="Gustavo-Pandas" height="30" width="40" src="https://raw.githubusercontent.com/devicons/devicon/master/icons/pandas/pandas-original.svg" style="max-width: 100%;"></a>
</br>
</div>

## Topics
* [Project Description](#project-description) :us:
* [Descrição do Projeto](#descrição-do-projeto) :brazil:
  
## Project Description
<p align="justify">
This project is a structured process to parse and process data from HTML and Excel files. The steps in the process include importing modules, reading HTML and Excel files, parsing and extracting specific elements from HTML, performing data cleaning, merging and filtering on the dataframes, and saving the output to new Excel files. This process is beneficial for preparing data for further data analysis tasks, where clean and correctly formatted data is critical for obtaining reliable insights.
</p>

## Code Description
<p align="justify">
The given script is essentially a structured procedure to handle data extraction, transformation, and loading (ETL). It begins with the import of necessary modules, including pandas for data handling, lxml for HTML parsing, and re for regular expression operations. 
<br>
The script reads data from an HTML file and several Excel files into dataframes. It uses lxml to parse the HTML file and extract specific elements based on XPath queries. The HTML document is broken down into different tables, which are then processed separately. The script also uses pandas to read data from Excel files.
<br>
The transformation process involves cleaning and formatting the dataframes. Duplicate values are removed, specific columns are selected, new columns are added, and filters are applied based on certain conditions. 
<br>
Lastly, the updated dataframes are concatenated and written back to their respective Excel files. String manipulation and regular expressions are used extensively to clean and format the data. The script ensures a systematic workflow from extraction, transformation, and loading.
</p>

## Getting Started
<p align="justify">
To get started, make sure Python is installed in your environment. The script relies heavily on the pandas, lxml, and re libraries, so make sure to install these using pip: `pip install pandas lxml regex`. Replace the placeholders for file paths in the script with the actual paths of your HTML and Excel files. Please also make sure to update the backup directory path.
</p>

## Executing Program
<p align="justify"> 
To run the program, ensure that this script and the HTML and Excel source files are located in the same directory. Open your terminal, navigate to the directory containing the script, and run the script using the command `python script_name.py`. Make sure to replace `script_name.py` with the actual name of your script.
</p>

## Author
<p align="justify"> Gustavo de Souza Pessanha da Costa. 
</p>

## License
<p align="justify"> This project is licensed under the MIT license. 
</p>

:small_orange_diamond: :small_orange_diamond: :small_orange_diamond:
