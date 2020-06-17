
**AbsentiaVR Coding Assignment**

###  Prerequisites
*  python > 3.6 0r higher
*   VS Code editor (preferred)
###  Getting Started
clone the repo
```bash
$ git clone https://github.com/akshay951228/AbsentiaVR.git
```
###  Installation
*  ``` pip install -r requirement.txt```
### Running:
* Two step process :
    * In this it will  take template and csv raw file and processing data and save in .xlsx file and later we need to open it once ,because XlsxWriter donot compute ,  refer this ```https://stackoverflow.com/a/22492975``` , will give more clarity 
        * python excel.py process --template_path "TEMPLATE_FILE_PATH" --csv_path "CSV_PATH" --output_path "INTERMINATED_FILE_PATH" --output_pkl_path "PKL_SAVING_PATH" --sample "JUST TO A BUNCH"
        * Here i kept sample parameter to process few records like 1k-10k for testing and checking
    * In the second step it just take previous step outputs as inputs and output csv file with required field

### Todo
* Its will take two input and output fields from template , haven't work temp field contains files 
* It will work for formula contain only one parameter(single parameter with multiple in formal will work). for more than 1 unique parameter it will break. 
* Final output saving in not order
* Haven't did any multi-process,for huge data it will take lot of time

### Important points:
* After first step , we need open excel and save it again ,In windows it auto calculate(haven't check), but in libreoffice we need to refresh ```ctrl+shift+f9``` and saving again(I have check and it worked)


## Sample outputs

```https://drive.google.com/drive/folders/17CrjMxgyRCztgGUM-YuETbCBOLj8K87K?usp=sharing```


