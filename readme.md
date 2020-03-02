## get the project

1. download the zip 
!['github download button'](https://github.com/MkLHX/python3_wmi_windows10/blob/master/img/img.png)

2. clone the repository
> git clone https://github.com/MkLHX/python3_wmi_windows10.git

## install python3 on windows
[https://www.python.org/downloads/](https://www.python.org/downloads/)

## set python3 virtual environnement 
> c:\<path_to_your_python3>/python.exe -m venv venv

## activate virtual environnement
> venv/Scripts/activate

### if cannot activate venv
run windows Powershell as administrator and execute
>Set-ExecutionPolicy Unrestricted -Force

then retry

> venv/Scripts/activate

## update pip
> python -m pip install --upgrade pip

## install dependencies
> pip3 install -r requirements.txt

## run the script
> c:\<path_to_your_python3>/python.exe main.py