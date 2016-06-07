# excely

# Purpose
Read and write Excel files.

# References

## Python excel info

### Working with Excel Files in Python
http://www.python-excel.org

### openpyxl
https://openpyxl.readthedocs.io/en/default

# Results


---

## Appendix virtual environment venv

The project uses a virtual environment.

https://docs.python.org/3/library/venv.html

This can hold a python version and pip installed packages such as "openpyxl".

https://github.com/kennethreitz/requests

### Install virtual environment in directory named "venv"

    $ cd <project root directory>
    $ pyvenv venv

### Before activating virtual environment

On my machine, active python is 2.7.11

    ➜  excely git:(master) ✗ which python
    /usr/local/bin/python
    ➜  excely git:(master) python --version
    Python 2.7.11

On my machine, to use python3 must specify python3

    ➜  excely git:(master) which python3
    /usr/local/bin/python3

### Activate virtual environment

    ➜  excely git:(master) source ./venv/bin/activate

### Now active python is in venv and is version 3.5.1

Notice command prompt shows venv is active

    (venv) ➜  excely git:(master) which python
    /Users/stevebaker/Documents/projects/pythonProjects/excely/venv/bin/python
    (venv) ➜  excely git:(master) python --version
    Python 3.5.1


### Deactivate virtual environment
In shell run deactivate
    (venv) ➜  excely git:(master) ✗ deactivate

## Appendix pip install dependencies
With virtualenv active

    pip install openpyxl

## Appendix clone app from github to another machine
After cloning app from github, activating venv did still showed system python.
Fixed as follows:

    delete ./venv
    Re-run pyvenv venv
    pip re-install packages.
