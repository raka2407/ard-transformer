# ard-transformer
Transforming data generated from CA ARD to ALM Octane

## Purpose
This custom solution bridges the gap between two widely known tools in the market, namely CA Agile Requirements Designer [ARD] and MicroFocus ALM Octane.

Currently, CA ARD provides an integration capability with HP QC/ALM and automatically uploads the test cases generated from ARD to QC/ALM. However, CA ARD and does not integrate with the advanced version of QC/ALM, which is MicroFocus ALM Octane. 

A custom solution is written in Python which reads the data generated from CA ARD [.xlsx], transforms and writes the data [.xlsx] to Octane undestandable format. The transformed data in [.xlsx] is further used for import into ALM Octane.

For the easy of usage, custom solution is wrapped as a standard python package using PYPI and made available as **ard-transformer**   

![Transformer](https://upload.wikimedia.org/wikipedia/commons/6/60/ARD-Octane.jpg)


## Usage

#### Prerequisites
1. Python Intepreter to be installed on the machine where this utility is planned to be run. Python can be downloaded from https://www.python.org/downloads/

#### Installation
1. Run the following command to install the package
    
    pip install ard-transformer

#### Execution
1. Utility can be run using the command below

       transformer

2. Use one of the command below to list the possible command line arguments

       transformer -h
       transformer --help

3. Usage 
    
        transformer [-h] [-v] [-p PATH] [-i INPUT] [-o OUTPUT]
   
        `Arguments`

            -p PATH     Path of input and output files
            -i INPUT    Name of ARD file
            -o OUTPUT   Name of Octane file
          
3. Below is a sample example

        transformer -p C:/work -i ard.xlsx -o octane.xlsx
    
 
