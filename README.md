# fmdenq 
Package `fmdenq` provides calculations of intra- and interclass correlation coefficients for FMD data.


### Description 
Package `fmdenq` imports FMD result data from a predefined xlsx file, calculates intra- and interclass correlation coefficients on some predefined and some optional custom data groupings / splittings, and saves the statistics results to an xlsx file.

#### Cave 
'fmdenq' is "strictly BETA" and under development, we are in the process of defining the data acquisition protocol and custom-tailoring the pre-build groupings for data analysis.

#### Calculating the ICC
For calculating the resp. ICC, function `ICC` from package `psych` is used.


## Installation
Download from github:
```
library(devtools)
install_github(repo="bpollner/fmdenq", ref="main")
```
*** 
## Usage
### Set up example
First copy the example folder 'FMD' to your desktop:
```
library(fmdenq)
to <- "~/desktop"
from <- paste0(path.package("fmdenq"), "/examples/FMD")
file.copy(from, to, recursive = TRUE) 
```
### Prepare Folder Structure
The folder 'FMD' living on your desktop can be, in this example, the home-folder of a single experiment / measurement series.
Open the R-Studio project called 'FMD' in this folder. As can be seen in the file `script.R`, execute the following code:
```
prepare()
````
This is generating two folders named 'results_fmd' and 'results_stat', holding the results from the FMD-tests and the statistics results, respectively. 
Also, a template of the predefined xlsx file used to collect the results of the FMD tests is copied into the folder 'FMD'. 

### Acquire Data
Now you go and perform your FMD tests, with the results collected in the predefined xlsx-file obtained above. 
For this demonstrational purpose, we will simulate the FMD results by copying a `results_fmd.xlsx` file from package `fmdenq`.
````
to <- "~/desktop/FMD"
from <- paste0(path.package("fmdenq"), "/examples/results_fmd.xlsx")
file.copy(from, to, recursive = TRUE) 
````
Move this file into the folder 'results_fmd'.

### Calculate ICCs
With the simulated results from FMD-tests in place, everything should be ready now to calculate inter- and intraclass correlation coefficients.
For that, there are two options:
* **A)** Step by step: 
```
fmdf <- importData()
res <- calcIntraInter(fmdf)
```
* **B)** All in one step:
```
res <- runFmdenq()
```
Please see ```?runFmdenq``` for further information on input and output parameters.

### Inspect Results
Go to the folder 'results_stat' and inspect the xlsx-file therein.


Enjoy ! 




 
