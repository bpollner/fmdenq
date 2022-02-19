#### General ####
library(fmdenq)

prepare() # if necessary

# now run muscle tests and acquire data

# fill in muscle test results into the provided xlsx template

# move excel file holding fmd results data into the folder 'results_fmd',
# and possibly rename to default input-name 'results_fmd.xlsx'

# A) step by step
fmdf <- importData()
res <- calcIntraInter(fmdf)

# or B) all in one step
res <- runFmdenq()

# go to the folder 'results_stat' and inspect the xlsx-file therein.



#### specific example ####
library(fmdenq)
fmdf <- importData()
tsList <- list(c("Ther#1", "Ther#2", "Ther#3"),
               c("Ther#4", "Ther#5", "Ther#6")
            ) # end list
calcIntraInter(fmdf, ther.sub=tsList, fnOut="#1_results_stat_Thsub")
calcIntraInter(fmdf, ther.sub=tsList, exclOneMT=TRUE, lmer=FALSE, fnOut="#2_results_stat_MuThsub")
# note: when excluding muscle test, the ICC calculation using lmer is prone to not work,
# (probably not enougth data in the demo data file) therefore here "lmer=FALSE"

