genfs <- function() {
	resFolder <- "results_fmd"
	statFolder <- "results_stat"
	#
	if (!dir.exists(resFolder)) {
		dir.create(resFolder)
	}
	if (!dir.exists(statFolder)) {
		dir.create(statFolder)
	}
} # EOF

copyXlsxTemplate <- function(from="pack") {
	if (from == "pack") {
		from <- paste0(path.package("fmdenq"), "/templates/results_fmd_template.xlsx")
	} # end if
	to <- getwd()
	file.copy(from, to, overwrite=TRUE)
} # EOF

#' @title Generate Folders, Copy Template
#' @description Generate the required folders and copy a template of the xlsx
#' input file, all in the current working directory.
#' @export
prepare <- function() {
	genfs()
	ok <- copyXlsxTemplate()
	if (ok) {
		cat("OK \n")
	}
	return(invisible(NULL))
} # EOF

importDataFromXlsx <- function(fnIn="results_fmd", sel=FALSE) {
	folder <- "results_fmd"
	#
	fnIn <- paste0(folder, "/", fnIn, ".xlsx")
	if (sel == TRUE) {
		fnIn <- try(file.choose(new=FALSE), silent=TRUE)
	}
	if (class(fnIn) == "try-error") {
		stop(paste0("Sorry, choosing the path to a xlsx file holding the FMD-results does not seem to work interactively.\nPlease provide a valid path to an existing xlsx file in the argument 'fnIn' manually.\n"), call.=FALSE)
	} # end if class(where) == "try-error"
	#
	message(paste0("Importing data from file '", basename(fnIn), "'."))
	fmddf <- openxlsx::read.xlsx(fnIn, sheet=1) # comes out as data.frame
	assign("fnIn", basename(fnIn), pos=-1)
	return(fmddf)
} # EOF

checkData <- function(fmddf, filename=NULL, verbose=TRUE) {
	goodColNames <- gcn <-  c("Patient_ID", "Therapist_ID", "Repetition", "MuscleTest_ID", "Result")
	goodRes <- sort(c("w", "n", "ht"))
	#
	# first check the column names
	cns <- sort(colnames(fmddf))
	if (!identical(cns, sort(goodColNames))) {
		stop(paste0("Sorry, the column names in the provided xlsx file '",filename ,"' differ from the required names as provided in the template.\nThe correct column names should be:\n", paste(goodColNames, collapse=", ")), call.=FALSE)
	} # end if !identical
	#
	# describe content of columns
	uniPatID <- sort(unique(fmddf$Patient_ID))
	uniTherID <- sort(unique(fmddf$Therapist_ID))
	uniRep <- sort(unique(fmddf$Repetition))
	uniMtID <- sort(unique(fmddf$MuscleTest_ID))
	uniRes <- sort(unique(fmddf$Result))
	#
	assign("allMuscleTestIDs", uniMtID, envir=fmdenqEnv)
	assign("allTherapistIDs", uniTherID, envir=fmdenqEnv)
	assign("allPatientIDs", uniPatID, envir=fmdenqEnv)
	#
	if (!identical(goodRes, uniRes)) {
		ind <- which(! uniRes %in% goodRes)
		wrongRes <- paste(uniRes[ind], collapse=", ")
		stop(paste0("Sorry, some values for the results (", wrongRes, ") are not within the permitted values (w, n, ht). Please check the entries in the column 'Result'."), call.=FALSE)
	} # end if !identical
	#
	if (verbose) {
		cat("The following values are present in their respective column:\n")
		cat(paste0(gcn[1], ": ", paste(uniPatID, collapse=", "), "\n"))
		cat(paste0(gcn[2], ": ", paste(uniTherID, collapse=", "), "\n"))
		cat(paste0(gcn[3], ": ", paste(uniRep, collapse=", "), "\n"))
		cat(paste0(gcn[4], ": ", paste(uniMtID, collapse=", "), "\n"))
	} # end if verbose
} # EOF

cleanOutRowsAnyNA <- function(fmddf, verbose) {
    indEx <- which(apply(fmddf, 1, function(x) any(is.na(x))))
    if (length(indEx) != 0) {
    	if (verbose) {
	    	cat(paste0("The following ", length(indEx), " row(s) were excluded:\n"))
	    	print(fmddf[indEx,])
    		cat("\n\n")
    	} # end if verbose
    	fmddf <- fmddf[-indEx, ]
    } # end if
    #
    return(fmddf)
} # EOF

exchangeCharForNumbers <- function(fmddf) {
	doChange <- function(wChar, wNr, daf) {
		ind <- which(fmddf$Result == wChar)
		daf[ind, "Result"] <- wNr
		return(daf)
	} # EOIF
	fmddf <- doChange("w", 0, fmddf)
	fmddf <- doChange("n", 1, fmddf)
	fmddf <- doChange("ht", 2, fmddf)
	fmddf$Result <- as.numeric(fmddf$Result)
	return(fmddf)
} # EOF

#' @title Import Data
#' @description Import data from predefined xlsx sheet, check range of result
#' values and describe range of other values.
#' @param fnIn Character length one, the name of the data-input file, what is
#' an xlsx file residing in the folder 'results_fmd'. (The name is provided
#' without the '.xlsx' ending.)
#' @param selIn Logical, if an xlsx file should be selected interactively instead.
#' Defaults to FALSE.
#' @param verbose Logical, if the range of the values of each individual column
#' should be reported. Defaults to TRUE.
#' @return A data frame containing FMD result data.
#' @examples
#' \dontrun{
#' fmdf <- importData() # leave everything at the defaults
#' fmdf <- importData(fnIn="results_fmd_foo") # specify custom name of fmd results
#' # file
#' fmdf <- importData(sel=TRUE) # to interactively select the xlsx file to import
#' }
#' @seealso \code{\link{runFmdenq}}
#' @export
importData <- function(fnIn="results_fmd", selIn=FALSE, verbose=TRUE) {
	fmddf <- importDataFromXlsx(fnIn, selIn) # is assigning !basename fnIn (as it could have been interactively selected as well)
	fmddf <- cleanOutRowsAnyNA(fmddf, verbose)
	checkData(fmddf, filename=fnIn, verbose)
	if (class(fmddf) == "data.frame") {
		message("Data import successful.")
	}
	fmddf <- exchangeCharForNumbers(fmddf)
	return(new("fmddf", fmddf))
} # EOF

#' @title Exclude Single Class Variables
#' @description Define a column, define a value, and exclude or only include these 
#' data
#' @param fmddf A data frame containing FMD data.
#' @param cn Column name
#' @param val Value
#' @param include Logical
#' @return The dataset modified
#' @export
ssc <- function(fmddf, cn, val, include=TRUE) {
	ind <- which(fmddf[,cn] == val)
	if (include) {
		return(fmddf[ind,])
	} else {
		return(fmddf[-ind,])
	}
} # EOF

resh <- function(fmddf) {
	mdf <- reshape2::melt(fmddf, id.vars=1:4, measure.vars=5)
	out <-  reshape2::dcast(mdf, Patient_ID + Therapist_ID + MuscleTest_ID  ~ Repetition)
	return(out)
} # EOF

resh_Inter <- function(fmddf) {
	fmddf <- resh(fmddf) # has all repetitions in extra columns
	therId <- sort(unique(fmddf$Therapist_ID))
	reps <- fmddf[, -(1:3), drop=FALSE] # keep only the numeric values
	if (ncol(reps) > 1) { # so we have at least 2 repeated measurements
		type <- "ICC2k"
	} else { # so we have only one measurement
		type <- "ICC2"
	}
	repsMean <- apply(reps, 1, mean) # possibly verage together all repeates
	repsMean <- round(repsMean, 2)
	meanDf <- cbind(fmddf[, 1:3], repsMean) # we have one column with (possibly) results averaged together
	#
	meanResh <- reshape2::melt(meanDf, id.vars=1:3, measure.vars=4)
	meanResh <- reshape2::dcast(meanResh, Patient_ID  + MuscleTest_ID  ~ Therapist_ID)
	#
	out <- new("fmddf_reshInter", meanResh, type=type, nrReps=ncol(reps), therapist=therId)
	return(out)
} # EOF

resh_Intra <- function(fmddf) {
	fmddf <- resh(fmddf) # has all repetitions in extra columns
	therId <- unique(fmddf$Therapist_ID)
	#
	if (length(therId) != 1) {
		stop(paste0("Sorry, the Therapist_ID has to be unique for intra-class calculations"), call.=FALSE)
	} # end if stop
	#
	fmddf <- fmddf[, -2] # kick out the colummn with therpist id
	#
	if (ncol(fmddf) == 3) {# that means Patient_ID, MuscleTest_ID and only one more for data
		stop(paste0("Sorry, you need at least two repetitions (test, re-test) in order to calculate intra-class statistics"), call.=FALSE)
	} # end if stop
	#
	cns <- colnames(fmddf)
	cnsOld <- cns[1:2] # Patient_ID and MuscleTest_ID
	cnsNew <- cns[-(1:2)] # has all the repetitions
	cnsNew <- paste0(therId, "_", cnsNew)
	colnames(fmddf) <- c(cnsOld, cnsNew)
	return(new("fmddf_reshIntra", fmddf, type="ICC3", nrReps=ncol(fmddf)-2, therapist=therId))
} # EOF


#' @title Get and Display a Subset of the Data
#' @description Use an expression to provide a definition of a subset of the 
#' data, reshape the data to either inter- or intra-class analysis, possibly 
#' print these subset and return it.
#' @details Use 'colnames(object)' to determine the name of the columns in the 
#' data frame. (Of course.)
#' @param fmddf A data frame containing FMD data.
#' @param expr An expression defining the subset of data. If left at the default 
#' NULL, no subset of the data is created, but he data is merely reshaped.
#' @param resh Character length one. What type of reshape should be applied to 
#' the output. Possible values are 'inter' (default) and 'intra'.
#' @param showData Logical If the resulting data frame should be printed. Defaults 
#' to TRUE.
#' @return A data frame with the subset of data, reshaped for either inter- or 
#' intra-class analysis.
#' @examples
#' \dontrun{
#' fmddf <- importData()
#' expr <- expression(Patient_ID == "Pat#2")
#' showGetData(fmddf, expr)
#' expr <- expression(Patient_ID == "Pat#2" & Therapist_ID != "Ther#4")
#' showGetData(fmddf, expr)
#' }
showGetData <- function(fmddf, expr=NULL, resh=NULL, showData=TRUE) {
	pv_resh <- c("inter", "intra")
	#
	if (is.null(resh)) {
		stop("Please provide either 'inter' or 'intra' to the argument 'resh'", call.=FALSE)
	} # end if

	if (class(fmddf) != "fmddf") {
		stop("Please provide an object of class 'fmddf' to the argument 'fmddf'.", call.=FALSE)
	} # end if
	if (!is.expression(expr) & ! is.null(expr)) {
		stop("Please provide an expression to the argument 'expr'", call.=FALSE)
	} # end if
	#
	if (! resh %in% pv_resh) {
		stop("Please provide either 'inter' or 'intra' to the argument 'resh'.", call.=FALSE)
	} # end if
	#
	if (!is.null(expr)) {
		TFsel <- with(fmddf, eval(expr))
		fmddf <- fmddf[TFsel, ]
	} # end if
	
	if (resh == "inter") {
		fmddf <- resh_Inter(fmddf)
	} else { # so it must be intra
		fmddf <- resh_Intra(fmddf)
	} # end else 
	if (showData) {
		print(fmddf)
	} # end if
	return(invisible(fmddf))
} # EOF

#' @title Calculate ICC for Provided Data
#' @description Calculate either ICC2 or ICC3 on the provided data, possibly 
#' print the results and return them.
#' @section CAVE: It is the users responsibility to decide if ICC2 or ICC3 is 
#' appropriate for the provided data input. 
#' @param fmddf A reshaped data frame containing FMD data as produced by 
#' \code{\link{showGetData}}.
#' @param showICC Logical, if the calculated ICC should be printed. Defaults to 
#' TRUE.
#' @param showData Logical, if the input-data should be printed as well. Defaults 
#' to TRUE.
#' @param expr An expression defining the subset of data. If left at the default 
#' NULL, no subset of the data is created, but he data is merely reshaped.
#' @param resh Character length one. What type of reshape should be applied to 
#' the output. Possible values are 'inter' and 'intra'.
#' @inheritParams calcIntraInter
#' @return An (invisible) data frame containing the results of the ICC calculation.
#' @examples
#' \dontrun{
#' XXX
#' }
#' @export
showCalcICC <- function(fmddf, expr=NULL, resh=NULL, showData=TRUE, showICC=TRUE, lmer=FALSE) {
	rndIcc <- 3
	rndpVal <- 4
	#
	if (is.null(resh)) {
		stop("Please provide either 'inter' or 'intra' to the argument 'resh'", call.=FALSE)
	} # end if
	if (! any(grepl("fmddf", class(fmddf)))) {
		stop("Please provide an fmdenq data object to the argument 'fmddf'", call.=FALSE)
	} #
	if (is.expression(expr) & class(fmddf) != "fmddf") {
		stop(paste0("Please provide an object of class 'fmddf' to the argument 'fmddf' in order to properly generate data subset"), call.=FALSE)
	} #
	if (!is.expression(expr) & ! is.null(expr)) {
		stop("Please provide an expression to the argument 'expr'", call.=FALSE)
	} # end if
	#
	if (class(fmddf) == "fmddf") {
		fmddf <- showGetData(fmddf, expr, resh, showData=FALSE) # here gets reshaped
	} # end if
	#
	iccList <- calcICC(fmddf, useLmer=lmer)
	icc <- round(iccList$icc, rndIcc)
	pVal <- round(iccList$pVal, rndpVal)
	nrJ <- iccList$nrJ
	nrObs <- iccList$nrObs
	type <- fmddf@type
	nrReps <- fmddf@nrReps
	if (grepl("ICC2", type)) { # so we have INTER
		#
	} # end if
	if (grepl("ICC3", type)) { # so we have INTRA
		nrJ <- 1
	} # end if
	#
	out <- data.frame(icc, pVal, type, nrJ, nrObs, nrReps, lmer)
	colnames(out) <- c("ICC", "p-value", "Type", "nrJ", "nrOb", "nrRe", "lmer")
	#	
	if (showData) {
		print(fmddf)
		cat("\n")
	} # end if
	if (showICC) {
		print(out)
 		cat("\n\n")
	} # end if
	return(invisible(out))
} # EOF

calcICC <- function(reshDf, useLmer = TRUE) { # ICC2 or ICC3
	type <- reshDf@type
	onlyData <- reshDf[, -c(1:2), drop=FALSE] # kick out Patient_ID and MuscleTest_ID
	odul <- length(unique(as.numeric(as.matrix(onlyData, nrow=1)))) # only data unique length. odul. yea.
	icRes <- psych::ICC(onlyData, lmer=useLmer)
	res <- icRes$results
	ind <- which(res$type == type)  # type from the first column of the results data frame
	icc <- res[ind, "ICC"]
	pVal <- res[ind, "p"]
	nrJ <- icRes$n.judge
	nrObs <- icRes$n.obs
	if  (odul == 1) {
		icc <- 1
		pVal <- 0
	}
	return(list(type=type, icc=icc, pVal=pVal, nrJ=nrJ, nrObs=nrObs))
} # EOF

assignPValChar <- function(pv) {
	if (is.nan(pv)) {return("NaN")}
	if (pv >= 0.1 ) {return("")}
	if (pv >= 0.05 & pv < 0.1 ) {return(".")}
	if (pv >= 0.01 & pv < 0.05 ) {return("*")}
	if (pv >= 0.001 & pv < 0.01 ) {return("**")}
	if (pv < 0.001 ) {return("***")}
	} # EOF

assignIccWord <- function(icc) {
	if (is.nan(icc)) {return("NaN")}
	if (icc < 0.5 ) {return("poor")}
	if (icc >= 0.5 & icc < 0.75 ) {return("moderate")}
	if (icc >= 0.75 & icc < 0.9 ) {return("good")}
	if (icc >= 0.9 ) {return("excellent")}
} # EOF

makeSegment <- function(fmddf, reshType, useLmer=TRUE) {
	rndPvalue <- 4
	rndICC <- 3
	#
	if (reshType == "inter") {
		reshDf <- resh_Inter(fmddf)
	} else { # so it must be intra
		reshDf <- resh_Intra(fmddf)
	} # end else 
	
	ther <- reshDf@therapist # comes already sorted
	iccType <- reshDf@type
	nrReps <- reshDf@nrReps
	pati <- sort(unique(reshDf$Patient_ID))
	muTe <- sort(unique(reshDf$MuscleTest_ID))
	#
	iccList <- calcICC(reshDf, useLmer)
	icc <- round(iccList$icc, rndICC)
	pVal <- round(iccList$pVal, rndPvalue)
	pValChar <- assignPValChar(pVal)
	nrJ <- iccList$nrJ
	nrObs <- iccList$nrObs
	#
	iccWord <- assignIccWord(icc)
	#
	if (grepl("ICC3", iccType)) { # so we have INTRA
		nrJ <- 1
	} # end if
	#
	if (identical(muTe, fmdenqEnv$allMuscleTestIDs)) { # has been assigned during data import
		muTeChar <- "all"
	} else {
		muTeChar <- paste(muTe, collapse=", ")
	} # end else
	#
	if (identical(ther, fmdenqEnv$allTherapistIDs)) { # has been assigned during data import
		therChar <- "all"
	} else {
		therChar <- paste(ther, collapse=", ")
	} # end else
	#
	if (identical(pati, fmdenqEnv$allPatientIDs)) { # has been assigned during data import
		patiChar <- "all"
	} else {
		patiChar <- paste(pati, collapse=", ")
	} # end else
	#
	out <- data.frame(therChar, patiChar, muTeChar, icc, iccWord, pValChar, pVal, iccType, nrJ, nrObs, nrReps, useLmer)
	colnames(out) <- c("Therapist", "Patient", "MuscleTest", "ICC", "Rel.", "sig.", "p-value", "Type", "nrJ", "nrOb", "nrRe", "lmer")
	return(out)
} # EOF

# for INTRA class, --> intra-rater   split by therapist
intraRater <- function(fmddf, exclOneMT = FALSE, inclOneMT=TRUE, lmer=TRUE) {
	#
	fnTher <- function(datf, lmer) {
		plyr::ddply(datf, .variables=c("Therapist_ID"), .fun=makeSegment, reshType="intra", useLmer=lmer)[,-1]
	} # EOIF
	fnTherPat <- function(datf, lmer) {
		plyr::ddply(datf, .variables=c("Therapist_ID", "Patient_ID"), .fun=makeSegment, reshType="intra", useLmer=lmer)[,-(1:2)]
	} # EOIF
	#
	fnTherSiMu <- function(datf, lmer) {
		plyr::ddply(datf, .variables=c("Therapist_ID", "MuscleTest_ID"), .fun=makeSegment, reshType="intra", useLmer=lmer)[,-(1:2)]
	} # EOIF
	#
	
	## split by therapist; by therapist and patient ##
	dfTher <- fnTher(fmddf, lmer)
	dfTher$Split <- "Th"
	dfTherPat <- fnTherPat(fmddf, lmer)
	dfTherPat$Split <- "ThPa"
	#
	
	## split by therapist and include only one muscle test
	dfTherSiMu <- NULL
	if (inclOneMT) {
		dfTherSiMu <- fnTherSiMu(fmddf, lmer)
		dfTherSiMu$Split <- "ThSiMu"
	} # end if inclOneMT
	#
	
	## Exclude one muscle test only ##
	dfTherMu <- dfTherPatMu <- NULL
	if (exclOneMT) {
		muTe <- sort(unique(fmddf$MuscleTest_ID))
		for (i in 1:length(muTe)) {
			fmdSel <- ssc(fmddf, "MuscleTest_ID", muTe[i], include=FALSE)
			dfTherMu <- rbind(dfTherMu, fnTher(fmdSel, lmer))
			dfTherPatMu <- rbind(dfTherPatMu, fnTherPat(fmdSel, lmer))
		} # end for i
		dfTherMu$Split <- "ThMuex"
		dfTherPatMu$Split <- "ThPaMuex"
	} # end if (exclOneMT)
	#
	
	out <- rbind(dfTher, dfTherPat, dfTherSiMu, dfTherMu, dfTherPatMu)
	return(out)
} # EOF

# ther.sub  list, can contain the values of Therapist_ID to be used for sub-grouping (one subgroup per list-element)
interRater <- function(fmddf, ther.sub=NULL, exclOneMT=FALSE, inclOneMT=TRUE, lmer=TRUE) {
	#
	fnPat <- function(datf, lmer) {
		plyr::ddply(datf, .variables=c("Patient_ID"), .fun=makeSegment, reshType="inter", useLmer=lmer)[,-1]
	} # EOIF
	#
	fnMuscle <- function(datf, lmer) {
		plyr::ddply(datf, .variables=c("MuscleTest_ID"), .fun=makeSegment, reshType="inter", useLmer=lmer)[,-1]
	} # EOIF
	#
	
	## first all in ##
	dfAll <- makeSegment(fmddf, "inter", useLmer=lmer)
	dfAll$Split <- "No"
	#
	dfPat <- fnPat(fmddf, lmer)
	dfPat$Split <- "Pa"
	#

	## Therapist Subgroups ##
	dfAllTherSub <-  dfPatTherSub <- NULL
	if (!is.null(ther.sub)) {
		# XXX to do: check if all the list elements are valid values present in the dataset
		for (i in 1: length(ther.sub)) {
			vals <- ther.sub[[i]]
			indCol <- NULL
			for (k in 1: length(vals)) {
				aa <- which(fmddf$Therapist_ID == vals[k])
				indCol <- c(indCol, aa) # collect together all the indices to keep
			} # end for k (going through the values)
			fmdSel <- fmddf[indCol,]
			dfAllTherSub <- rbind(dfAllTherSub, makeSegment(fmdSel, "inter", useLmer=lmer))
			dfPatTherSub <- rbind(dfPatTherSub, fnPat(fmdSel, lmer))
		} # end for i (going through the list elements)
		dfAllTherSub$Split <- "Thsg"
		dfPatTherSub$Split <- "PaThsg"
	} # end !is.null (ther.sub)
	#
	
	## Include only one muslce test ##		# i.e. group by MuscleTest_ID 
	# all in, only grouped by single muscle test
	dfAllSiMu <- NULL
	if (inclOneMT) {
		dfAllSiMu <- fnMuscle(fmddf, lmer)
		dfAllSiMu$Split <- "SiMu"
	} # end if inclOneMT
	#
	
	## Combination of single muscle test and therapist subgrouping  ##
	dfTherSubSiMu <- NULL
	if (!is.null(ther.sub) & inclOneMT) {
		# XXX to do: check if all the list elements are valid values present in the dataset
		for (i in 1: length(ther.sub)) {
			vals <- ther.sub[[i]]
			indCol <- NULL
			for (k in 1: length(vals)) {
				aa <- which(fmddf$Therapist_ID == vals[k])
				indCol <- c(indCol, aa) # collect together all the indices to keep
			} # end for k (going through the values)
			fmdSel <- fmddf[indCol,] # a subselection by therapist subgroup
			# currently no additional subgrouping by patient or patient subgroups, --> not enougth data, this better to do via an extra function focusing on single muscle tests
			dfTherSubSiMu <- rbind(dfTherSubSiMu, fnMuscle(fmdSel, lmer))
		} # end for i (going through the ther.sub list elements)
		dfTherSubSiMu$Split <- "ThsgSiMu"
	} # end if (!is.null(ther.sub) & inclOneMT)		
	#

	## Exclude one muscle test ##
	dfAllMu <- dfPatMu <- NULL
	if (exclOneMT) {
		muTe <- sort(unique(fmddf$MuscleTest_ID))
		for (i in 1:length(muTe)) {
			fmdSel <- ssc(fmddf, "MuscleTest_ID", muTe[i], include=FALSE)
			dfAllMu <- rbind(dfAllMu, makeSegment(fmdSel, "inter", useLmer=lmer))
			dfPatMu <- rbind(dfPatMu, fnPat(fmdSel, lmer))
		} # end for i
		dfAllMu$Split <- "Muex"
		dfPatMu$Split <- "PaMuex"
	} # end if (exclOneMT)
	#

	## Combination of therapist subgroup and exclusion of one muscle test ##
	dfAllTherSubMu <- dfPatTherSubMu <- NULL
	if (!is.null(ther.sub) & exclOneMT) {
		# XXX to do: check if all the list elements are valid values present in the dataset
		for (i in 1: length(ther.sub)) {
			vals <- ther.sub[[i]]
			indCol <- NULL
			for (k in 1: length(vals)) {
				aa <- which(fmddf$Therapist_ID == vals[k])
				indCol <- c(indCol, aa) # collect together all the indices to keep
			} # end for k (going through the values)
			fmdSel <- fmddf[indCol,]
			muTe <- sort(unique(fmdSel$MuscleTest_ID))
			for (m in 1:length(muTe)) {
				fmdSel2 <- ssc(fmdSel, "MuscleTest_ID", muTe[m], include=FALSE)
				dfAllTherSubMu <- rbind(dfAllTherSubMu, makeSegment(fmdSel2, "inter", useLmer=lmer))
				dfPatTherSubMu <- rbind(dfPatTherSubMu, fnPat(fmdSel2, lmer))
			} # end for m (going through the muscle tests)
		} # end for i (going through the ther.sub list elements)
		dfAllTherSubMu$Split <- "ThsgMuex"
		dfPatTherSubMu$Split <- "PaThsgMuex"
	} # end if (!is.null(ther.sub) & exclOneMT)
	#
	out <- rbind(dfAll, dfPat, dfAllTherSub, dfPatTherSub, dfAllSiMu, dfTherSubSiMu, dfAllMu, dfPatMu, dfAllTherSubMu, dfPatTherSubMu)
	return(out)
} # EOF

calcBottomBorders <- function(resDf, sheetName) {
	ThsgChar <- "Thsg"
	#
	rowIndsInter <- rowIndsIntra <- NULL
	#
	if (sheetName == "Inter") {
		## look for therapist subgroup boundaries ##
		ind <- which(grepl(ThsgChar, resDf$Split))
		uniTher <- unique(resDf$Therapist[ind])
		rl <- rle(resDf$Therapist)
		vals <- rl$values
		lengs <- rl$lengths
		csLengs <- cumsum(lengs)
		for (i in 1: (length(uniTher)-1) ) {
			ind <- which(vals == uniTher[i]) # more than one maybe
			for (k in 1: length(ind)) {
				rowIndsInter <- c(rowIndsInter, csLengs[ind[k]] ) ## hahaha . cumsum does the trick!
			} # end for k
		} # end for i
	} # end if sheet Name == "Inter"
	#
	if (sheetName == "Intra") { # simply run rle over Therapist
		uniTher <- unique(resDf$Therapist)
		rl <- rle(resDf$Therapist)
		vals <- rl$values
		lengs <- rl$lengths
		csLengs <- cumsum(lengs)
		for (i in 1: (length(uniTher)) ) {
			ind <- which(vals == uniTher[i]) # more than one maybe
			for (k in 1: length(ind)) {
				rowIndsIntra <- c(rowIndsIntra, csLengs[ind[k]] ) ## hahaha . cumsum does the trick!
			} # end for k
		} # end for i	
	} # end if (sheetName == "Intra")
	
	return(c(rowIndsInter, rowIndsIntra))	
} # EOF

addWorkbookStyles <- function(wb, resDf, sheetName) { 
	grayCol <- "gray46"
	#
	uniSplit <- unique(resDf$Split)
	nrCols <- aaa <- length(uniSplit)
	if (nrCols > 9) {
		singlePool <- RColorBrewer::brewer.pal(9, "Pastel1")
		pastColsPool <- rep(singlePool, 20) # gives 180 (repeating) colors. Should be enough.
	} else {
		if (nrCols < 3) {aaa <- 3}
		pastColsPool <- RColorBrewer::brewer.pal(aaa, "Pastel1") # we have to take minium 3 (do not know why.)
	} # end else
	#
	for (i in 1: length(uniSplit)) {
    	indRows <- which(resDf$Split == uniSplit[i])
    	fgFillStyle <- openxlsx::createStyle(fgFill=pastColsPool[i])
		openxlsx::addStyle(wb, sheet=sheetName, fgFillStyle, rows=indRows+1, cols = 1:ncol(resDf), gridExpand=TRUE, stack=TRUE)    	
	} # end for i 
	#
	headerStyle <- openxlsx::createStyle(textDecoration="bold")
	openxlsx::addStyle(wb, sheet=sheetName, headerStyle, rows=1, cols = 1:ncol(resDf), gridExpand=TRUE, stack=TRUE)
	#
	grayTextStyle <- openxlsx::createStyle(fontColour=grayCol)
	openxlsx::addStyle(wb, sheet=sheetName, grayTextStyle, rows=(1:nrow(resDf))+1, cols = 8:ncol(resDf), gridExpand=TRUE, stack=TRUE)
	#
	borderStyle <- openxlsx::createStyle(border="bottom", borderColour=grayCol)
	openxlsx::addStyle(wb, sheet=sheetName, borderStyle, rows=calcBottomBorders(resDf, sheetName)+1, cols = 1:ncol(resDf), gridExpand=TRUE, stack=TRUE)
	#
	openxlsx::setColWidths(wb, sheet=sheetName, cols = 1:ncol(resDf), widths = "auto")
	#	
	return(wb)
} # EOF

addWBInfoStyles <- function(wb) {
	headerStyle <- openxlsx::createStyle(textDecoration="bold")
	openxlsx::addStyle(wb, sheet="Info", headerStyle, rows=1, cols = 1:2, gridExpand=TRUE, stack=TRUE)
	openxlsx::setColWidths(wb, sheet="Info", cols = 1:2, widths = "auto")
	return(wb)
} # EOF

createInfoDf <- function() {
	allMtId <- fmdenqEnv$allMuscleTestIDs
	allTherId <- fmdenqEnv$allTherapistIDs
	allPatId <- fmdenqEnv$allPatientIDs
	#
	outDf <- data.frame(c("Patient_ID", "Therapist_ID", "MuscleTest_ID"), c(paste(allPatId, collapse=", "), paste(allTherId, collapse=", "), paste(allMtId, collapse=", ") ))
	colnames(outDf) <- c("Column Name", "All Values")
	return(outDf)
} # EOF

generateSaveWorkbook <- function(intraDf, interDf, fnOut, verbose) {
	exportFolderName <- "results_stat"
	wsZoom <- 140
	#
	wb <- openxlsx::createWorkbook("IntraInter")
	openxlsx::addWorksheet(wb, sheetName="Intra", zoom=wsZoom)
	openxlsx::addWorksheet(wb, sheetName="Inter", zoom=wsZoom)
	openxlsx::addWorksheet(wb, sheetName="Info", zoom=wsZoom)		
	#
	openxlsx::writeData(wb, sheet="Intra", intraDf)
	openxlsx::writeData(wb, sheet="Inter", interDf)
	openxlsx::writeData(wb, sheet="Info", createInfoDf())
	#
	wb <- addWorkbookStyles(wb, intraDf, "Intra")
	wb <- addWorkbookStyles(wb, interDf, "Inter")
	wb <- addWBInfoStyles(wb)
	#
	filename <-  paste0(exportFolderName, "/", fnOut, ".xlsx")
	openxlsx::saveWorkbook(wb, filename, overwrite = TRUE)
	#
	if (verbose) {
		cat(paste0("Statistics results saved under '", basename(filename), "'.\n"))
	} # end if verbose
} # EOF

#' @title Calculate Intra- and Interclass Correlation Coefficients
#' @description Takes the dataset as provided in argument 'fmddf' and calculates
#' intra- and interclass correlation coefficients for some pre-defined and some
#' customisable groupings / data-splits.
#' @details For calculating the resp. ICC, function \code{\link[psych]{ICC}} from
#' package \code{\link{psych}} is used.
#' @param fmddf A dataframe containing FMD data.
#' @param exclOneMT Logical, if additional groupings with one muscle test
#' excluded should be performed. Defaults to FALSE.
#' @param inclOneMT Logical, if additional groupings with only one muscle 
#' test included should be performed. Defaults to TRUE.
#' @param ther.sub NULL or List. If left at the default NULL, no additional
#' splitting with therapist subgroups is performed. Provide a list with
#' n therapist IDs in each list element to calculate additional inter-rater
#' correlation coefficients on data split by therapist subgroups defined in
#' each list element; see examples.
#' @param lmer Logical. Should the \code{\link[lme4]{lmer}} function from
#' packakge \code{\link[lme4]{lme4}} be used for calculating the
#' \code{\link[psych]{ICC}} via package \code{\link{psych}}.
#' @param verbose Logical, if status info should be given; defaults to TRUE.
#' @param fnOut Character length one. The filename (without the ending '.xlsx')
#' of the xlsx file where the statistics results will be saved to.
#' @param toXls Logical, if statistics results should be saved to xlsx.
#' Defaults to TRUE.
#' @return A list length two, with a dataframe holding the statistics
#' results for the intra-class correlation coefficients in the first list
#' element, and the inter-class correlation coefficients in the second.
#' @examples
#' \dontrun{
#' fmdf <- importData()
#' res <- calcIntraInter(fmdf) # run with everything left at defaults
#' # additionally split one excluded muscle test, save under different name
#' res <- calcIntraInter(fmdf, exclOneMT=TRUE, fnOut="results_stat_MTex")
#' # define two therapist subgroups (in a list length two)
#' tsList <- list(c("Ther#1", "Ther#2", "Ther#3"), c("Ther#4", "Ther#5", "Ther#6"))
#' res <- calcIntraInter(fmdf, exclOneMT=FALSE, ther.sub=tsList, fnOut="results_stat_Thersub")
#' res
#' }
#' @seealso \code{\link{runFmdenq}}
#' @export
calcIntraInter <- function(fmddf, exclOneMT=FALSE, inclOneMT=TRUE, ther.sub=NULL, lmer=TRUE, verbose=TRUE, fnOut="results_stat", toXls=TRUE) {
	if (verbose) {
		cat("Calculating intra-rater statistics ... \n")
	}
	intraDf <- intraRater(fmddf, exclOneMT, inclOneMT, lmer)
	if (verbose) {
		cat("Calculating inter-rater statistics ... \n")
	}
	interDf <- interRater(fmddf, ther.sub, exclOneMT, inclOneMT, lmer)
	#
	# export to xlsx
	if (toXls) {
		generateSaveWorkbook(intraDf, interDf, fnOut, verbose)
	} # end toXls
	#
	return(invisible(list(intra=intraDf, inter=interDf)))
} # EOF

#' @title Import Data and Calculate Intra- and Interclass Correlation Coefficients
#' @description Imports FMD result data from a predefined xlsx file, calculates
#' intra- and interclass correlation coefficients on some predefined and some
#' optional custom data grouoings / splittings, and saves the statistics results to an
#' xlsx file.
#' @details The name of the input-excel file can be specified or it can be
#' interactively selected, the name of the output-excel file can be specified.
#' This basically is a convenient wrapper for the two functions
#' \code{\link{importData}} and \code{\link{calcIntraInter}}; see examples there.
#' @section Calculating ICC: For calculating the resp. ICC, function
#' \code{\link[psych]{ICC}} from package \code{\link{psych}} is used.
#' @inheritParams importData
#' @inheritParams calcIntraInter
#' @param verbose Logical, if various status reports / feedback should be given.
#' Defaults to TRUE.
#' @examples
#' \dontrun{
#' runFmdenq()
#' runFmdenq(selIn=TRUE, fnOut="results_stat_foo") # interactively select xlsx file for data import
#' }
#' @seealso \code{\link{importData}}, \code{\link{calcIntraInter}}
#' @export
runFmdenq <- function(fnIn="results_fmd", selIn=FALSE, exclOneMT=FALSE, inclOneMT=TRUE, ther.sub=NULL, lmer=TRUE, verbose=TRUE, fnOut="results_stat", toXls=TRUE) {
	fmddf <- importData(fnIn, selIn, verbose)
	return(calcIntraInter(fmddf, exclOneMT, inclOneMT, ther.sub, lmer, verbose, fnOut, toXls))
} # EOF
