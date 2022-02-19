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
	checkData(fmddf, filename=fnIn, verbose)
	if (class(fmddf) == "data.frame") {
		message("Data import successful.")
	}
	fmddf <- exchangeCharForNumbers(fmddf)
	return(fmddf)
} # EOF

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

# type from the first column of the results data frame
calcICC <- function(reshDf, type, useLmer = TRUE) { # ICC2 or ICC3
#	reshDf <- resh(fmddf)
	icRes <- psych::ICC(reshDf[,-(1:3)], lmer=useLmer)
	res <- icRes$results
	ind <- which(res$type == type)
	icc <- res[ind, "ICC"]
	pVal <- res[ind, "p"]
	nrJ <- icRes$n.judge
	return(list( type=type, icc=icc, pVal=pVal, nrJ=nrJ))
} # EOF

assignPValChar <- function(pv) {
	if (is.nan(pv)) {return("NaN")}
	if (pv >= 0.1 ) {return("")}
	if (pv >= 0.05 & pv < 0.1 ) {return(".")}
	if (pv >= 0.01 & pv < 0.05 ) {return("*")}
	if (pv >= 0.001 & pv < 0.01 ) {return("**")}
	if (pv < 0.001 ) {return("***")}
	} # EOF

makeSegment <- function(fmddf, iccType, rnd=3, useLmer=TRUE) {
	reshDf <- resh(fmddf)
	ther <- sort(unique(reshDf$Therapist_ID))
	pati <- sort(unique(reshDf$Patient_ID))
	muTe <- sort(unique(reshDf$MuscleTest_ID))
	#
	iccList <- calcICC(reshDf, iccType, useLmer)
	icc <- signif(iccList$icc, rnd)
	pVal <- signif(iccList$pVal, rnd)
	pValChar <- assignPValChar(pVal)
	nrJ <- iccList$nrJ
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
	out <- data.frame(therChar, patiChar, muTeChar, icc, pValChar, pVal, iccType, nrJ, useLmer)
	colnames(out) <- c("Therapist", "Patient", "MuscleTest", "ICC", "sig.", "p-value", "Type", "nrJ", "lmer")
	return(out)
} # EOF

# for INTRA class, --> intra-rater   split by therapist

intraRater <- function(fmddf, exclOneMT = FALSE, lmer=TRUE) {
	#
	fnTher <- function(datf, lmer) {
		plyr::ddply(datf, .variables=c("Therapist_ID"), .fun=makeSegment, iccType="ICC3", useLmer=lmer)[,-1]
	} # EOIF
	fnTherPat <- function(datf, lmer) {
		plyr::ddply(datf, .variables=c("Therapist_ID", "Patient_ID"), .fun=makeSegment, iccType="ICC3", useLmer=lmer)[,-(1:2)]
	} # EOIF
	#
	dfTher <- fnTher(fmddf, lmer)
	dfTher$Split <- "Th"
	dfTherPat <- fnTherPat(fmddf, lmer)
	dfTherPat$Split <- "ThPa"
	#
	dfTherMu <- dfTherPatMu <- NULL
	if (exclOneMT) {
		muTe <- sort(unique(fmddf$MuscleTest_ID))
		for (i in 1:length(muTe)) {
			fmdSel <- ssc(fmddf, "MuscleTest_ID", muTe[i], include=FALSE)
			aa <- fnTher(fmdSel, lmer)
			dfTherMu <- rbind(dfTherMu, aa)
			bb <- fnTherPat(fmdSel, lmer)
			dfTherPatMu <- rbind(dfTherPatMu, bb)
		} # end for i
		dfTherMu$Split <- "ThMu"
		dfTherPatMu$Split <- "ThPaMu"
	} # end if (exclOneMT)
	#
	out <- rbind(dfTher, dfTherPat, dfTherMu, dfTherPatMu)
	return(out)
} # EOF

# ther.sub  list, can contain the values of Therapist_ID to be used for sub-grouping (one subgroup per list-element)
interRater <- function(fmddf, ther.sub=NULL, exclOneMT=FALSE, lmer=TRUE) {

	## first all in ##
	dfAll <- makeSegment(fmddf, "ICC2", useLmer=lmer)
	dfAll$Split <- "No"
	#
	fnPat <- function(datf, lmer) {
		plyr::ddply(datf, .variables=c("Patient_ID"), .fun=makeSegment, iccType="ICC2", useLmer=lmer)[,-1]
	} # EOIF
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
			aa <- makeSegment(fmdSel, "ICC2", useLmer=lmer)
			dfAllTherSub <- rbind(dfAllTherSub, aa)
			bb <- fnPat(fmdSel, lmer)
			dfPatTherSub <- rbind(dfPatTherSub, bb)
		} # end for i (going through the list elements)
		dfAllTherSub$Split <- "Thsg"
		dfPatTherSub$Split <- "PaThsg"
	} # end !is.null (ther.sub)
	#

	## Exclude one muscle test ##
	dfAllMu <- dfPatMu <- NULL
	if (exclOneMT) {
		muTe <- sort(unique(fmddf$MuscleTest_ID))
		for (i in 1:length(muTe)) {
			fmdSel <- ssc(fmddf, "MuscleTest_ID", muTe[i], include=FALSE)
			aa <- makeSegment(fmdSel, "ICC2", useLmer=lmer)
			dfAllMu <- rbind(dfAllMu, aa)
			bb <- fnPat(fmdSel, lmer)
			dfPatMu <- rbind(dfPatMu, bb)
		} # end for i
		dfAllMu$Split <- "Mu"
		dfPatMu$Split <- "PaMu"
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
				aa <- makeSegment(fmdSel2, "ICC2", useLmer=lmer)
				dfAllTherSubMu <- rbind(dfAllTherSubMu, aa)
				bb <- fnPat(fmdSel2, lmer)
				dfPatTherSubMu <- rbind(dfPatTherSubMu, bb)
			} # end for m (going through the muscle tests)
		} # end for i (going through the ther.sub list elements)
		dfAllTherSubMu$Split <- "ThsgMu"
		dfPatTherSubMu$Split <- "PaThsgMu"
	} # end if (!is.null(ther.sub) & exclOneMT)
	#
	out <- rbind(dfAll, dfPat, dfAllTherSub, dfPatTherSub, dfAllMu, dfPatMu, dfAllTherSubMu, dfPatTherSubMu)
	return(out)
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
calcIntraInter <- function(fmddf, exclOneMT=FALSE, ther.sub=NULL, lmer=TRUE, verbose=TRUE, fnOut="results_stat", toXls=TRUE) {
	exportFolderName <- "results_stat"
	#
	if (verbose) {
		cat("Calculating intra-rater statistics ... \n")
	}
	intraDf <- intraRater(fmddf, exclOneMT, lmer)
	if (verbose) {
		cat("Calculating inter-rater statistics ... \n")
	}
	interDf <- interRater(fmddf, ther.sub, exclOneMT, lmer)
	#
	# export to xlsx
	if (toXls) {
		wb <- openxlsx::createWorkbook("IntraInter")
		openxlsx::addWorksheet(wb, sheetName="Intra")
		openxlsx::addWorksheet(wb, sheetName="Inter")
		openxlsx::writeData(wb, sheet="Intra", intraDf)
		openxlsx::writeData(wb, sheet="Inter", interDf)
		filename <-  paste0(exportFolderName, "/", fnOut, ".xlsx")
		openxlsx::saveWorkbook(wb, filename, overwrite = TRUE)
		if (verbose) {
 			cat(paste0("Statistics results saved under '", basename(filename), "'.\n"))
		}
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
runFmdenq <- function(fnIn="results_fmd", selIn=FALSE, exclOneMT=FALSE, ther.sub=NULL, lmer=TRUE, verbose=TRUE, fnOut="results_stat", toXls=TRUE) {
	fmddf <- importData(fnIn, selIn, verbose)
	return(calcIntraInter(fmddf, exclOneMT, ther.sub, lmer, verbose, fnOut, toXls))
} # EOF
