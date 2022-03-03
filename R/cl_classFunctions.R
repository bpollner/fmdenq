


show_fmddf_reshInter <- function(object) {
	datRnd <- round(object[,-c(1:2)], 1)
	showDf <- cbind(object[,c(1:2)], datRnd)
	#
	txtIntro <- "Object of class 'fmddf_reshInter'\n"
	txtType <- paste0("Slot 'type':\n", object@type, "\n")
	txtNrReps <- paste0("Slot 'nrReps'\n", object@nrReps, "\n")
	#
	cat(txtIntro)
	print(showDf)
	cat("\n")
	cat(txtType)
	cat("\n")
	cat(txtNrReps)
	cat("\n")
} # EOF
