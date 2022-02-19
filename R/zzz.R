

.onLoad <- function(libname, pkgname) { 
	nsp <- "pgk_fmdenq_envs"												## is defining the name on the search path
	if (!any(grepl(nsp, search()))) {attach(what=NULL, name=nsp)}			## create a new entry on the search path if not already there
	assign("fmdenqEnv", new.env(), pos=nsp)									## create a new environment called "fmdenqEnv"
} # EOF



.onUnload <- function(libpath) {
	if (any(grepl("pgk_fmdenq_envs", search()))) {detach(name="pgk_fmdenq_envs")}
} # EOF



