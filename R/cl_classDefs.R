#' @include fmdenq.R
#' @include cl_classFunctions.R


#### classes ####
# setClassUnion(name="dfNull", members = c("data.frame", "NULL"))
#
setClass("fmddf", contains="data.frame")
setClass("fmddf_reshInter", slots=c(type="character", nrReps="numeric"), contains="data.frame")
setClass("fmddf_reshIntra", slots=c(type="character", nrReps="numeric"), contains="data.frame")







#### methods ####


# #' @rdname fdmat-methods
# #' @export
#			setMethod("[", signature(x = "fmddf_resh"), definition = function(x, i) {
#				cns <- colnames(x)
#				out <- as.data.frame(x@.Data)[i,,drop=FALSE]
#				colnames(out) <- cns
#				return(new("fmddf_resh", out))
#			} ) # end set method


# setMethod("show", signature(object = "fmddf_reshInter"), definition = show_fmddf_reshInter)
