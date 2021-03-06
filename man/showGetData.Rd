% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/functions.R
\name{showGetData}
\alias{showGetData}
\title{Get and Display a Subset of the Data}
\usage{
showGetData(fmddf, expr = NULL, resh = NULL, showData = TRUE)
}
\arguments{
\item{fmddf}{A data frame containing FMD data.}

\item{expr}{An expression defining the subset of data. If left at the default 
NULL, no subset of the data is created, but he data is merely reshaped.}

\item{resh}{Character length one. What type of reshape should be applied to 
the output. Possible values are 'inter' (default) and 'intra'.}

\item{showData}{Logical If the resulting data frame should be printed. Defaults 
to TRUE.}
}
\value{
A data frame with the subset of data, reshaped for either inter- or 
intra-class analysis.
}
\description{
Use an expression to provide a definition of a subset of the 
data, reshape the data to either inter- or intra-class analysis, possibly 
print these subset and return it.
}
\details{
Use 'colnames(object)' to determine the name of the columns in the 
data frame. (Of course.)
}
\examples{
\dontrun{
fmddf <- importData()
expr <- expression(Patient_ID == "Pat#2")
showGetData(fmddf, expr)
expr <- expression(Patient_ID == "Pat#2" & Therapist_ID != "Ther#4")
showGetData(fmddf, expr)
}
}
