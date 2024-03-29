% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/create_objectives_slide.R
\name{create_objectives_slide}
\alias{create_objectives_slide}
\title{Create Objectives Slide}
\usage{
create_objectives_slide(
  x,
  slide_title = " ",
  slide_subtitle = " ",
  slide_table,
  slide_section = "",
  footer_text = "",
  layout_name = "Objectives",
  master_name = "EVA - Standard",
  autofit = F
)
}
\arguments{
\item{x}{The PowerPoint document to add the slide to.}

\item{slide_title}{The title to be added to the slide. Default is "Default Title for OKRs".}

\item{slide_subtitle}{The subtitle to be added to the slide. Default is "".}

\item{slide_table}{The table to be added to the slide.}

\item{slide_section}{The section to be added to the slide.}

\item{footer_text}{The footer text of the slide}

\item{layout_name}{The name of the slide layout to be used. Default is "Objectives".}

\item{master_name}{The name of the slide master to be used. Default is "EVA - Standard".}

\item{autofit}{A boolean to specify if the table should be autofitted. Default is F.}
}
\value{
A PowerPoint slide with the selected layout and added title, table, caption, body text and section.
}
\description{
This function creates a slide in a PowerPoint template containing a table about the objectives of the project. It adds a slide with the selected layout and adds a title, table, caption, body text and section to the slide.
}
\examples{
\dontrun{
create_objectives_slide(x,
"Table Title",
"Table Caption",
my_table,
"Some text",
"Section A",
"Objectives",
"EVA - Standard",
F)
}
}
