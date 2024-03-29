% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/create_title_slide.R
\name{create_title_slide}
\alias{create_title_slide}
\title{Create Title Slide}
\usage{
create_title_slide(
  x,
  slide_title,
  slide_subtitle,
  picture,
  layout_name = "Title - Black",
  master_name = "EVA - Standard"
)
}
\arguments{
\item{x}{The PowerPoint document to add the slide to.}

\item{slide_title}{The title to be added to the slide.}

\item{slide_subtitle}{The subtitle to be added to the slide.}

\item{picture}{The path to the picture to be added to the slide.}

\item{layout_name}{The name of the slide layout to be used. Default is "Title - Black".}

\item{master_name}{The name of the slide master to be used. Default is "EVA - Standard".}
}
\value{
A PowerPoint slide with the selected layout and added title, subtitle and picture.
}
\description{
This function creates a title slide in a PowerPoint template. It adds
a slide with the selected layout and adds a title, subtitle and picture
to the slide.
}
\examples{
\dontrun{
create_title_slide(x, "Title", "Subtitle", "picture.png", "Title - Blue", "EVA - Standard")
}
}
