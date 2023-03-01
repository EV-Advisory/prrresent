#' Read PowerPoint Template
#'
#' This function reads in a PowerPoint template file and returns a pptx object to be used for further modifications.
#'
#' @param folder_location The folder that hosts the powerpoint template
#' @param document_name The name of the PowerPoint template file to be read in.
#'
#' @return A pptx object for further modifications.
#' @export
#'
#' @importFrom here here
#' @importFrom officer read_pptx
#' @examples
#' \dontrun{
#' doc <- read_ppt_template("default-eva-template.potx")
#' }
#'
read_ppt_template <- function(folder_location = "inst/extdata",document_name = "default-eva-template.potx") {

  if(!interactive())doc<-officer::read_pptx(system.file("extdata","default-eva-template.potx",package = "prrresent"))
  if(interactive())doc <- officer::read_pptx(here::here(folder_location, document_name))
  return(doc)
}
