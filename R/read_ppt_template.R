#' Read PowerPoint Template
#'
#' This function reads in a PowerPoint template file and returns a pptx object to be used for further modifications.
#'
#' @param template_path The complete location of the powerpoint template.
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
read_ppt_template <-
  function(template_path = here::here("inst/extdata", "default-eva-template.potx")) {
    #Find the template specified that is not the default
    if (!interactive() &
        is.null(template_path))
      doc <-
        officer::read_pptx(template_path)

    #Find the default template of the powerpoint
    if (!interactive())
      doc <-
        officer::read_pptx(system.file("extdata", "default-eva-template.potx", package = "prrresent"))
    #If a different template exists and is specified, find the template in the location provided
    if (interactive())
      doc <- officer::read_pptx(template_path)
    return(doc)
  }
