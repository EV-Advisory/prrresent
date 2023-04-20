#' Add a disclaimer to a PowerPoint slide
#'
#' This function adds a disclaimer text to a PowerPoint slide. By default, the disclaimer text is "PRIVATE AND STRICTLY CONFIDENTIAL", but it can be customized to a different text by providing a value for the "val" parameter.
#'
#' @param x A PowerPoint document.
#' @param val The text to use for the disclaimer. Default is "PRIVATE AND STRICTLY CONFIDENTIAL".
#'
#' @return A PowerPoint slide with a disclaimer text box.
#' @importFrom officer ph_with
#' @importFrom officer ph_location_label
#'
#' @examples
#' \dontrun{
#' ph_disclaimer(my_presentation, "This information is for internal use only.")
#' }
ph_disclaimer <- function(x, val = "PRIVATE AND STRICTLY CONFIDENTIAL") {
  officer::ph_with(
    x = x,
    value = val,
    location = officer::ph_location_label(ph_label = "Disclaimer")
  )
}
