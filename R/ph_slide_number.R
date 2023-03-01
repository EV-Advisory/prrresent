
#' Add slide numbers to PowerPoint slides
#'
#' This function adds slide numbers to a PowerPoint slide. It retrieves the maximum slide ID using pptx_summary and adds it to the slide as a text box.
#'
#' @param x A PowerPoint document.
#' @importFrom officer pptx_summary
#' @return A PowerPoint slide with a slide number.
#'
#' @examples
#' \dontrun{
#' ph_slide_number(my_presentation)
#' }
ph_slide_number <- function(x) {
  slide_summary <- pptx_summary(x)

  officer::ph_with(
    x = x,
    value = max(slide_summary$slide_id),
    location = officer::ph_location_label(ph_label = "Slide Number")
  )
}
