#' Create Title Slide
#'
#' This function creates a title slide in a PowerPoint template. It adds
#' a slide with the selected layout and adds a title, subtitle and picture
#' to the slide.
#'
#' @param x The PowerPoint document to add the slide to.
#' @param slide_title The title to be added to the slide.
#' @param slide_subtitle The subtitle to be added to the slide.
#' @param picture The path to the picture to be added to the slide.
#' @param layout_name The name of the slide layout to be used. Default is "Title - Black".
#' @param master_name The name of the slide master to be used. Default is "EVA - Standard".
#'
#' @return A PowerPoint slide with the selected layout and added title, subtitle and picture.
#' @export
#'
#' @examples
#' \dontrun{
#' create_title_slide(x, "Title", "Subtitle", "picture.png", "Title - Blue", "EVA - Standard")
#' }
create_title_slide <- function(x,
                               slide_title,
                               slide_subtitle,
                               picture,
                               layout_name = "Title - Black",
                               master_name = 'EVA - Standard') {
  slide_layout <- layout_name
  stopifnot('Layout is not a title slide' = grepl("Title ", layout_name))
  stopifnot("Layout doesn't exist" = layout_name %in% layout_summary(x)$layout)
  officer::add_slide(x = x, layout = slide_layout, master = master_name) %>%
    officer::ph_with(value = slide_title,
                     location = officer::ph_location_label(ph_label = "Title")) %>%
    officer::ph_with(value = slide_subtitle,
                     location = officer::ph_location_label(ph_label = "Subtitle")) %>%
    {
      `if`(
        !is.null(picture),
        ph_with(
          x = .,
          value = officer::external_img(picture),
          location = officer::ph_location_label(ph_label = "Image Location")
        ),.

      )
    }%>%
    ph_disclaimer(x=.)%>%
    ph_slide_number(x=.)%>%
    return(.)
}
