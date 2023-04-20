#' Create Point Slide
#'
#' This function creates a point slide in a PowerPoint template. It adds
#' a slide with the selected layout and adds a title, subtitle and picture
#' to the slide.
#'
#' @param x The PowerPoint document to add the slide to.
#' @param slide_title The title to be added to the slide.
#' @param slide_subtitle The subtitle to be added to the slide.
#' @param slide_body The body text or contents to be added to the slide.
#' @param layout_name The name of the slide layout to be used. Default is "Title - Black".
#' @param master_name The name of the slide master to be used. Default is "EVA - Standard".
#'
#' @return A PowerPoint slide with the selected layout and added title, subtitle and picture.
#' @export
#'
#' @examples
#' \dontrun{
#' create_point_slide(x, "Title", "Subtitle", "Body Text", "Point - Detail Bullet", "EVA - Standard")
#' }
create_point_slide <- function(x,
                               slide_title = "Default Title for Point Slide",
                               slide_subtitle = "Default Subtitle for Point Slide",
                               slide_body = "Body Text",
                               layout_name = "Point - Detail Bullet",
                               master_name = 'EVA - Standard') {
  stopifnot('Layout is not a point slide' = grepl("Point - ", layout_name))
  stopifnot("Layout doesn't exist" = layout_name %in% layout_summary(x)$layout)
  officer::add_slide(x = x,
                     layout = layout_name,
                     master = master_name) %>%
    officer::ph_with(value = slide_title,
                     location = officer::ph_location_label(ph_label = "Title")) %>%
    officer::ph_with(value = slide_subtitle,
                     location = officer::ph_location_label(ph_label = "Subtitle")) %>%
    {
      `if`(
        !is.null(slide_body),
        ph_with(
          x = .,
          value = slide_body,
          location = officer::ph_location_label(ph_label = "Body")
        ),
        .

      )
    } %>%
    ph_disclaimer(x = .) %>%
    ph_slide_number(x = .) %>%
    return(.)
}
