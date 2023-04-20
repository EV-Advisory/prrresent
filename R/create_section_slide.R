#' Create Section Slide
#'
#' This function creates a title slide in a PowerPoint template. It adds a slide with the selected layout and adds a title and subtitle to the slide.
#'
#' @param x The PowerPoint document to add the slide to.
#' @param slide_title The title to be added to the slide.
#' @param slide_subtitle The subtitle to be added to the slide.
#' @param layout_name The name of the slide layout to be used. Default is "Section - Black".
#' @param master_name The name of the slide master to be used. Default is "EVA - Standard".
#' @param logo The location of a logo file for a client presentation. If NULL, ignore
#' @return A PowerPoint slide with the selected layout and added title, subtitle and picture.
#' @export
#'
#' @examples
#' \dontrun{
#' create_title_slide(doc,
#' "Title",
#' "Subtitle",
#' "Section - Blue",
#' "EVA - Standard")
#' }
create_section_slide <- function(x,
                                 slide_title,
                                 slide_subtitle,
                                 layout_name = "Section - Black",
                                 master_name = 'EVA - Standard',
                                 logo = NULL) {
  slide_layout <- layout_name
  stopifnot('Layout is not a section slide' = grepl("Section ", layout_name))
  officer::add_slide(x, layout = slide_layout, master = master_name) %>%
    officer::ph_with(value = slide_title,
                     location = officer::ph_location_label(ph_label = "Title")) %>%
    officer::ph_with(value = slide_subtitle,
                     location = officer::ph_location_label(ph_label = "Subtitle"))# %>%
    # {
    #   ifelse(is.null(logo), ., {
    #     officer::ph_with(
    #       value = slide_subtitle,
    #       value = officer::external_img(logo),
    #       location = officer::ph_location_label(ph_label = )
    #     )
    #   })
    # }
}
