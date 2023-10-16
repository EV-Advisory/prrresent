#' Create Graphic Slide
#'
#' This function creates a slide in a PowerPoint template containing a graphic.
#' It adds a slide with the selected layout and adds a title, table, caption,
#' body text and section to the slide.
#'
#' @param x The PowerPoint document to add the slide to.
#' @param slide_title The title to be added to the slide. Default is "Default Title for Table".
#' @param slide_caption The caption to be added to the slide. Default is "".
#' @param slide_graphic The graphic to be added to the slide.
#' @param slide_body The body text to be added to the slide. Default is "Some text".
#' @param slide_section The section to be added to the slide.
#' @param footer_text The footer text of the slide
#' @param layout_name The name of the slide layout to be used. Default is "graphic Full".
#' @param master_name The name of the slide master to be used. Default is "EVA - Standard".
#' @importFrom officer layout_properties
#' @importFrom officer layout_summary
#' @return A PowerPoint slide with the selected layout and added title, table, caption, body text and section.
#' @export
#'
#' @examples
#' \dontrun{
#' create_table_slide(x,
#' "Table Title",
#' "Table Caption",
#' my_graphic,
#' "Some text",
#' "Section A",
#' "Graphic Full",
#' "EVA - Standard",
#' F)
#' }

create_graphic_slide <- function(x,
                                 slide_title = "Default Title for Graphic",
                                 slide_caption = "",
                                 slide_graphic,
                                 slide_body = "Some text",
                                 slide_section = "",
                                 footer_text = "",
                                 layout_name = "Graphic Full",
                                 master_name = 'EVA - Standard') {
  # browser()
  slide_layout_properties <-
    layout_properties(x, layout = layout_name)
  stopifnot('Layout is not a graphic slide' = grepl("Graphic ", layout_name))
  stopifnot("Layout doesn't exist" = layout_name %in% layout_summary(x)$layout)
  if (sum(grepl("*.png|*.jpeg|*.jpg|*.svg", slide_graphic)) > 0)
    slide_graphic <- officer::external_img(slide_graphic)
  officer::add_slide(x, layout = layout_name, master = master_name) %>%
    {
      `if`(
        !grepl("No Title", layout_name),
        officer::ph_with(
          x = .,
          value = slide_title,
          location = officer::ph_location_label(ph_label = "Title")
        ),
        .
      )
    } %>%
    {
      `if`(
        grepl("Graphic Full", layout_name),
        officer::ph_with(
          x = .,
          value = slide_caption,
          location = officer::ph_location_label(ph_label = "Caption")
        ),
        .
      )
    } %>%
    {
      `if`(
        !grepl("Graphic Full", layout_name),
        officer::ph_with(
          x = .,
          value = slide_body,
          location = officer::ph_location_label(ph_label = "Body Text")
        ),
        .
      )
    } %>%
    {
      `if`(
        !grepl("Graphic Full", layout_name),
        officer::ph_with(
          x = .,
          value = slide_graphic,
          location = officer::ph_location_label(ph_label = "Content")
        ),
        officer::ph_with(
          x = .,
          value = slide_graphic,
          location = officer::ph_location_label(ph_label = "Body")
        )
      )
    }  %>%
    officer::ph_with(value = slide_section,
                     location = officer::ph_location_label(ph_label = "Section")) %>% {
                       #Add the disclaimer, if it exists
                       `if`("Disclaimer" %in% slide_layout_properties$ph_label,
                            ph_disclaimer(x = .),
                            .)
                     } %>% {
                       #Add the slide number, if it exists
                       `if`("Slide Number" %in% slide_layout_properties$ph_label,
                            ph_slide_number(x = .),
                            .)
                     } %>% {
                       #Add the footer, if it exists
                       `if`(
                         "Footer" %in% slide_layout_properties$ph_label,
                         officer::ph_with(
                           value = footer_text,
                           location = officer::ph_location_label(ph_label = slide_layout_properties$ph_label[grepl("Footer", slide_layout_properties$ph_label)])
                         ),
                         .
                       )
                     } %>%
    return(.)

}
