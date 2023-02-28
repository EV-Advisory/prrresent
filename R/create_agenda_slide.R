#' Create an Agenda Slide
#'
#' This function creates an agenda slide in a PowerPoint presentation using the officer package in R.
#'
#' @param x a PowerPoint presentation object.
#' @param slide_title Title of the slide.
#' @param slide_body Content of the slide, created using the officer::unordered_list function. Default list has 3 levels and is formatted with three different colors (red, blue, and orange).
#' @param section_title Title of the section.
#' @param footer_text Text to be displayed in the footer.
#' @param layout_name Name of the layout used for the slide. Default is "Agenda".
#' @param master_name Name of the master layout used for the slide. Default is "EVA - Standard".
#'
#' @return A PowerPoint slide object with the specified title, content, section title, and footer text.
#'
#' @import officer
#' @importFrom officer layout_summary add_slide ph_with ph_location_label
#' @export
#'
#' @examples
#' \dontrun{
#' ppt <- read_pptx()
#' create_agenda_slide(ppt, "Agenda", "Welcome", section_title = "Introduction")
#' }
create_agenda_slide <- function(x,
                                slide_title,
                                slide_body = officer::unordered_list(
                                  level_list = c(1, 2, 3),
                                  str_list = c("Level1", "Level2", "Level1"),
                                  style = list(
                                    fp_text(color = "#C32900", font.size = 0),
                                    fp_text(color = "#006699", font.size = 0),
                                    fp_text(color = "#f2af00", font.size = 0)
                                  )
                                ),
                                section_title = "Agenda",
                                footer_text = "Client Engagement",
                                layout_name = "Agenda",
                                master_name = 'EVA - Standard') {
  slide_layout <- layout_name
  stopifnot('Layout is not an agenda slide' = grepl("Agenda", layout_name))
  stopifnot("Layout doesn't exist" = layout_name %in% layout_summary(x)$layout)
  officer::add_slide(x, layout = slide_layout, master = master_name) %>%
    officer::ph_with(value = slide_title,
                     location = officer::ph_location_label(ph_label = "Title")) %>%
    officer::ph_with(value = slide_body,
                     location = officer::ph_location_label(ph_label = "Text")) %>%
    officer::ph_with(
      value = section_title,
      location = officer::ph_location_label(ph_label = "Section Title")
    ) %>%
    officer::ph_with(value = footer_text,
                     location = officer::ph_location_label(ph_label = "Footer")) %>%
    ph_disclaimer(x = .) %>%
    ph_slide_number(x = .)
}
