#' Create Decision Point Slide
#'
#' This function creates a decision point slide in a PowerPoint template. It adds a slide with the selected layout and adds a main header, actions, decision, section, date, attendees, and footer to the slide.
#'
#' @param x The PowerPoint document to add the slide to.
#' @param dp_header The main header text for the decision point slide.
#' @param actions_body The body text for the actions section of the slide. Default is an empty unordered list.
#' @param decision_body The body text for the decision section of the slide. Default is an empty unordered list.
#' @param section_text The section text for the slide. Default is "Section".
#' @param date_text The date to be displayed on the slide. Default is the current date.
#' @param attendees_list The list of attendees to be displayed on the slide. Default is an empty unordered list.
#' @param footer_text The footer text to be displayed on the slide. Default is an empty string.
#' @param layout_name The name of the slide layout to be used. Default is "Decision Point".
#' @param master_name The name of the slide master to be used. Default is "EVA - Standard".
#'
#' @return A PowerPoint slide with the selected layout and added header, actions, decision, section, date, attendees, and footer.
#' @export
#'
#' @importFrom magrittr `%>%`
#' @importFrom officer unordered_list
#' @importFrom officer fp_text
#' @examples
#' \dontrun{
#' create_decision_point_slide(doc,
#' "Main Decision",
#' actions_body,
#' decision_body,
#' "Section1",
#' "2023-02-22",
#' attendees_list,
#' "Footer")
#' }
create_decision_point_slide <- function(x,
                               dp_header = "Primary decision point text",
                               actions_body =officer::unordered_list(
                                 level_list = c(1, 2, 3),
                                 str_list = c("Level1", "Level2", "Level1"),
                                 style = list(
                                   fp_text(color = "#000000", font.size = 0),
                                   fp_text(color = "#000000", font.size = 0),
                                   fp_text(color = "#000000", font.size = 0)
                                 )) ,
                               decision_body = officer::unordered_list(
                                 level_list = c(1, 2, 3),
                                 str_list = c("Level1", "Level2", "Level1"),
                                 style = list(
                                   fp_text(color = "#000000", font.size = 0),
                                   fp_text(color = "#000000", font.size = 0),
                                   fp_text(color = "#000000", font.size = 0)
                                 )) ,
                               section_text = "Section",
                               date_text = Sys.Date(),
                               attendees_list = officer::unordered_list(
                                 level_list = c(1, 1),
                                 str_list = c("Level1", "Level2"),
                                 style = list(
                                   fp_text(color = "#000000", font.size = 0),
                                   fp_text(color = "#000000", font.size = 0)
                                 )) ,
                               footer_text = "",

                               layout_name = "Decision Point",
                               master_name = 'EVA - Standard') {
  #Set the layout
  slide_layout <- layout_name
  #Determine if the correct layout was passed
  stopifnot('Layout is not a decision point slide' = grepl("Decision ", layout_name))
  stopifnot("Layout doesn't exist" = layout_name %in% layout_summary(x)$layout)
  #Create the slide
  officer::add_slide(x = x,
                     layout = slide_layout,
                     master = master_name) %>%
    #Fill in the header, the primary text
    officer::ph_with(
      value = dp_header,
      location = officer::ph_location_label(ph_label = "Decision Point - Main Body")
    ) %>%
    #Actions list
    officer::ph_with(
      value = actions_body,
      location = officer::ph_location_label(ph_label = "Actions - Body")
    )%>%
    #Decision list
    officer::ph_with(
      value = decision_body,
      location = officer::ph_location_label(ph_label = "Decision - Body")
    ) %>%
    #Name of the section
    officer::ph_with(
      value = section_text,
      location = officer::ph_location_label(ph_label = "Section")
    ) %>%
    #Date
    officer::ph_with(
      value = as.character(date_text),
      location = officer::ph_location_label(ph_label = "Date - Body")
    ) %>%
    #List of attendees for the decision
    officer::ph_with(
      value = attendees_list,
      location = officer::ph_location_label(ph_label = "Attendees - Body")
    )%>%
    #Disclaimer for privacy
    ph_disclaimer(x=.)%>%
    #Automatic inclusion of the slide number
    ph_slide_number(x=.)

}
