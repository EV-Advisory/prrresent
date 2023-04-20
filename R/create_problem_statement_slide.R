#' Create Problem Statement Slide
#'
#' This function creates a problem statement slide in a PowerPoint template. It adds
#' a slide with the selected layout and adds a title, subtitle and body text for
#' additional details to the slide.
#'
#' @param x The PowerPoint document to add the slide to.
#' @param slide_problem_statement The title to be added to the slide.
#' @param slide_subtitle The subtitle to be added to the slide.
#' @param slide_body The body text or contents to be added to the slide.
#' @param layout_name The name of the slide layout to be used. Default is "Problem Statement".
#' @param master_name The name of the slide master to be used. Default is "EVA - Standard".
#'
#' @return A PowerPoint slide with the selected layout and added title, subtitle
#' and picture.
#' @export
#'
#' @examples
#' \dontrun{
#' create_problem_state_slide(x, "Title", "Subtitle", "Body Text",
#' "Problem Statement", "EVA - Standard")
#' }
create_problem_statement_slide <- function(x,
                                           slide_problem_statement = "Insert Problem Statement Here",
                                           slide_subtitle = "Caption for the problem statement",
                                           slide_body = "Body Text",
                                           layout_name = "Problem Statement",
                                           master_name = 'EVA - Standard') {
  stopifnot('Layout is not a problem statement slide' = grepl("Problem", layout_name))
  stopifnot("Layout doesn't exist" = layout_name %in% layout_summary(x)$layout)
  officer::add_slide(x = x,
                     layout = layout_name,
                     master = master_name) %>%
    officer::ph_with(value = slide_problem_statement,
                     location = officer::ph_location_label(ph_label = "Title")) %>%
    officer::ph_with(
      value = slide_subtitle,
      location = officer::ph_location_label(ph_label = "Body Subtitle")
    ) %>%
    {
      `if`(
        !is.null(slide_body),
        ph_with(
          x = .,
          value = slide_body,
          location = officer::ph_location_label(ph_label = "Main Body Text")
        ),
        .

      )
    } %>%
    ph_disclaimer(x = .) %>%
    ph_slide_number(x = .) %>%
    return(.)
}
