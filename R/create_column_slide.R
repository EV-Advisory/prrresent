#' Create Column Slide
#'
#' This function creates a slide in a PowerPoint template containing a dynamic number of columns.
#' It adds a slide with the selected layout and adds a title, table, caption,
#' body text and section to the slide.
#'
#' @param x The PowerPoint document to add the slide to.
#' @param slide_title The title to be added to the slide. Default is "Default Title for Columns".
#' @param slide_caption The caption to be added to the slide. Default is "".
#' @param column_data The data used to fill in the corresponding columns a list including title + body
#' @param slide_section The section to be added to the slide.
#' @param footer_text The footer text of the slide
#' @param layout_name The name of the slide layout to be used. Default is "Two Column".
#' @param master_name The name of the slide master to be used. Default is "EVA - Standard".
#' @importFrom officer layout_properties
#' @importFrom officer layout_summary
#' @return A PowerPoint slide with the selected layout and added title, table, caption, body text and section.
#' @export
#'
#' @examples
#' \dontrun{
#' doc%>%create_column_slide(
#' x = .,
#' column_data = list(
#' list('title' = "Title of Column 1", 'body' = "Body of Column 1"),
#' list('title' = "Title of Column 2", 'body' = "Body of Column 2"),
#' list('title' = "Title of Column 3", 'body' = "Body of Column 3"),
#' list('title' = "Title of Column 4", 'body' = "Body of Column 4")
#' ),
#' layout_name = "Four Column"
#' )
#' }

create_column_slide <- function(x,
                                slide_title = "Default Title for Columns",
                                slide_caption = "",
                                # slide_graphic,
                                column_data = NULL,
                                slide_section = "",
                                footer_text = "",
                                layout_name = "Two Column",
                                master_name = 'EVA - Standard') {
  # browser()
  slide_layout <- layout_name
  stopifnot('Layout is not a column slide' = grepl("Column", layout_name))
  stopifnot("Layout doesn't exist" = layout_name %in% layout_summary(x)$layout)
  stopifnot("Column Data Missing" = !is.null(column_data))


  #Determine the number of columns from the layout name
  number_of_columns <-
    unlist(strsplit(layout_name, "[[:blank:]]"))[1]
  # Convert the string to a numeric value
  ncols <- switch(
    number_of_columns,
    "Two" = 2,
    "Three" = 3,
    "Four" = 4,
    "Five" = 5
  )
  #Determine whether or not the column data matches the selected layout
  #To do: dynamically detect the number of columns necessary to avoid this check
  stopifnot("Column Data does not match layout template" = length(column_data) ==
              ncols)

  #Select the slide template
  officer::add_slide(x, layout = slide_layout, master = master_name) %>%
    {
      #Determine if a main title exists
      `if`(
        !grepl("No Title", layout_name),
        officer::ph_with(
          x = .,
          value = slide_title,
          location = officer::ph_location_label(ph_label = "Title")
        ),
        .
      )
    }  %>% (function(xx) {
      #For the given columns, run a for loop of the given information to populate the theme
      for (i in 1:ncols) {
        xx <- xx %>% officer::ph_with(
          x = .,
          value = column_data[[i]][['title']],
          location = officer::ph_location_label(ph_label = paste0("Body Title - ", i))
        ) %>% officer::ph_with(
          x = .,
          value = column_data[[i]][['body']],
          location = officer::ph_location_label(ph_label = paste0("Body - ", i))
        )
      }
      return(xx)
    }) %>%
    #Set the section of the corresponding slide deck
    officer::ph_with(value = slide_section,
                     location = officer::ph_location_label(ph_label = "Section")) %>%
    #Populate the footer text
    officer::ph_with(
      value = footer_text,
      location = officer::ph_location_label(ph_label = "Footer Text")
    ) %>%
    #Update the disclaimer, which is the default, typically
    ph_disclaimer(x = .) %>%
    #Add the slide number dynamically
    ph_slide_number(x = .)

}
