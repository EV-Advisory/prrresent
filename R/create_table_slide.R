#' Create Table Slide
#'
#' This function creates a slide in a PowerPoint template containing a table. It adds a slide with the selected layout and adds a title, table, caption, body text and section to the slide.
#'
#' @param x The PowerPoint document to add the slide to.
#' @param slide_title The title to be added to the slide. Default is "Default Title for Table".
#' @param slide_caption The caption to be added to the slide. Default is "".
#' @param slide_table The table to be added to the slide.
#' @param slide_body The body text to be added to the slide. Default is "Some text".
#' @param slide_section The section to be added to the slide.
#' @param footer_text The footer text of the slide
#' @param layout_name The name of the slide layout to be used. Default is "Table Full".
#' @param master_name The name of the slide master to be used. Default is "EVA - Standard".
#' @param autofit A boolean to specify if the table should be autofitted. Default is F.
#'
#' @return A PowerPoint slide with the selected layout and added title, table, caption, body text and section.
#' @export
#'
#'@importFrom flextable autofit
#'@importFrom flextable flextable
#'@importFrom methods is
#' @examples
#' \dontrun{
#' create_table_slide(x,
#' "Table Title",
#' "Table Caption",
#' my_table,
#' "Some text",
#' "Section A",
#' "Table Full",
#' "EVA - Standard",
#' F)
#' }

create_table_slide <- function(x,
                               slide_title = "Default Title for Table",
                               slide_caption = "",
                               slide_table,
                               slide_body = "Some text",
                               slide_section = "",
                               footer_text = "",
                               layout_name = "Table Full",
                               master_name = 'EVA - Standard',
                               autofit = F) {
  #Local variable declaration
  slide_layout <- layout_name

  #Layout checks
  stopifnot('Layout is not a table slide' = grepl("Table ", layout_name))
  stopifnot("Layout doesn't exist" = layout_name %in% layout_summary(x)$layout)

  #Find all of the placeholders on the slide
  placeholders <-layout_properties(x, layout = layout_name)$ph_label

  #Determine the title placeholder and the corresponding index
  title_type <-`if`("Table with Title" %in% layout_name, "Title", "Body Title")
  title_idx <- grepl(title_type, placeholders)

  #For the corresponding table type, determine the dimensions of the placeholder
  table_line_info <-layout_properties(x, layout = layout_name)[grepl("Table", placeholders),]

  #Depending on the table type, set the dimensions of the table
  table_width <-
    `if`(
      layout_name == "Table Full",
      (table_line_info$cx + table_line_info$offx) / 15,
      (table_line_info$cx + table_line_info$offx) / 15
    )
  table_height <-
    `if`(
      layout_name == "Table Full",
      (table_line_info$cy - table_line_info$offy) / 10,
      (table_line_info$cy - table_line_info$offy) / 8
    )

  #Determine if the table is or is not a flextable
  if (!is(slide_table, "flextable")) {
    warning("Not a flextable, converting to autofit flextable")
    slide_table <- flextable::flextable(
      slide_table,
      cwidth = table_width,
      cheight = table_height,
      defaults = F
    )
  }
  # if (layout_name == "Table Full")
  #
  #   slide_table <-
  #   flextable(
  #     as.data.frame(slide_table),
  #     cwidth = (table_line_info$cx + table_line_info$offx),
  #     cheight = (table_line_info$cy - table_line_info$offy)
  #   )
  # if (autofit == T)
  #   slide_table <- flextable::autofit(slide_table)


  #Add the corresponding table slide
  officer::add_slide(x, layout = slide_layout, master = master_name) %>%
    #Anonymous function to determine if the title exists
    {
      `if`(
        sum(title_idx) > 0,
        officer::ph_with(
          x = .,
          value = slide_title,
          location = officer::ph_location_label(ph_label = placeholders[title_idx])
        ),
        .
      )
    } %>%
    #Add the value of the table to the table content placeholder
    officer::ph_with(value = slide_table,
                     location = officer::ph_location_label(ph_label = "Table")) %>%
    #Anonymous function to determine if the caption exists
    {
      `if`(
        "Caption" %in% placeholders,
        officer::ph_with(
          x = .,
          value = slide_caption,
          location = officer::ph_location_label(ph_label = "Caption")
        ),
        .
      )
    } %>%
    #Anonymous function to determine if the Body Text field exists
    {
      `if`(
        "Body Text" %in% placeholders,
        officer::ph_with(
          x = .,
          value = slide_body,
          location = officer::ph_location_label(ph_label = "Body Text")
        ),
        .
      )
    }  %>%
    #Adding the value for the corresponding section
    officer::ph_with(value = slide_section,
                     location = officer::ph_location_label(ph_label = "Section")) %>%
    #Adding the footer text to the presentation
    officer::ph_with(
      value = footer_text,
      location = officer::ph_location_label(ph_label = "Footer Text")
    ) %>%
    #Disclaimer of the presentation
    ph_disclaimer(x = .) %>%
    #Placeholder addition for the automatic slide number
    ph_slide_number(x = .)

}
