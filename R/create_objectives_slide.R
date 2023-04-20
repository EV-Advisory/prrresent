#' Create Objectives Slide
#'
#' This function creates a slide in a PowerPoint template containing a table about the objectives of the project. It adds a slide with the selected layout and adds a title, table, caption, body text and section to the slide.
#'
#' @param x The PowerPoint document to add the slide to.
#' @param slide_title The title to be added to the slide. Default is "Default Title for OKRs".
#' @param slide_subtitle The subtitle to be added to the slide. Default is "".
#' @param slide_table The table to be added to the slide.
#' @param slide_section The section to be added to the slide.
#' @param footer_text The footer text of the slide
#' @param layout_name The name of the slide layout to be used. Default is "Objectives".
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
#' create_objectives_slide(x,
#' "Table Title",
#' "Table Caption",
#' my_table,
#' "Some text",
#' "Section A",
#' "Objectives",
#' "EVA - Standard",
#' F)
#' }

create_objectives_slide <- function(x,
                                    slide_title = " ",
                                    slide_subtitle = " ",
                                    slide_table,
                                    slide_section = "",
                                    footer_text = "",
                                    layout_name = "Objectives",
                                    master_name = 'EVA - Standard',
                                    autofit = F) {
  #Local variable declaration
  slide_layout <- layout_name

  #Layout checks
  stopifnot('Layout is not an objectives slide' = grepl("Objectives", layout_name))
  stopifnot("Layout doesn't exist" = layout_name %in% layout_summary(x)$layout)

  #Find all of the placeholders on the slide
  placeholders <-
    layout_properties(x, layout = layout_name)$ph_label

  #Determine the title placeholder and the corresponding index
  title_type <-
    `if`("Table with Title" %in% layout_name, "Title", "Body Title")
  title_idx <- grepl(title_type, placeholders)

  #For the corresponding table type, determine the dimensions of the placeholder
  table_line_info <-
    layout_properties(x, layout = layout_name)[grepl("Table", placeholders), ]

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

  #Add the corresponding table slide
  officer::add_slide(x, layout = slide_layout, master = master_name) %>%
    officer::ph_with(
      x = .,
      value = slide_title,
      location = officer::ph_location_label(ph_label = "OKRs Title")
    ) %>%
    #Add the value of the table to the table content placeholder
    officer::ph_with(value = slide_table,
                     location = officer::ph_location_label(ph_label = "Table")) %>%
    #Anonymous function to determine if the caption exists
    {
      `if`(
        "Vision and Aim Subtitle Text" %in% placeholders,
        officer::ph_with(
          x = .,
          value = slide_subtitle,
          location = officer::ph_location_label(ph_label = "Vision and Aim Subtitle Text")
        ),
        .
      )
    } %>%
    #Adding the value for the corresponding section
    officer::ph_with(
      value = slide_section,
      location = officer::ph_location_label(ph_label = "Section Title")
    ) %>%
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
