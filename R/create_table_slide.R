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
  # browser()
  slide_layout <- layout_name
  stopifnot('Layout is not a table slide' = grepl("Table ", layout_name))
  stopifnot("Layout doesn't exist" = layout_name %in% layout_summary(x)$layout)

  placeholders <-
    layout_properties(x, layout = layout_name)$ph_label
  title_type <-
    `if`("Table with Title" %in% layout_name, "Title", "Body Title")
  title_idx <- grepl(title_type, placeholders)

  table_line_info <-
    layout_properties(x, layout = layout_name)[grepl("Table", placeholders), ]
  table_width<-`if`(layout_name == "Table Full",(table_line_info$cx + table_line_info$offx)/15,(table_line_info$cx + table_line_info$offx)/20)
  table_height<-`if`(layout_name == "Table Full",(table_line_info$cy - table_line_info$offy)/10,(table_line_info$cy - table_line_info$offy)/10)
  print(table_line_info)
  if (is(slide_table,"flextable")) {
    warning("Not a flextable, converting to autofit flextable")
    slide_table <- flextable::flextable(slide_table,
                                        cwidth = table_width,
                                        cheight =table_height,defaults = F)
    # slide_table<-flextable::flextable_dim(slide_table,"cm")
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

  officer::add_slide(x, layout = slide_layout, master = master_name) %>%
    {
      `if`(
        sum(title_idx) > 0,
        officer::ph_with(x = .,
          value = slide_title,
          location = officer::ph_location_label(ph_label = placeholders[title_idx])
        ),
        .
      )
    } %>%
    officer::ph_with(value = slide_table,
                     location = officer::ph_location_label(ph_label = "Table"))%>%
    {
      `if`(
        "Caption"%in%placeholders,
        officer::ph_with(x = .,
                         value = slide_caption,
                         location = officer::ph_location_label(ph_label = "Caption")
        ),
        .
      )
    }%>%
    {
      `if`(
        "Body Text"%in%placeholders,
        officer::ph_with(x = .,
                         value = slide_body,
                         location = officer::ph_location_label(ph_label = "Body Text")
        ),
        .
      )
    }  %>%
    officer::ph_with(value = slide_section,
                     location = officer::ph_location_label(ph_label = "Section"))%>%
    officer::ph_with(value = footer_text,
                     location = officer::ph_location_label(ph_label = "Footer Text"))%>%
    ph_disclaimer(x=.)%>%
    ph_slide_number(x=.)

}
