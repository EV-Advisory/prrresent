#' Create Title Slide
#'
#' This function creates a title slide in a PowerPoint template. It adds
#' a slide with the selected layout and adds a title, subtitle and picture
#' to the slide.
#'
#' @param x The PowerPoint document to add the slide to.
#' @param slide_title The title to be added to the slide.
#' @param slide_subtitle The subtitle to be added to the slide.
#' @param footer_text The footer text of the slide.
#' @param picture The path to the picture to be added to the slide.
#' @param layout_name The name of the slide layout to be used. Default is "Title - Black".
#' @param master_name The name of the slide master to be used. Default is "EVA - Standard".
#'
#' @return A PowerPoint slide with the selected layout and added title, subtitle and picture.
#' @export
#'
#' @examples
#' \dontrun{
#' create_title_slide(x, "Title", "Subtitle", "picture.png", "Title - Blue", "EVA - Standard")
#' }
create_title_slide <- function(x,
                               slide_title,
                               slide_subtitle,
                               picture,
                               footer_text = "",
                               layout_name = "Title - Black",
                               master_name = 'EVA - Standard') {
  #Load the layout into a local object
  slide_layout_properties <- layout_properties(x, layout = layout_name)
  #Check: Stop if the title slide doesn't exist
  stopifnot('Layout is not a title slide' = grepl("Title", layout_name))
  #Check: Stop if the layout doesn't exist
  stopifnot("Layout doesn't exist" = layout_name %in% layout_summary(x)$layout)
  #Create the complete slide
  officer::add_slide(x = x,
                     layout = layout_name,
                     master = master_name) %>%
    #Add the title
    officer::ph_with(value = slide_title,
                     location = officer::ph_location_label(ph_label = "Title")) %>%
    #Add the subtitle
    officer::ph_with(value = slide_subtitle,
                     location = officer::ph_location_label(ph_label = "Subtitle")) %>%
    #Add the picture, if the placeholder exists called "Image Location"
    {
      `if`(
        !is.null(picture),
        ph_with(
          x = .,
          value = officer::external_img(picture),
          location = officer::ph_location_label(ph_label = "Image Location")
        ),
        .

      )
    } %>% {
      #Add the disclaimer, if it exists
      `if`("Disclaimer" %in% slide_layout_properties$ph_label,
           ph_disclaimer(x = .),
           .)
    }%>% {
      #Add the slide number, if it exists
      `if`("Slide Number" %in% slide_layout_properties$ph_label,
           ph_slide_number(x = .),
           .)
    }%>% {
      #Add the footer, if it exists
      `if`("Footer" %in% slide_layout_properties$ph_label,
           officer::ph_with(x=.,
                            value = footer_text,
                            location = officer::ph_location_label(ph_label = "Footer")),
           .)
    } %>%
    return(.)
}
